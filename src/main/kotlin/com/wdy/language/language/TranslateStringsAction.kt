package com.wdy.language.language

import com.intellij.openapi.actionSystem.AnAction
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.fileChooser.FileChooserDescriptor
import com.intellij.openapi.fileChooser.FileChooserFactory
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.Messages
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.w3c.dom.Element
import org.w3c.dom.Node
import java.awt.BorderLayout
import java.io.File
import java.nio.file.Files
import java.nio.file.Paths
import javax.swing.*
import javax.swing.table.DefaultTableModel
import javax.xml.transform.OutputKeys
import javax.xml.transform.TransformerFactory
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult
import org.apache.poi.xssf.usermodel.XSSFWorkbook  // Apache POI 的 Excel 2007+ 工作簿
import java.io.FileOutputStream

class TranslateStringsAction : AnAction() {
    override fun actionPerformed(event: AnActionEvent) {
        val project = event.project
        if (project == null) {
            Messages.showErrorDialog("工程不正确，请检查工程是否有strings", "错误")
            return
        }

        val selectedFile = chooseExcelFile(project) ?: return

        try {
            parseExcelFile(selectedFile) { excelData ->
                if (excelData.isEmpty()) {
                    Messages.showErrorDialog("在所选的 Excel 文件中未找到有效数据，请检查表格。", "错误")
                    return@parseExcelFile
                }

                val stringsData = readStringsFromProject(project)
                val mergedData = mergeExcelWithStrings(excelData, stringsData)

                SwingUtilities.invokeLater {
                    showEditableTableWithSaveButton(
                        project,
                        mergedData,
                        "编辑 Strings 表格"
                    ) { updatedData ->
                        saveStringsToProject(project, updatedData, stringsData)
                    }
                }
            }
        } catch (e: Exception) {
            Messages.showErrorDialog("处理 Excel 文件时出错: ${e.message}", "错误")
            e.printStackTrace()
        }
    }
}

private fun chooseExcelFile(project: Project): File? {
    val descriptor = FileChooserDescriptor(true, false, false, false, false, false)
    descriptor.withFileFilter { it.name.endsWith(".xlsx") || it.name.endsWith(".xls") }
    val virtualFile =
        FileChooserFactory.getInstance().createFileChooser(descriptor, project, null).choose(project).firstOrNull()
    return virtualFile?.canonicalPath?.let { File(it) }
}

// -------------------- Excel 解析 --------------------
private fun parseExcelFile(file: File, callback: (Map<String, List<String>>) -> Unit) {
    val result = mutableMapOf<String, MutableList<String>>() // key=中文, value=翻译行

    val loadingDialog = JDialog().apply {
        title = "加载中..."
        setSize(300, 100)
        setLocationRelativeTo(null)
        isModal = false
        layout = BorderLayout()
        add(JLabel("正在读取 Excel 文件，请稍候..."), BorderLayout.CENTER)
    }
    loadingDialog.isVisible = true

    val worker = object : SwingWorker<Map<String, List<String>>, Void>() {
        override fun doInBackground(): Map<String, List<String>> {
            try {
                WorkbookFactory.create(file).use { workbook ->
                    for (sheetIndex in 0 until workbook.numberOfSheets) {
                        val sheet = workbook.getSheetAt(sheetIndex)
                        if (sheet.physicalNumberOfRows == 0) continue

                        val headerRow: Row? = sheet.getRow(0)?.takeIf { it.physicalNumberOfCells > 0 }
                            ?: sheet.getRow(1)?.takeIf { it.physicalNumberOfCells > 0 }
                        if (headerRow == null) continue

                        val headerRowIndex = headerRow.rowNum
                        val titleIndexMap = mutableMapOf<String, Int>()
                        for (cell in headerRow) {
                            val colIndex = cell.columnIndex
                            // 列被隐藏不读取
                            if (sheet.isColumnHidden(colIndex)) continue

                            val title = cell.toString()
                            when {
                                title.contains("文言") -> titleIndexMap["zh"] = cell.columnIndex
                                title.contains("英文") -> titleIndexMap["en"] = cell.columnIndex
                                title.contains("俄文") -> titleIndexMap["ru"] = cell.columnIndex
                                title.contains("阿拉伯文") -> titleIndexMap["ar"] = cell.columnIndex
                                title.contains("西班牙文") -> titleIndexMap["es"] = cell.columnIndex
                                title.contains("葡萄牙文") -> titleIndexMap["pt"] = cell.columnIndex
                                title.contains("法文") -> titleIndexMap["fr"] = cell.columnIndex
                            }
                        }
                        val zhIndex = titleIndexMap["zh"] ?: continue

                        for (rowIndex in (headerRowIndex + 1) until sheet.physicalNumberOfRows) {
                            val row = sheet.getRow(rowIndex) ?: continue
                            val zhCell = row.getCell(zhIndex)
                            val zhText = zhCell?.toString()?.trim()
                            if (zhText.isNullOrEmpty() || isStrikethrough(zhCell)) continue

                            val rowData = result.getOrPut(zhText) { mutableListOf(zhText, "", "", "", "", "", "") }

                            val langOrder = listOf("en", "ru", "ar", "es", "pt", "fr")
                            for ((langIdx, lang) in langOrder.withIndex()) {
                                val colIndex = titleIndexMap[lang]
                                val value = if (colIndex != null) {
                                    val cell = row.getCell(colIndex)
                                    if (isStrikethrough(cell)) "" else cell?.toString()?.trim() ?: ""
                                } else ""

                                // 只有非空值才覆盖
                                if (value.isNotEmpty()) {
                                    rowData[langIdx + 1] = value
                                }
                            }
                            result[zhText] = rowData
                        }
                    }
                }
            } catch (e: Exception) {
                Messages.showErrorDialog("读取 Excel 文件时出错: ${e.message}", "错误")
                e.printStackTrace()
            }
            return result
        }

        override fun done() {
            loadingDialog.dispose()
            try {
                callback(get())
            } catch (e: Exception) {
                Messages.showErrorDialog("处理 Excel 数据失败: ${e.message}", "错误")
                e.printStackTrace()
            }
        }
    }
    worker.execute()
}

// -------------------- 合并逻辑 --------------------
private fun mergeExcelWithStrings(
    excelData: Map<String, List<String>>,
    stringsData: Map<String, Map<String, String>>
): List<List<String>> {
    val merged = mutableListOf<List<String>>()

    val header = listOf("Name", "Chinese", "English", "Russian", "Arabic", "Spanish", "Portuguese", "French")
    merged.add(header)

    for ((name, translations) in stringsData) {
        val mergedRow = mutableListOf<String>()
        mergedRow.add(name)

        val defaultValue = translations["default"] ?: ""
        mergedRow.add(defaultValue)

        val excelRow = excelData[defaultValue]

        mergedRow.add(excelRow?.getOrNull(1) ?: translations["en"] ?: "")
        mergedRow.add(excelRow?.getOrNull(2) ?: translations["ru"] ?: "")
        mergedRow.add(excelRow?.getOrNull(3) ?: translations["ar"] ?: "")
        mergedRow.add(excelRow?.getOrNull(4) ?: translations["es"] ?: "")
        mergedRow.add(excelRow?.getOrNull(5) ?: translations["pt"] ?: "")
        mergedRow.add(excelRow?.getOrNull(6) ?: translations["fr"] ?: "")

        merged.add(mergedRow)
    }
    return merged
}

// -------------------- 其他工具方法 --------------------
private fun readStringsFromProject(project: Project): Map<String, Map<String, String>> {
    val stringsMap = mutableMapOf<String, MutableMap<String, String>>()
    val virtualFiles = Files.walk(Paths.get(project.basePath ?: ""))
        .filter {
            val file = it.toFile()
            file.isFile &&
                    file.extension.equals("xml", ignoreCase = true) &&
                    file.name.contains("string", ignoreCase = true)
        }
        .toList()

    for (filePath in virtualFiles) {
        val file = filePath.toFile()
        try {
            val documentBuilder = javax.xml.parsers.DocumentBuilderFactory.newInstance()
                .newDocumentBuilder()
            val document = documentBuilder.parse(file)
            val root = document.documentElement
            if (root == null || root.nodeName != "resources") continue

            val locale = getLocaleFromPath(filePath.toString())
            val nodes = document.getElementsByTagName("string")
            for (i in 0 until nodes.length) {
                val node = nodes.item(i)
                val name = node.attributes?.getNamedItem("name")?.nodeValue ?: continue
                val value = node.textContent
                stringsMap.computeIfAbsent(name) { mutableMapOf() }[locale] = value
            }
        } catch (_: Exception) {}
    }
    return stringsMap
}

private fun getLocaleFromPath(path: String): String {
    val regex = Regex("values-([a-z]{2})")
    val match = regex.find(path)
    return match?.groupValues?.get(1) ?: "default"
}

private fun isStrikethrough(cell: Cell?): Boolean {
    if (cell == null) return false
    val workbook = cell.sheet.workbook
    val style = cell.cellStyle ?: return false
    val font = workbook.getFontAt(style.fontIndexAsInt)
    return font.strikeout
}

private fun showEditableTableWithSaveButton(
    project: Project,
    data: List<List<String>>,
    title: String,
    saveAction: (List<List<String>>) -> Unit
) {
    val frame = JFrame(title)
    frame.layout = BorderLayout()

    val tableModel = DefaultTableModel()
    for (column in data[0]) {
        tableModel.addColumn(column)
    }
    for (row in data.drop(1)) {
        tableModel.addRow(row.toTypedArray())
    }

    val table = JTable(tableModel)
    val scrollPane = JScrollPane(table)

    val buttonPanel = JPanel() // 按钮区域
    val saveButton = JButton("保存修改")
    val exportButton = JButton("保存为Excel")

    saveButton.addActionListener {
        val updatedData = mutableListOf<List<String>>()
        for (i in 0 until tableModel.rowCount) {
            val row = mutableListOf<String>()
            for (j in 0 until tableModel.columnCount) {
                row.add(tableModel.getValueAt(i, j)?.toString() ?: "")
            }
            updatedData.add(row)
        }
        saveAction(updatedData)
        JOptionPane.showMessageDialog(
            frame,
            "修改内容已保存到 string.xml！",
            "保存成功",
            JOptionPane.INFORMATION_MESSAGE
        )
    }
// 导出为 Excel
    exportButton.addActionListener {
        val chooser = JFileChooser()
        chooser.dialogTitle = "选择保存路径"
        chooser.fileSelectionMode = JFileChooser.FILES_ONLY
        val result = chooser.showSaveDialog(frame)
        if (result == JFileChooser.APPROVE_OPTION) {
            val file = chooser.selectedFile
            try {
                val workbook = XSSFWorkbook()
                val sheet = workbook.createSheet("Translations")


// 写表头
                val headerRow = sheet.createRow(0)
                for (j in 0 until tableModel.columnCount) {
                    headerRow.createCell(j).setCellValue(tableModel.getColumnName(j))
                }


// 写数据
                for (i in 0 until tableModel.rowCount) {
                    val row = sheet.createRow(i + 1)
                    for (j in 0 until tableModel.columnCount) {
                        row.createCell(j).setCellValue(tableModel.getValueAt(i, j)?.toString() ?: "")
                    }
                }


                val targetFile = if (file.name.endsWith(".xlsx")) file else File(file.absolutePath + ".xlsx")
                FileOutputStream(targetFile).use { fos ->
                    workbook.write(fos)
                }
                workbook.close()


                JOptionPane.showMessageDialog(
                    frame,
                    "Excel 文件已保存到: ${targetFile.absolutePath}",
                    "保存成功",
                    JOptionPane.INFORMATION_MESSAGE
                )
            } catch (ex: Exception) {
                JOptionPane.showMessageDialog(
                    frame,
                    "保存 Excel 失败: ${ex.message}",
                    "错误",
                    JOptionPane.ERROR_MESSAGE
                )
                ex.printStackTrace()
            }
        }
    }


    buttonPanel.add(saveButton)
    buttonPanel.add(exportButton)
    frame.add(scrollPane, BorderLayout.CENTER)
    frame.add(buttonPanel, BorderLayout.SOUTH)
    frame.setSize(800, 600)
    frame.isVisible = true
}

private fun escapeXml(input: String): String {
    return input.replace("'", "\\'")
}

private fun saveStringsToProject(
    project: Project,
    updatedData: List<List<String>>,
    originalStringsData: Map<String, Map<String, String>>
) {
    val projectBasePath = project.basePath ?: return
    val stringsFiles = Files.walk(Paths.get(projectBasePath))
        .filter {
            val name = it.toFile().name
            name == "strings.xml" || name == "string.xml"
        }
        .toList()

    for (filePath in stringsFiles) {
        val file = filePath.toFile()
        val locale = getLocaleFromPath(filePath.toString())

        val document = javax.xml.parsers.DocumentBuilderFactory.newInstance()
            .newDocumentBuilder()
            .parse(file)
        removeWhitespaceNodes(document)

        val nodes = document.getElementsByTagName("string")
        val updatedStrings = updatedData.associateBy({ it[0] }, { it })

        for (i in 0 until nodes.length) {
            val node = nodes.item(i)
            if (node.nodeType == Node.ELEMENT_NODE) {
                val element = node as Element
                val name = element.getAttribute("name")
                val updatedRow = updatedStrings[name] ?: continue
                val updatedValue = when (locale) {
                    "zh" -> updatedRow.getOrNull(1) ?: ""
                    "en" -> updatedRow.getOrNull(2) ?: ""
                    "ru" -> updatedRow.getOrNull(3) ?: ""
                    "ar" -> updatedRow.getOrNull(4) ?: ""
                    "es" -> updatedRow.getOrNull(5) ?: ""
                    "pt" -> updatedRow.getOrNull(6) ?: ""
                    "fr" -> escapeXml(updatedRow.getOrNull(7) ?: "")
                    else -> continue
                }
                element.textContent = updatedValue
            }
        }

        val transformer = TransformerFactory.newInstance().newTransformer()
        transformer.setOutputProperty(OutputKeys.INDENT, "yes")
        transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4")
        val source = DOMSource(document)
        val result = StreamResult(file)
        transformer.transform(source, result)
    }
}

private fun removeWhitespaceNodes(node: Node) {
    val children = node.childNodes
    for (i in children.length - 1 downTo 0) {
        val child = children.item(i)
        if (child.nodeType == Node.TEXT_NODE && child.textContent.trim().isEmpty()) {
            node.removeChild(child)
        } else if (child.hasChildNodes()) {
            removeWhitespaceNodes(child)
        }
    }
}
