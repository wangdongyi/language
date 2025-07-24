package com.wdy.language.language

import com.intellij.openapi.actionSystem.AnAction
import com.intellij.openapi.actionSystem.AnActionEvent
import com.intellij.openapi.fileChooser.FileChooserDescriptor
import com.intellij.openapi.fileChooser.FileChooserFactory
import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.Messages
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

class TranslateStringsAction : AnAction() {

    override fun actionPerformed(event: AnActionEvent) {
        val project = event.project
        if (project == null) {
            Messages.showErrorDialog("工程不正确，请检查工程是否有strings", "错误")
            return
        }

        // 第一步：用户选择 Excel 文件
        val selectedFile = chooseExcelFile(project) ?: return

        try {
            // 第二步：解析 Excel 文件
            parseExcelFile(selectedFile) { excelData ->
                if (excelData.isEmpty()) {
                    Messages.showErrorDialog("在所选的 Excel 文件中未找到有效数据，请检查表格。", "错误")
                    return@parseExcelFile
                }

                // 第三步：读取 Android 工程中的 strings.xml 文件
                val stringsData = readStringsFromProject(project)

                // 第四步：合并数据
                val mergedData = mergeExcelWithStrings(excelData, stringsData)

                // 第五步：显示表格并允许编辑（回到 UI 线程）
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

private fun parseExcelFile(file: File, callback: (List<List<String>>) -> Unit) {
    val result = mutableListOf<List<String>>()

    val loadingDialog = JDialog().apply {
        title = "加载中..."
        setSize(300, 100)
        setLocationRelativeTo(null)
        isModal = false
        layout = BorderLayout()
        add(JLabel("正在读取 Excel 文件，请稍候..."), BorderLayout.CENTER)
    }
    // 显示 loading 弹窗
    loadingDialog.isVisible = true
    val worker = object : SwingWorker<List<List<String>>, Void>() {
        override fun doInBackground(): List<List<String>> {
            try {
                WorkbookFactory.create(file).use { workbook ->
                    for (sheetIndex in 0 until workbook.numberOfSheets) {
                        val sheet = workbook.getSheetAt(sheetIndex)
                        println("正在读取 Sheet: ${sheet.sheetName}")
                        if (sheet.physicalNumberOfRows == 0) continue
                        // 试着从第 0 行找标题，如果没有再从第 1 行找
                        val headerRow: Row? = sheet.getRow(0)?.takeIf { it.physicalNumberOfCells > 0 }
                            ?: sheet.getRow(1)?.takeIf { it.physicalNumberOfCells > 0 }

                        if (headerRow == null) {
                            println("跳过没有有效标题行的 sheet: ${sheet.sheetName}")
                            continue
                        }
                        // 找到标题行
                        val headerRowIndex = headerRow.rowNum
                        val titleIndexMap = mutableMapOf<String, Int>()
                        for (cell in headerRow) {
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

                        // 添加数据行
                        for (rowIndex in (headerRowIndex + 1) until sheet.physicalNumberOfRows) {
                            val row = sheet.getRow(rowIndex) ?: continue
                            val zhText = row.getCell(zhIndex)?.toString()?.trim()
                            if (zhText.isNullOrEmpty()) continue

                            val rowData = mutableListOf<String>()
                            rowData.add(zhText) // 中文

                            // 添加英文到法文的翻译
                            val langOrder = listOf("en", "ru", "ar", "es", "pt", "fr")
                            for (lang in langOrder) {
                                val colIndex = titleIndexMap[lang]
                                val value = colIndex?.let { row.getCell(it)?.toString()?.trim() } ?: ""
                                rowData.add(value)
                            }

                            result.add(rowData)
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
            if (root == null || root.nodeName != "resources") {
                println("跳过非资源文件：${file.name}")
                continue
            }

            val locale = getLocaleFromPath(filePath.toString())
            val nodes = document.getElementsByTagName("string")
            for (i in 0 until nodes.length) {
                val node = nodes.item(i)
                val name = node.attributes?.getNamedItem("name")?.nodeValue ?: continue
                val value = node.textContent
                stringsMap.computeIfAbsent(name) { mutableMapOf() }[locale] = value
            }

        } catch (e: Exception) {
            println("跳过无法解析的文件：${file.name} (${e.message})")
        }
    }
    return stringsMap
}

private fun getLocaleFromPath(path: String): String {
    val regex = Regex("values-([a-z]{2})")
    val match = regex.find(path)
    return match?.groupValues?.get(1) ?: "default"
}


private fun mergeExcelWithStrings(
    excelData: List<List<String>>,
    stringsData: Map<String, Map<String, String>>
): List<List<String>> {
    val merged = mutableListOf<List<String>>()

    // 添加表头
    val header = listOf("Name", "Chinese", "English", "Russian", "Arabic", "Spanish", "Portuguese", "French")
    merged.add(header)

    // 遍历 stringsData 中的每个 name
    for ((name, translations) in stringsData) {
        val mergedRow = mutableListOf<String>()
        mergedRow.add(name) // 第一列：显示 name

        // 获取 default 语言的 value
        val defaultValue = translations["default"] ?: ""
        mergedRow.add(defaultValue) // 第二列：显示 default 语言的 value

        // 遍历 Excel 数据，查找匹配的中文
        for (row in excelData) {
            if (row.isEmpty()) continue // 跳过空行
            val excelChinese = row[0] // Excel 中的中文
            // 如果 Excel 中的中文与 default 语言的 value 相同
            if (excelChinese == defaultValue) {
                // 将 Excel 中的翻译放入对应的国家翻译列中
                mergedRow.add(row.getOrNull(1) ?: translations["en"] ?: "") // English
                mergedRow.add(row.getOrNull(2) ?: translations["ru"] ?: "") // Russian
                mergedRow.add(row.getOrNull(3) ?: translations["ar"] ?: "") // Arabic
                mergedRow.add(row.getOrNull(4) ?: translations["es"] ?: "") // Spanish
                mergedRow.add(row.getOrNull(5) ?: translations["pt"] ?: "") // Portuguese
                mergedRow.add(row.getOrNull(6) ?: translations["fr"] ?: "") // French
                break // 找到匹配项后跳出循环
            }
        }
        merged.add(mergedRow)
    }
    return merged
}


private fun showExcelDataInTable(data: List<List<String>>, title: String) {
    val frame = JFrame(title)
    val tableModel = DefaultTableModel()
    for (column in data[0]) {
        tableModel.addColumn(column)
    }
    for (row in data.drop(1)) {
        tableModel.addRow(row.toTypedArray())
    }
    val table = JTable(tableModel)
    table.isEnabled = false // Make it non-editable
    frame.add(JScrollPane(table))
    frame.setSize(800, 600)
    frame.isVisible = true
}

private fun showEditableTableWithSaveButton(
    project: Project,
    data: List<List<String>>,
    title: String,
    saveAction: (List<List<String>>) -> Unit
) {
    val frame = JFrame(title)
    frame.layout = BorderLayout()

    // 创建表格模型
    val tableModel = DefaultTableModel()
    for (column in data[0]) {
        tableModel.addColumn(column) // 添加表头
    }
    for (row in data.drop(1)) {
        tableModel.addRow(row.toTypedArray()) // 添加每一行数据
    }

    // 创建表格
    val table = JTable(tableModel)
    val scrollPane = JScrollPane(table)

    // 创建保存按钮
    val saveButton = JButton("保存修改")
    saveButton.addActionListener {
        val updatedData = mutableListOf<List<String>>()
        for (i in 0 until tableModel.rowCount) {
            val row = mutableListOf<String>()
            for (j in 0 until tableModel.columnCount) {
                row.add(tableModel.getValueAt(i, j)?.toString() ?: "")
            }
            updatedData.add(row)
        }
        saveAction(updatedData) // 将修改后的数据传给保存逻辑
        JOptionPane.showMessageDialog(
            frame,
            "修改内容已保存到 string.xml！",
            "保存成功",
            JOptionPane.INFORMATION_MESSAGE
        )
    }

    // 布局
    frame.add(scrollPane, BorderLayout.CENTER)
    frame.add(saveButton, BorderLayout.SOUTH)
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

    // 遍历每个 strings.xml 文件，更新翻译内容
    for (filePath in stringsFiles) {
        val file = filePath.toFile()
        val locale = getLocaleFromPath(filePath.toString())

        val document = javax.xml.parsers.DocumentBuilderFactory.newInstance()
            .newDocumentBuilder()
            .parse(file)

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

        // 保存修改后的内容到文件
        val transformer = javax.xml.transform.TransformerFactory.newInstance().newTransformer()
        transformer.setOutputProperty(javax.xml.transform.OutputKeys.INDENT, "yes")
        val source = javax.xml.transform.dom.DOMSource(document)
        val result = javax.xml.transform.stream.StreamResult(file)
        transformer.transform(source, result)
    }

}




