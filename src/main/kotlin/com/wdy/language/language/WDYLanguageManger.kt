//package com.wdy.language.language
//
////import com.intellij.openapi.actionSystem.AnAction
////import com.intellij.openapi.actionSystem.AnActionEvent
////import com.intellij.openapi.fileChooser.FileChooser
////import com.intellij.openapi.fileChooser.FileChooserDescriptorFactory
////import com.intellij.openapi.ui.Messages
////import com.intellij.openapi.vfs.readText
////import com.intellij.openapi.vfs.writeText
////import org.apache.poi.ss.usermodel.WorkbookFactory
////import java.io.FileInputStream
////import com.intellij.openapi.actionSystem.AnAction
////import com.intellij.openapi.actionSystem.AnActionEvent
////import com.intellij.openapi.fileChooser.FileChooser
////import com.intellij.openapi.fileChooser.FileChooserDescriptorFactory
////import com.intellij.openapi.project.Project
////import com.intellij.openapi.ui.Messages
//import com.intellij.openapi.actionSystem.AnAction
//import com.intellij.openapi.actionSystem.AnActionEvent
//import com.intellij.openapi.fileChooser.FileChooserDescriptorFactory
//import com.intellij.openapi.fileEditor.FileDocumentManager
//import com.intellij.openapi.project.Project
//import com.intellij.openapi.ui.Messages
//import com.intellij.openapi.vfs.VirtualFile
//import com.intellij.psi.PsiDirectory
//import com.intellij.psi.PsiFile
//import com.intellij.psi.PsiManager
//import com.intellij.psi.search.FilenameIndex
//import com.intellij.psi.search.GlobalSearchScope
//import org.apache.poi.ss.usermodel.CellType
//import org.apache.poi.xssf.usermodel.XSSFWorkbook
//import org.jdom.Document
//import org.jdom.Element
//import org.jdom.input.SAXBuilder
//import org.jdom.output.Format
//import org.jdom.output.XMLOutputter
//import java.io.FileInputStream
//import java.io.StringReader
//import javax.swing.JFileChooser
//
//class WDYLanguageManger : AnAction() {
//    override fun actionPerformed(event: AnActionEvent) {
////        val project = event.project ?: return
////
////        // Select Excel file using a file chooser dialog
////        val excelFile = FileChooser.chooseFile(FileChooserDescriptorFactory.createSingleFileNoJarsDescriptor(), project, null)
////        if (excelFile == null || !excelFile.exists()) {
////            Messages.showMessageDialog(project, "Selected Excel file does not exist.", "Error", Messages.getErrorIcon())
////            return
////        }
////
////        try {
////            FileInputStream(excelFile.path).use { inputStream ->
////                val workbook = WorkbookFactory.create(inputStream)
////                val sheet = workbook.getSheetAt(0)
////
////                val translations = mutableListOf<Map<String, String>>()
////
////                for (row in sheet) {
////                    val chineseTextCell = row.getCell(3) // 假设中文在第四列
////                    val englishTextCell = row.getCell(5) // 假设英文在第六列
////                    val russianTextCell = row.getCell(6) // 假设俄语在第七列
////                    val arabicTextCell = row.getCell(7) // 假设阿拉伯语在第八列
////                    val spanishTextCell = row.getCell(8) // 假设西班牙语在第九列
////                    val portugueseTextCell = row.getCell(9) // 假设葡萄牙语在第十列
////                    val frenchTextCell = row.getCell(10) // 假设法语在第十一列
////
////                    if (chineseTextCell != null && englishTextCell != null &&
////                        russianTextCell != null && arabicTextCell != null &&
////                        spanishTextCell != null && portugueseTextCell != null &&
////                        frenchTextCell != null) {
////
////                        val translation = mapOf(
////                            "Chinese" to chineseTextCell.toString().trim(),
////                            "English" to englishTextCell.toString().trim(),
////                            "Russian" to russianTextCell.toString().trim(),
////                            "Arabic" to arabicTextCell.toString().trim(),
////                            "Spanish" to spanishTextCell.toString().trim(),
////                            "Portuguese" to portugueseTextCell.toString().trim(),
////                            "French" to frenchTextCell.toString().trim()
////                        )
////                        translations.add(translation)
////                    }
////                }
////
////                val message = translations.joinToString("\n") { translation ->
////                    "Chinese: ${translation["Chinese"]}\n" +
////                            "English: ${translation["English"]}\n" +
////                            "Russian: ${translation["Russian"]}\n" +
////                            "Arabic: ${translation["Arabic"]}\n" +
////                            "Spanish: ${translation["Spanish"]}\n" +
////                            "Portuguese: ${translation["Portuguese"]}\n" +
////                            "French: ${translation["French"]}\n"
////                }
////
////                Messages.showMessageDialog(
////                    project,
////                    message,
////                    "Translations",
////                    Messages.getInformationIcon()
////                )
////            }
////        } catch (e: Exception) {
////            Messages.showErrorDialog(
////                project,
////                "Error reading Excel file: ${e.message}",
////                "Error"
////            )
////        }
//       init(event)
//    }
//    private fun init(event: AnActionEvent) {
//        val project = event.project ?: return
//
//        // Show file chooser to select the Excel file
//        val fileChooser = JFileChooser()
//        fileChooser.dialogTitle = "Choose an Excel File"
//        fileChooser.fileFilter = javax.swing.filechooser.FileNameExtensionFilter("Excel Files", "xlsx")
//
//        if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
//            val excelFile = fileChooser.selectedFile
//            val translationsMap = readExcelFile(excelFile)
//
//            if (translationsMap.isEmpty()) {
//                Messages.showMessageDialog(project, "No valid data found in the Excel file.", "Error", Messages.getErrorIcon())
//                return
//            }
//
//            // Find all values directories in the project
//            val valuesDirectories = findValuesDirectories(project)
//
//            if (valuesDirectories.isEmpty()) {
//                Messages.showMessageDialog(project, "No 'values' directory found in the project.", "Error", Messages.getErrorIcon())
//                return
//            }
//
//            var stringsFiles = mutableListOf<PsiFile>()
//
//            for (directory in valuesDirectories) {
//                val virtualFiles = FilenameIndex.getVirtualFilesByName("strings.xml", GlobalSearchScope.directoryScope(directory))
//                stringsFiles.addAll(virtualFiles.mapNotNull { PsiManager.getInstance(project).findFile(it) })
//            }
//
//            if (stringsFiles.isEmpty()) {
//                Messages.showMessageDialog(project, "No strings.xml file found in the 'values' directory.", "Error", Messages.getErrorIcon())
//                return
//            }
//
//            val results = mutableMapOf<String, MutableMap<String, String>>()
//
//            for (stringsFile in stringsFiles) {
//                val content = stringsFile.text
//                val builder = SAXBuilder()
//                val document: Document? = try {
//                    builder.build(StringReader(content))
//                } catch (e: Exception) {
//                    Messages.showMessageDialog("Error parsing strings.xml: ${e.message}", "Error", Messages.getErrorIcon())
//                    continue
//                }
//
//                val rootElement = document?.rootElement
//                if (rootElement == null || rootElement.name != "resources") {
//                    Messages.showMessageDialog("Invalid strings.xml format.", "Error", Messages.getErrorIcon())
//                    continue
//                }
//
//                val stringElements = rootElement.getChildren("string")
//                for (element in stringElements) {
//                    val name = element.getAttributeValue("name")
//                    val zhText = element.text
//
//                    if (zhText in translationsMap) {
//                        val translations = translationsMap[zhText]!!
//
//                        for ((languageCode, translation) in translations) {
//                            val targetDir = getOrCreateTargetDirectory(project, languageCode)
//                            val targetStringsFile = getOrCreateStringsFile(targetDir)
//
//                            updateOrAddString(targetStringsFile, name!!, translation)
//                            results.putIfAbsent(languageCode, mutableMapOf())
//                            results[languageCode]!![name!!] = translation
//                        }
//                    }
//                }
//            }
//
//            showResultsTable(results)
//        }
//    }
//
//    private fun readExcelFile(file: VirtualFile): Map<String, Map<String, String>> {
//        val translationsMap = mutableMapOf<String, Map<String, String>>()
//
//        FileInputStream(file.path).use { inputStream ->
//            XSSFWorkbook(inputStream).use { workbook ->
//                val sheet = workbook.getSheetAt(0)
//                for (rowNum in 1 until sheet.lastRowNum + 1) {
//                    val row = sheet.getRow(rowNum) ?: continue
//                    val chineseCell = row.getCell(3) ?: continue // D column
//                    if (chineseCell.cellType == CellType.STRING) {
//                        val chineseText = chineseCell.stringCellValue.trim()
//                        val englishText = row.getCell(4)?.stringCellValue?.trim() ?: ""
//                        val russianText = row.getCell(5)?.stringCellValue?.trim() ?: ""
//                        val arabicText = row.getCell(6)?.stringCellValue?.trim() ?: ""
//                        val spanishText = row.getCell(7)?.stringCellValue?.trim() ?: ""
//                        val portugueseText = row.getCell(8)?.stringCellValue?.trim() ?: ""
//                        val frenchText = row.getCell(9)?.stringCellValue?.trim() ?: ""
//
//                        translationsMap[chineseText] = mapOf(
//                            "en" to englishText,
//                            "ru" to russianText,
//                            "ar" to arabicText,
//                            "es" to spanishText,
//                            "pt" to portugueseText,
//                            "fr" to frenchText
//                        )
//                    }
//                }
//            }
//        }
//
//        return translationsMap
//    }
//
//    private fun findValuesDirectories(project: Project): List<PsiDirectory> {
//        val resourceDirs = FilenameIndex.getAllFilesByExt(project, "xml", GlobalSearchScope.allScope(project))
//            .filter { it.parent?.name == "res" }
//            .mapNotNull { it.parent }
//
//        return resourceDirs.flatMap { dir ->
//            dir.subdirectories.filter { it.name.startsWith("values") }
//        }
//    }
//
//    private fun getOrCreateTargetDirectory(project: Project, languageCode: String): PsiDirectory {
//        val resDir = project.baseDir.findChild("src/main/res")
//            ?: throw IllegalArgumentException("res directory not found")
//
//        val valuesDirName = when (languageCode) {
//            "en" -> "values-en"
//            "ru" -> "values-ru"
//            "ar" -> "values-ar"
//            "es" -> "values-es"
//            "pt" -> "values-pt"
//            "fr" -> "values-fr"
//            else -> throw IllegalArgumentException("Unsupported language code: $languageCode")
//        }
//
//        var targetDir = resDir.findSubdirectory(valuesDirName)
//        if (targetDir == null) {
//            targetDir = resDir.createSubdirectory(valuesDirName)
//        }
//
//        return targetDir
//    }
//
//    private fun getOrCreateStringsFile(directory: PsiDirectory): PsiFile {
//        val existingFile = directory.findFile("strings.xml")
//        if (existingFile != null) {
//            return existingFile
//        }
//
//        val newFile = directory.createFile("strings.xml")
//        val initialContent = """<?xml version="1.0" encoding="utf-8"?>
//<resources>
//</resources>"""
//
//        FileDocumentManager.getInstance().getDocument(newFile.virtualFile)?.setText(initialContent)
//        return newFile
//    }
//    private fun updateOrAddString(stringsFile: PsiFile, name: String, value: String) {
//        val content = stringsFile.text
//        val builder = SAXBuilder()
//        val document: Document? = try {
//            builder.build(StringReader(content))
//        } catch (e: Exception) {
//            throw RuntimeException("Error parsing strings.xml", e)
//        }
//
//        val rootElement = document?.rootElement ?: throw IllegalStateException("Invalid strings.xml format")
//
//        val existingStringElement = rootElement.getChild("string", rootElement.namespace)?.children
//            ?.firstOrNull { it.getAttributeValue("name") == name }
//
//        if (existingStringElement != null) {
//            existingStringElement.setText(value)
//        } else {
//            val newStringElement = Element("string").setAttribute("name", name).setText(value)
//            rootElement.addContent(newStringElement)
//        }
//
//        val outputFormat = Format.getPrettyFormat().apply {
//            encoding = "UTF-8"
//            indent = "    "
//        }
//        val xmlOutputter = XMLOutputter(outputFormat)
//        val updatedContent = xmlOutputter.outputString(document)
//
//        FileDocumentManager.getInstance().getDocument(stringsFile.virtualFile)?.setText(updatedContent)
//    }
//
//    private fun showResultsTable(results: Map<String, Map<String, String>>) {
//        val columnNames = arrayOf("Language Code", "Key", "Translation")
//        val data = mutableListOf<Array<String>>()
//
//        for ((languageCode, translations) in results) {
//            for ((key, translation) in translations) {
//                data.add(arrayOf(languageCode, key, translation))
//            }
//        }
//
//        ResultsTableDialog(data.toTypedArray(), columnNames).show()
//    }
//}