package com.wdy.language.language

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File

class StringsTranslatorService {

    fun parseExcel(file: File): Map<String, Map<String, String>> {
        val translations = mutableMapOf<String, Map<String, String>>()
        val formatter = DataFormatter() // 用于安全地格式化单元格值

        try {
            WorkbookFactory.create(file).use { workbook ->
                val sheet = workbook.getSheetAt(0)

                for (row in sheet) {
                    if (row.rowNum == 0) continue // 跳过表头

                    val key = row.getCell(0)?.stringCellValue ?: continue
                    val translationsForKey = mutableMapOf(
                        "ar" to getCellValue(row.getCell(4), formatter),
                        "es" to getCellValue(row.getCell(5), formatter),
                        "pt" to getCellValue(row.getCell(6), formatter),
                        "fr" to getCellValue(row.getCell(7), formatter)
                    )

                    translations[key] = translationsForKey
                }
            }
        } catch (e: Exception) {
            throw RuntimeException("Failed to parse Excel file: ${e.message}", e)
        }

        return translations
    }

    private fun getCellValue(cell: Cell?, formatter: DataFormatter): String {
        if (cell == null) return ""
        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> formatter.formatCellValue(cell) // 格式化数值为字符串
            CellType.BOOLEAN -> cell.booleanCellValue.toString()
            else -> ""
        }
    }
}