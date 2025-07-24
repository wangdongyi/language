package com.wdy.language.language

import com.intellij.openapi.ui.DialogWrapper
import javax.swing.JComponent
import javax.swing.JScrollPane
import javax.swing.JTable
import javax.swing.table.DefaultTableModel

class StringsDialog(private val data: Array<Array<String>>, private val columnNames: Array<String>) : DialogWrapper(true) {

    init {
        title = "Strings from strings.xml"
        init()
    }

    override fun createCenterPanel(): JComponent? {
        val tableModel = DefaultTableModel(data, columnNames)
        val table = JTable(tableModel)
        table.autoResizeMode = JTable.AUTO_RESIZE_ALL_COLUMNS
        return JScrollPane(table)
    }
}