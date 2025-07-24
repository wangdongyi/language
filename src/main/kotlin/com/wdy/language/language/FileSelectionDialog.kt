package com.wdy.language.language

import com.intellij.openapi.project.Project
import com.intellij.openapi.ui.DialogWrapper
import javax.swing.DefaultListModel
import javax.swing.JComponent
import javax.swing.JList
import javax.swing.JScrollPane

class FileSelectionDialog(project: Project, private val fileNames: List<String>) : DialogWrapper(project) {

    private val listModel = DefaultListModel<String>()
    private val fileList = JList(listModel)

    init {
        title = "Select strings.xml File"
        for (fileName in fileNames) {
            listModel.addElement(fileName)
        }
        init()
    }

    override fun createCenterPanel(): JComponent? {
        val scrollPane = JScrollPane(fileList)
        scrollPane.preferredSize = java.awt.Dimension(400, 300) // Set preferred size for better visibility
        return scrollPane
    }

    val selectedFile: String?
        get() = fileList.selectedValue
}