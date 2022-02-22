package com.example.poi


import org.junit.Assert.*

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.*
import org.junit.Test

import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream


val inputPath = "C:\\Users\\lionel_zc\\Desktop\\a.xls"
val outPath = "C:\\Users\\lionel_zc\\Desktop\\aaaa.xls"


/**
 * Example local unit test, which will execute on the development machine (host).
 *
 * See [testing documentation](http://d.android.com/tools/testing).
 */
class ExampleUnitTest {
    @Test
    fun addition_isCorrect() {

        val workbook = HSSFWorkbook(FileInputStream(File(inputPath)))
        val sheet = workbook.getSheetAt(0);
        val cellStyle = workbook.createCellStyle()
            .apply { fillBackgroundColor = HSSFColor.HSSFColorPredefined.RED.index }

        sheet.forEach { row ->
            row.forEach {
                val c = it.columnIndex
                val r = it.rowIndex
                if (r == 2) {
                    println("value: ${it.getCellValue()}")
                    it.cellStyle = cellStyle
                }
            }
        }
        val outputStream = FileOutputStream(outPath)
        val workbook1 = HSSFWorkbook()
        val sheet1 = workbook1.createSheet(sheet.sheetName)

        sheet.forEach { row ->
            val createRow = sheet1.createRow(row.rowNum)
            row.forEach {
                val c = it.columnIndex
                val cellStyle_ = workbook1.getStyle(IndexedColors.RED.index)
                val createCell = createRow.createCell(c)
                createCell.setCellValue(it.getCellValue())
                createCell.setCellStyle(cellStyle_)
            }
        }

        workbook1.write(outputStream)
        workbook1.close()
    }

    private val styleMap = HashMap<Short, CellStyle>()
    private fun HSSFWorkbook.getStyle(color: Short): CellStyle {
        return styleMap[color].run {
            if (this == null) {
                val cellStyle = createCellStyle()
                cellStyle.setFillForegroundColor(color)
                cellStyle.fillPattern = FillPatternType.SOLID_FOREGROUND
                styleMap[color] = cellStyle
                cellStyle
            } else this
        }
    }

    private fun Cell.getCellValue(): String {
        return when (this.cellType) {
            CellType._NONE -> ""
            CellType.NUMERIC -> this.numericCellValue.toString()
            CellType.STRING -> this.stringCellValue
            CellType.FORMULA -> this.richStringCellValue.toString()
            CellType.BLANK -> ""
            CellType.BOOLEAN -> this.booleanCellValue.toString()
            CellType.ERROR -> this.errorCellValue.toString()
            else -> ""
        }
    }

}