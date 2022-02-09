package com.lionel.zc.exceldemo;

import android.content.Context;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;

import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.RGB;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * author: Lionel
 * date: 2022-02-09 22:24
 */
class ExcelManager {
    Context context;
    String dirPath;
    private static String filePrefix = "excelFile_";
    private Executor executor = Executors.newSingleThreadExecutor();

    public ExcelManager(Context context) {
        this.context = context;
        this.dirPath = context.getFilesDir().getPath() + "/export";
    }

    public void writeExcel() {
        executor.execute(this::export);
    }

    //导出 excel 文件..
    private void export() {
        //检查 文件目录是否存在..
        File dir = new File(dirPath);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        File[] files = dir.listFiles();
        int index = files == null ? 0 : files.length;
        String fileName = filePrefix + String.valueOf(index) + ".xls";

        WritableWorkbook wwb = null;
        OutputStream os = null;
        try {
            os = new FileOutputStream(dirPath + "/" + fileName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        if (os == null) return;
        try {
            wwb = Workbook.createWorkbook(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (wwb == null) return;

        //创建一个单元表sheet
        WritableSheet sheet = wwb.createSheet("sheet1", 0);
        write(sheet, 1, 1, "hello word!!", Colour.GREEN);

        flush(wwb);
    }

    private void write(WritableSheet sheet, int column, int row, String content) {
        write(sheet, column, row, content, null);
    }

    /**
     * @param sheet   当前的 表的 sheet
     * @param colunm  要写入的单元格所在的列
     * @param row     要写入的单元格所在的行
     * @param content 内容
     * @param color   背景颜色..
     */
    private void write(WritableSheet sheet, int colunm, int row, String content, Colour color) {
        WritableCellFormat cf = new WritableCellFormat();
        try {
            cf.setBackground(color);
        } catch (WriteException e) {
            e.printStackTrace();
        }
        Label cell2 = new Label(0, 0, content, cf);
        try {
            sheet.addCell(cell2);
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    //所有 数据编辑完成后,  同意写入. 然后 结束写入流...
    private void flush(WritableWorkbook workbook) {
        try {
            workbook.write();
            workbook.close();
        } catch (IOException | WriteException e) {
            e.printStackTrace();
        }
    }


}
