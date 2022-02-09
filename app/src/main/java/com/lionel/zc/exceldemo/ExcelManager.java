package com.lionel.zc.exceldemo;

import android.content.Context;
import android.util.Log;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Random;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.RGB;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.WritableWorkbookImpl;

/**
 * author: Lionel
 * date: 2022-02-09 22:24
 */
class ExcelManager {
    String dirPath;
    private static String filePrefix = "excelFile_";
    private static String fixFileName = "excelFile_0.xls";
    private Executor executor = Executors.newSingleThreadExecutor();

    public ExcelManager(Context context) {
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
//      String fileName = filePrefix + index + ".xls";.

        //使用固定的文件名... 测试  修改 excel 表内容. .
        String fileName = fixFileName;

        WritableWorkbook wwb = null;
        Workbook in = null;
        File workbookFile = new File(dirPath + "/" + fileName);
        try {
            if (workbookFile.exists()) {
                //如果文件已经存在...那么将in  加入到  Writable 中.
                in = Workbook.getWorkbook(workbookFile);
            }
        } catch (IOException | BiffException e) {
            e.printStackTrace();
        }
        try {
            //如果 原有的文件不存在或者导入失败, 创建新的. 否则,将原有的文件导入..执行新的修改..
            wwb = in == null ? Workbook.createWorkbook(workbookFile) : Workbook.createWorkbook(workbookFile, in);
        } catch (IOException e) {
            e.printStackTrace();
        }

        if (wwb == null) return;

        //创建一个单元表sheet, 如果已经存在直接获取, 如果不存在, 创建一个新的 sheet 表.
        WritableSheet sheet;
        WritableSheet[] sheets = wwb.getSheets();
        if (sheets == null || sheets.length == 0) {
            sheet = wwb.createSheet("sheet1", 0);
        } else {
            sheet = sheets[0];
        }

        //写入新的数据....
        write(sheet, new Random().nextInt(15), new Random().nextInt(15), "hello word!!", Colour.GREEN);
        write(sheet, new Random().nextInt(15), new Random().nextInt(15), "hello word!!", Colour.GREEN);
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
        Label cell2 = new Label(colunm, row, content, cf);
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
