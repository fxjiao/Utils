package com.example.excel;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;

/**
 * Created by fxjiao on 16/10/28.
 */

public class ExcelUtil {

    /**
     * 获取可读的workbook
     *
     * @param filePath  excel文件路径
     * @return
     */
    public static Workbook getReadableWorkBook(String filePath){
        try {
            Workbook    workbook = Workbook.getWorkbook(new File(filePath));
            return workbook;
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取可写的workbook
     *
     * @param filePath  excel文件路径
     * @return
     */
    public static WritableWorkbook getWritableWorkbook(String filePath){
        WritableWorkbook    workbook  = null;
        try {
            workbook = Workbook.createWorkbook(new File(filePath));
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

}
