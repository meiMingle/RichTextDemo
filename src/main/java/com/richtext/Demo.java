package com.richtext;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class Demo {
    public static void main(String[] args) {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(new XSSFWorkbook(/* "D://test.xlsx" */), 100, false, true);
             FileOutputStream fos = new FileOutputStream("D://test_.xlsx")
        ) {

            Font font = workbook.createFont();
            font.setColor(IndexedColors.RED.index);
            font.setBold(true);
            font.setItalic(true);
            font.setStrikeout(true);
            // SXSSFSheet sheet0 = workbook.getSheetAt(0);
            SXSSFSheet sheet0 = workbook.createSheet();
            SXSSFRow row0 = sheet0.createRow(2);
            SXSSFCell cell0 = row0.createCell(2);
            WPSXSSFRichTextString rts = new WPSXSSFRichTextString();
            rts.append("模1");
            rts.append("块1", (XSSFFont) font);

            cell0.setCellValue(rts);
            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
