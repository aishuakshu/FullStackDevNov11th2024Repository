package com.gentech.excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
public class Program2 {
    public static void main(String[] args) {
        writeContent();
    }

    private static void writeContent() {
        FileOutputStream fout = null;
        Workbook wb = null;
        Sheet sh = null;
        Row row = null;
        Cell cell = null;

        String[] fruits = {
                "Apple", "Banana", "Cherry", "Date", "Grape", "Kiwi", "Lemon", "Mango", "Nectarine", "Orange",
                "Papaya", "Peach", "Pear", "Pineapple", "Plum", "Pomegranate", "Raspberry", "Strawberry", "Watermelon", "Blueberry"
        };

        try {
            wb = new XSSFWorkbook();
            sh = wb.createSheet("Sheet1");
            row = sh.createRow(0);

            for (int i = 0; i < fruits.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(fruits[i]);
            }
            fout = new FileOutputStream("D:\\Excel\\Fruits.xlsx");
            wb.write(fout);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fout != null) fout.close();
                if (wb != null) wb.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
