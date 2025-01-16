package com.gentech.excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
public class Program3 {
    public static void main(String[] args) {
        writeContent();
    }

    private static void writeContent() {
        FileOutputStream fout = null;
        Workbook wb = null;
        Sheet sh = null;
        Row row = null;
        Cell cell = null;

        String[] cities = {
                "Bangalore", "Mysore", "Hubli", "Mangalore", "Belgaum", "Shimoga", "Tumkur", "Davangere", "Bijapur", "Bagalkot",
                "Chitradurga", "Udupi", "Hospet", "Kolar", "Chikmagalur", "Mandya", "Hassan", "Karwar", "Raichur", "Gulbarga"
        };
        try {
            wb = new XSSFWorkbook();
            sh = wb.createSheet("Sheet1");
            row = sh.createRow(9);

            for (int i = 0; i < cities.length; i++)
            {
                cell = row.createCell(i);
                cell.setCellValue(cities[i]);
            }

            fout = new FileOutputStream("D:\\Excel\\KarnatakaCities.xlsx");
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
