package com.gentech.excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
public class Program1 {
        public static void main(String[] args) {
            writeContent();
        }

        private static void writeContent() {
            FileOutputStream fout = null;
            Workbook wb = null;
            Sheet sh = null;
            Row row = null;
            Cell cell = null;
            String[] flowers = {
                    "Rose", "Tulip", "Daisy", "Sunflower", "Lily", "Orchid", "Chrysanthemum", "Carnation", "Marigold", "Lavender",
                    "Jasmine", "Poppy", "Violet", "Peony", "Daffodil", "Bluebell", "Iris", "Begonia", "Freesia", "Geranium"
            };
            try {
                wb = new XSSFWorkbook();
                sh = wb.createSheet("Sheet1");

                for (int i = 0; i < flowers.length; i++) {
                    row = sh.createRow(i);
                    cell = row.createCell(0);
                    cell.setCellValue(flowers[i]);
                }

                fout = new FileOutputStream("D:\\Excel\\Flowers.xlsx");
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

