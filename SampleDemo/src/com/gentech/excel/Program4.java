package com.gentech.excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
public class Program4 {
        public static void main(String[] args) {
            writeContent();
        }

        private static void writeContent() {
            FileOutputStream fout = null;
            Workbook wb = null;
            Sheet sh = null;
            Row row = null;
            Cell cell = null;
            String[] countries = {
                    "USA", "Canada", "Brazil", "Argentina", "Germany", "France", "Italy", "Spain", "Japan", "Australia",
                    "India", "China", "Russia", "South Korea", "Mexico", "South Africa", "Nigeria", "Egypt", "Saudi Arabia", "Turkey"
            };
            try {
                wb = new XSSFWorkbook();
                sh = wb.createSheet("Sheet1");

                for (int i = 0; i < countries.length; i++) {
                    row = sh.createRow(i);
                    cell = row.createCell(i);
                    cell.setCellValue(countries[i]);
                }
                fout = new FileOutputStream("D:\\Excel\\CountriesDiagonally.xlsx");
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
