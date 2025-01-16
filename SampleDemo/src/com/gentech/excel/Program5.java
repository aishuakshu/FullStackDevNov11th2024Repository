package com.gentech.excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
public class Program5 {
    public static void main(String[] args) {
        writeContent();
    }

    private static void writeContent() {
        FileOutputStream fout = null;
        Workbook wb = null;
        Sheet sh = null;
        Row row = null;
        Cell cell = null;
        String[] colors = {
                "Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White",
                "Gray", "Violet", "Indigo", "Cyan", "Magenta", "Turquoise", "Crimson", "Lavender", "Beige", "Gold"
        };

        try {
            wb = new XSSFWorkbook();
            sh = wb.createSheet("Sheet1");
            for (int i = 0; i < colors.length; i++) {
                row = sh.createRow(i);
                cell = row.createCell(4);
                cell.setCellValue(colors[i]);
            }

            fout = new FileOutputStream("D:\\Excel\\Colors.xlsx");
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
