package com.gentech.excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Program6 {
    public static void main(String[] args) {
        try {
            readWriteContent();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void readWriteContent() throws IOException {
        FileInputStream fin = null;
        FileOutputStream fout = null;
        Workbook wb = null;
        Sheet sh1 = null;
        Sheet sh2 = null;
        Row rowSh1 = null;
        Row rowSh2 = null;
        Cell cellSh1 = null;
        Cell cellSh2 = null;

        try {
            fin = new FileInputStream("D:\\Excel\\Flowers.xlsx");
            wb = new XSSFWorkbook(fin);
            sh1 = wb.getSheetAt(0);
            sh2 = wb.getSheet("Sheet2");

            if (sh2 == null) {
                sh2 = wb.createSheet("Sheet2");
            }
            int rc = sh1.getPhysicalNumberOfRows();

            int rowIndexInSheet2 = 4;

            for (int r = 0; r < rc; r++) {
                rowSh1 = sh1.getRow(r);
                if (rowSh1 != null) {
                    cellSh1 = rowSh1.getCell(0);

                    if (cellSh1 != null) {
                        String fruitName = cellSh1.getStringCellValue();

                        rowSh2 = sh2.getRow(rowIndexInSheet2);

                        if (rowSh2 == null) {
                            rowSh2 = sh2.createRow(rowIndexInSheet2);
                        }

                        cellSh2 = rowSh2.createCell(0);
                        cellSh2.setCellValue(fruitName);

                        rowIndexInSheet2++;
                    }
                }
            }
            fout = new FileOutputStream("D:\\Excel\\Demo1.xlsx");
            wb.write(fout);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fin != null) fin.close();
                if (fout != null) fout.close();
                if (wb != null) wb.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
