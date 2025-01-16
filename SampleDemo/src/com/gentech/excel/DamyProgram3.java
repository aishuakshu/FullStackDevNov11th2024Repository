package com.gentech.excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class DamyProgram3 {

        public static void main(String[] args) {
            readWriteContent();
        }

        private static void readWriteContent() {
            FileInputStream fin = null;
            FileOutputStream fout = null;
            Workbook wb = null;
            Sheet sh1 = null;
            Sheet sh3 = null;
            Row rowsh1 = null;
            Row rowsh3 = null;
            Cell cellSh1 = null;
            Cell cellsh3 = null;

            try {
                fin = new FileInputStream("D:\\Excel\\Book1.xlsx");
                wb = new XSSFWorkbook(fin);

                sh1 = wb.getSheet("Sheet1");
                if (sh1 == null) {
                    throw new RuntimeException("Sheet1 does not exist in the file.");
                }

                sh3 = wb.getSheet("Sheet3");
                if (sh3 == null) {
                    sh3 = wb.createSheet("Sheet3");
                }

                int rc = sh1.getPhysicalNumberOfRows();
                for (int r = 0; r < rc; r++) {
                    rowsh1 = sh1.getRow(r);
                    rowsh3 = sh3.getRow(r);
                    if (rowsh3 == null) {
                        rowsh3 = sh3.createRow(r);
                    }
                    int cc = rowsh1.getPhysicalNumberOfCells();
                    for (int c = 0; c < cc; c++) {
                        cellSh1 = rowsh1.getCell(c);
                        String data = cellSh1.getStringCellValue();
                        cellsh3 = rowsh3.getCell(c);
                        if (cellsh3 == null) {
                            cellsh3 = rowsh3.createCell(c);
                        }
                        cellsh3.setCellValue(data);
                    }
                }

                fout = new FileOutputStream("D:\\Excel\\Demo.xlsx");
                wb.write(fout);

            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                try {
                    if (fin != null) {
                        fin.close();
                    }
                    if (fout != null) {
                        fout.close();
                    }
                    if (wb != null) {
                        wb.close();
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }


