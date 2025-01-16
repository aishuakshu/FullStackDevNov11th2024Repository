package com.gentech.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Program7Demo {
    public static void main(String[] args) {
        readAndWriteReversedBirdNames();
    }

    private static void readAndWriteReversedBirdNames() {
        FileInputStream fin = null;
        FileOutputStream fout = null;
        Workbook wb = null;
        Sheet sh1 = null;
        Sheet sh2 = null;

        try {
            // Open the Excel file
            fin = new FileInputStream("D:\\Excel\\BirdFile.xlsx");
            wb = new XSSFWorkbook(fin);

            // Access Sheet1 and create Sheet2
            sh1 = wb.getSheet("Sheet1");
            if (sh1 == null) {
                System.out.println("Sheet1 does not exist in the file!");
                return;
            }
            sh2 = wb.getSheet("Sheet2");
            if (sh2 == null) {
                sh2 = wb.createSheet("Sheet2");
            }

            int rc = sh1.getPhysicalNumberOfRows(); // Total number of rows in Sheet1
            int columnIndex = 4; // The 5th column index (0-based)

            // Array to store bird names
            String[] birdNames = new String[rc - 1]; // Exclude header row

            // Read bird names from the 5th column, starting from row 1 (skip header)
            for (int r = 1; r < rc; r++) {
                Row rowsh1 = sh1.getRow(r);
                if (rowsh1 != null) {
                    Cell cell = rowsh1.getCell(columnIndex);
                    if (cell != null) {
                        birdNames[r - 1] = cell.getStringCellValue();
                    } else {
                        birdNames[r - 1] = ""; // Handle empty cells
                    }
                } else {
                    birdNames[r - 1] = ""; // Handle missing rows
                }
            }

            // Write bird names in reverse order to the 5th column of Sheet2
            for (int r = 0; r < birdNames.length; r++) {
                Row rowsh2 = sh2.getRow(r + 1); // Start writing from row 1 (skip header)
                if (rowsh2 == null) {
                    rowsh2 = sh2.createRow(r + 1);
                }
                Cell cellsh2 = rowsh2.getCell(columnIndex);
                if (cellsh2 == null) {
                    cellsh2 = rowsh2.createCell(columnIndex);
                }
                cellsh2.setCellValue(birdNames[birdNames.length - 1 - r]); // Reverse order
            }

            // Write header for Sheet2
            Row headerRow = sh2.getRow(0);
            if (headerRow == null) {
                headerRow = sh2.createRow(0);
            }
            Cell headerCell = headerRow.getCell(columnIndex);
            if (headerCell == null) {
                headerCell = headerRow.createCell(columnIndex);
            }
            headerCell.setCellValue("ReversedBirdNames");

            // Save the output to a new file
            File outputFile = new File("D:\\Excel\\ReversedBirds.xlsx");
            if (outputFile.exists()) {
                outputFile.delete(); // Delete if file exists
            }
            fout = new FileOutputStream(outputFile);
            wb.write(fout);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fin != null) fin.close();
                if (fout != null) fout.close();
                if (wb != null) wb.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
