package com.gentech.excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class WriteExcelFileData {
    public static void main(String[] args) {
        writeContent();
    }
    private static void writeContent()
    {
        FileOutputStream fout=null;
        Workbook wb=null;
        Sheet sh=null;
        Row row=null;
        Cell cell=null;
        try
        {
            wb=new XSSFWorkbook();
            sh=wb.createSheet("Credentials");
            //1st row
            row=sh.createRow(0);
            cell=row.createCell(0);
            cell.setCellValue("UserName");
            cell=row.createCell(1);
            cell.setCellValue("Password");
            //2nd row
            row=sh.createRow(1);
            cell=row.createCell(0);
            cell.setCellValue("admin");
            cell=row.createCell(1);
            cell.setCellValue("manager");
            fout=new FileOutputStream("D:\\Excel\\Welcome.xlsx");
            wb.write(fout);
        }catch(Exception e)
        {
            e.printStackTrace();
        }
        finally {
            try
            {
                fout.close();
                wb.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

}
