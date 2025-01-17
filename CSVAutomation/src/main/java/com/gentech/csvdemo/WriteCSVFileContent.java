package com.gentech.csvdemo;

import com.opencsv.CSVWriter;

import java.io.FileWriter;

public class WriteCSVFileContent {
    public static void main(String[] args) {
        writeContent();
    }

    private static void writeContent()
    {
        FileWriter fw=null;
        CSVWriter writer=null;
        try
        {
            fw=new FileWriter("F:\\Demo\\Test\\Employees.csv");
            writer=new CSVWriter(fw);
            String header[]={"EmpNo","Ename","JobName","Salary","DateOfJoin"};
            writer.writeNext(header);
            String row1[]={"101","Santosh","Analyst","12000","12-Jan-2025"};
            writer.writeNext(row1);
            String row2[]={"102","Adams","Clerk","11000","15-Jan-2025"};
            writer.writeNext(row2);
        }catch(Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            try
            {
                writer.flush();
                writer.close();
            }catch(Exception e)
            {
                e.printStackTrace();
            }
        }
    }
}
