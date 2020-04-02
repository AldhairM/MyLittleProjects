package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MyClass {
    private static final String FILE_NAME = "C:/Users/test.admin/Documents/MyFirstExcel.xlsx";
    
    static List<String> data = new LinkedList<String>();
    
    public static void main(String[] args) {
//    public static void main(String[] args) {
        try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator iterator = datatypeSheet.iterator();

            while(iterator.hasNext()) {
                Row currentRow = (Row)iterator.next();
                Iterator cellIterator = currentRow.iterator();

                while(cellIterator.hasNext()) {
                    Cell currentCell = (Cell)cellIterator.next();
                    if (currentCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    	double data0= currentCell.getNumericCellValue();
                        data.add(""+data0);
                    }else if(currentCell.getCellType()==Cell.CELL_TYPE_STRING){
                    	String data0= currentCell.getStringCellValue();
                        data.add(data0);
                    }
                    
         
                }
            }
        } catch (FileNotFoundException var8) {
            var8.printStackTrace();
        } catch (IOException var9) {
            var9.printStackTrace();
        }
        
        for(String i: data) {
        	System.out.println(i);
        }

    }
}