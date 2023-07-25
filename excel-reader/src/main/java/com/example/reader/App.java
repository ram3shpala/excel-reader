package com.example.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



/**
 * Hello world!
 *
 */
public class App {
    public static void main(String[] args) {
        // #write program to read excel file and print out the content in first column
        // of each row using POI library

        try {   
            //create code to read excel file from resources folder using file stream and pass it to readExcelFile method to read the content of excel file and print it out to console
             



            readExcelFile("myexcel.xlsx");
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        // #pring output program completed without error
        System.out.println("Program completed without error");

    }


    public static void readExcelFile(String fileName) throws IOException {

        ClassLoader classLoader = App.class.getClassLoader();
        InputStream file = classLoader.getResourceAsStream(fileName);
 
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.println(cell.getStringCellValue());
                }
            }
        file.close();
    }
}
