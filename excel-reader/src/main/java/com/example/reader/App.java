package com.example.reader;

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
            readExcelFile("myexcel.xlsx");
        } catch (IOException e) {
            System.out.println("Error reading file");
            e.printStackTrace();
        }
    }

    public static void readExcelFile(String fileName) throws IOException {

        ClassLoader classLoader = App.class.getClassLoader();
        InputStream file = classLoader.getResourceAsStream(fileName);

        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();

        // print here output "excel output"
        System.out.println("==========Excel File output Start=========");
        int i = 1;
        // skip first row
        rowIterator.next();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.println(i++ + ". " + cell.getStringCellValue());
            }
        }
        System.out.println("==========Excel File output End=========");
        file.close();
    }
}
