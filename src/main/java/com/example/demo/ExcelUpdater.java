package com.example.demo;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUpdater {

    public static void main(String[] args) throws IOException {
        String filePath = "F:\\1.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(filePath)) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0); // Assuming you want to update the first sheet

            // Example: Update cell value at row 1, column 2
            Row rowToUpdate = sheet.createRow(7);
            Cell cellToUpdate = rowToUpdate.createCell(7);
            
            //rowToUpdate.createCell(5);
            cellToUpdate.setCellValue("Updated dddddddValue");
            

            // Save changes back to the file
            try (FileOutputStream fileOutputStream = new FileOutputStream(filePath)) {
                workbook.write(fileOutputStream);
            }
        }
    }
}

