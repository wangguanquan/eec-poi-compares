/*
 * Copyright (c) 2019, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ttzero.compares.index.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

/**
 * Create by guanquan.wang at 2019-09-26 17:17
 */
public class PoiService {
    private File myFile;
    private XSSFWorkbook myWorkBook;
    private XSSFSheet mySheet;

    public PoiService() throws IOException {
        myFile = new File(getClass().getClassLoader().getResource(".").getFile(), "Employee.xlsx");
        FileInputStream fis = new FileInputStream(myFile);
        myWorkBook = new XSSFWorkbook (fis);
        // Finds the workbook instance for XLSX file
        myWorkBook = new XSSFWorkbook (fis);
    }

    public void reader() throws IOException {

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();

        // Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // For each row, iterate through each columns
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {

                Cell cell = cellIterator.next();

//                switch (cell.getCellType()) {
//                    case STRING:
//                        System.out.print(cell.getStringCellValue() + "\t");
//                        break;
//                    case NUMERIC:
//                        System.out.print(cell.getNumericCellValue() + "\t");
//                        break;
//                    case BOOLEAN:
//                        System.out.print(cell.getBooleanCellValue() + "\t");
//                        break;
//                    default :
//
//                }
            }
            System.out.println();
        }
    }

    public void writer() throws IOException {
        Map<String, Object[]> data = new HashMap<>();
        data.put("7", new Object[] {7d, "Sonya", "75K", "SALES", "Rupert"});
        data.put("8", new Object[] {8d, "Kris", "85K", "SALES", "Rupert"});
        data.put("9", new Object[] {9d, "Dave", "90K", "SALES", "Rupert"});

        // Set to Iterate and add rows into XLS file
        Set<String> newRows = data.keySet();

        // get the last row number to append new data
        int rownum = mySheet.getLastRowNum();

        for (String key : newRows) {

            // Creating a new Row in existing XLSX sheet
            Row row = mySheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Boolean) {
                    cell.setCellValue((Boolean) obj);
                } else if (obj instanceof Date) {
                    cell.setCellValue((Date) obj);
                } else if (obj instanceof Double) {
                    cell.setCellValue((Double) obj);
                }
            }
        }

        // open an OutputStream to save written data into XLSX file
        FileOutputStream os = new FileOutputStream(myFile);
        myWorkBook.write(os);
        System.out.println("Writing on XLSX file Finished ...");
    }

    public static void main(String[] args) throws IOException {
        new PoiService().writer();
    }
}
