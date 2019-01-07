package com;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.util.Iterator;
import java.util.WeakHashMap;

/**
 * Author:   YYX
 * Date:     2019/1/7
 */
public class Demo01 {
    public static void main(String[] args) throws Exception {
        Workbook excel1997 = new HSSFWorkbook(); // excel 1997
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        excel1997.write(fileOut);
        fileOut.close();

        Workbook excel2007 = new XSSFWorkbook(); // excel 2007
        fileOut = new FileOutputStream("workbook.xlsx");
        excel2007.write(fileOut);
        fileOut.close();

        Workbook wb = new HSSFWorkbook();
    }

    @Test
    public void demo02() throws Exception {
        Workbook wb = new XSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow((short) 0);
        // Create a cell and put a value in it.
        Cell cell = row.createCell(0);
        cell.setCellValue(1);

        // Or do it on one line.
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(
                createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);

        // Write the output to a file
//        OutputStream fileOut = new FileOutputStream("workbook04.xls");
        OutputStream fileOut = new FileOutputStream("workbook04.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    @Test
    public void demo03() {
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream("workbook04.xlsx");
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (row == null) {
                    continue;
                }
                for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                    Cell cell = row.getCell(i);
                    String cellValue = "";
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            cellValue = cell.getRichStringCellValue().getString();
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                cellValue = cell.getDateCellValue().toString();
                            } else {
                                cellValue = String.valueOf(cell.getNumericCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            cellValue = String.valueOf(cell.getCellFormula());
                            break;
                        case Cell.CELL_TYPE_BLANK:
                            break;
                        default:
                    }
                    System.out.println("CellNum"+i+"=>CellValue:"+cellValue);
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if (inputStream != null){
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
