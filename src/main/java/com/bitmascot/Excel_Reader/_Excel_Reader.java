package com.bitmascot.Excel_Reader;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;

public class _Excel_Reader {


    private String path;
    private FileInputStream inputStream;
    public FileOutputStream fileOutputStream;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private XSSFRow row;
    private XSSFCell cell;

    public _Excel_Reader(String path) {
        this.path = path;
        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheetAt(0);
            inputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // returns the row count in a sheet
    public int getRowCount(String sheetName) {
        int index = workbook.getSheetIndex(sheetName);
        if (index == -1)
            return 0;
        else {
            sheet = workbook.getSheetAt(index);
            int number = sheet.getLastRowNum() + 1;
            return number;
        }

    }

    // find whether sheets exists
    public boolean isSheetExist(String sheetName) {
        int index = workbook.getSheetIndex(sheetName);
        if (index == -1) {
            index = workbook.getSheetIndex(sheetName.toUpperCase());
            if (index == -1)
                return false;
            else
                return true;
        } else
            return true;
    }





    // returns true if sheet is created successfully else false
    public boolean addSheet(String sheetname) {

        FileOutputStream fileOut;
        try {
            workbook.createSheet(sheetname);
            fileOut = new FileOutputStream(path);
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    // returns true if sheet is removed successfully else false if sheet does not exist
    public boolean removeSheet(String sheetName) {
        int index = workbook.getSheetIndex(sheetName);
        if (index == -1)
            return false;

        FileOutputStream fileOut;
        try {
            workbook.removeSheetAt(index);
            fileOut = new FileOutputStream(path);
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    // returns true if data is set successfully else false
    public boolean setCellData(String sheetName, int colNum, int rowNum, String data) {

        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);

            if (rowNum <= 0)
                return false;

            int index = workbook.getSheetIndex(sheetName);
            if (index == -1)
                return false;


            sheet = workbook.getSheetAt(index);


            row = sheet.getRow(rowNum);
            if (row == null)
                row = sheet.createRow(rowNum - 1);

            if (colNum < 0)
                return false;

            cell = row.getCell(colNum);

            if (cell == null)
                cell = row.createCell(1);

            cell.setCellValue(data);

            fileOutputStream = new FileOutputStream(path);
            workbook.write(fileOutputStream);

            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }


    public boolean setCellData(String sheetName, int colNum, int rowNum, int data, String message) {

        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);

            if (rowNum <= 0)
                return false;

            int index = workbook.getSheetIndex(sheetName);
            if (index == -1)
                return false;


            sheet = workbook.getSheetAt(index);


            row = sheet.getRow(rowNum);
            if (row == null)
                row = sheet.createRow(rowNum - 1);

            if (colNum < 0)
                return false;

            cell = row.getCell(colNum);

            if (cell == null)
                cell = row.createCell(1);


            cell.setCellValue(data);

            fileOutputStream = new FileOutputStream(path);
            workbook.write(fileOutputStream);

            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    // returns true if column is created successfully
    public boolean addColumn(String sheetName, String colName) {

        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);
            int index = workbook.getSheetIndex(sheetName);
            if (index == -1)
                return false;


            sheet = workbook.getSheetAt(index);

            row = sheet.getRow(0);
            if (row == null)
                row = sheet.createRow(0);


            if (row.getLastCellNum() == -1)
                cell = row.createCell(0);
            else
                cell = row.createCell(row.getLastCellNum());

            cell.setCellValue(colName);

            fileOutputStream = new FileOutputStream(path);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    // removes a column and all the contents
    public boolean removeColumn(String sheetName, int colNum) {
        try {
            if (!isSheetExist(sheetName))
                return false;
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheet(sheetName);
            XSSFCreationHelper createHelper = workbook.getCreationHelper();

            for (int i = 0; i < getRowCount(sheetName); i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    cell = row.getCell(colNum);
                    if (cell != null) {
                        row.removeCell(cell);
                    }
                }
            }
            fileOutputStream = new FileOutputStream(path);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    public int findCloLoc(String sheetName, String columnName) {
        int ret;
        DataFormatter formatter = new DataFormatter();

        int index = workbook.getSheetIndex(sheetName);

        sheet = workbook.getSheetAt(index);

        int rowStart = sheet.getFirstRowNum();
        int rowEnd = sheet.getLastRowNum();

        for (int rows = rowStart; rows < rowEnd; rows++) {
            Row row = sheet.getRow(rows);
            if (row == null) {
                continue;
            }
            int lastColumn = row.getLastCellNum();
            for (int columns = 0; columns < lastColumn; columns++) {
                Cell cell = row.getCell(columns);
                if (cell != null) {
                    String text = formatter.formatCellValue(cell);
                    if (text.contains(columnName)) {
                        return cell.getColumnIndex();

                    }
                }
            }
        }
        return -1;
    }


    public int getRowIndex(String sheetName, String data) throws Throwable {
        inputStream = inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);

        DataFormatter formatter = new DataFormatter();

        int index = workbook.getSheetIndex(sheetName);

        sheet = workbook.getSheetAt(index);

        int rowStart = sheet.getFirstRowNum();
        int rowEnd = sheet.getLastRowNum();

        return getRowIndexByRagne(sheetName, rowStart, rowEnd, data);

    }

    public int getRowIndexByRagne(String sheetName, int startRow, int endRow, String data) throws Throwable {
        inputStream = inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);

        DataFormatter formatter = new DataFormatter();

        int index = workbook.getSheetIndex(sheetName);

        sheet = workbook.getSheetAt(index);

        int rowStart = startRow;
        int rowEnd = endRow;

        for (int rows = rowStart; rows < rowEnd; rows++) {
            Row row = sheet.getRow(rows);
            if (row == null) {
                continue;
            }

            int lastColumn = row.getLastCellNum();

            for (int columns = 0; columns < lastColumn; columns++) {
                Cell cell = row.getCell(columns);
                if (cell != null) {
                    String text = formatter.formatCellValue(cell);
                    if (text.contains(data)) {
                        return cell.getRowIndex();

                    }
                }
            }
        }
        return -1;

    }

    public int getRowIndexByFormulaValue(String sheetName, int startRow, int endRow, String data) throws Throwable {
        Map<String, Integer> allTeamInfo = new HashMap<>();
        inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);

        DataFormatter formatter = new DataFormatter();

        int index = workbook.getSheetIndex(sheetName);

        sheet = workbook.getSheetAt(index);

        int rowStart = startRow;
        int rowEnd = endRow;

        for (int rows = rowStart; rows < rowEnd; rows++) {
            Row row = sheet.getRow(rows);
            if (row == null) {
                continue;
            }

            int lastColumn = row.getLastCellNum();

            for (int columns = 0; columns < lastColumn; columns++) {
                Cell cell = row.getCell(columns);
                if (cell != null && cell.getCellTypeEnum() == CellType.FORMULA) {

                    FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
                    String text = formulaEval.evaluate(cell).formatAsString();
                    if (text.toLowerCase().contains(data.toLowerCase())) {
                        return cell.getRowIndex();
                    }
                }
            }
        }
        return -1;

    }

    public int searchforBlankSpaceIncolunm(String sheetName, int startRow, int endRow, int column) throws Throwable {
        inputStream = inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);

        DataFormatter formatter = new DataFormatter();

        int index = workbook.getSheetIndex(sheetName);

        sheet = workbook.getSheetAt(index);

        int rowStart = startRow;
        int rowEnd = endRow;

        for (int rows = rowStart; rows < rowEnd; rows++) {
            Row row = sheet.getRow(rows);
            if (row == null) {
                continue;
            }
            Cell cell = row.getCell(column);
            if (cell.getCellTypeEnum() == CellType.BLANK) {
                return cell.getRowIndex();
            }
        }
        return -1;

    }


}
