package com.lowman.creatpivotsheets;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

    private File excelFile;

    public void setExcelFile(File excelFile) {
        this.excelFile = excelFile;
    }

    public File getExcelFile() {
        return excelFile;
    }

    public boolean areThereOver25Lines(int sheetNumber) throws IOException {
        //zip bomb error
        ZipSecureFile.setMinInflateRatio(0);

        //link to book
        FileInputStream fileInputStream = new FileInputStream(excelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(sheetNumber);

        //check if sheet has over 20 entries
        int rowNumber = sheet.getLastRowNum();

        workbook.close();
        fileInputStream.close();

        return rowNumber > 20;
    }

    public void createPivotSheet(int sheetNumber) throws IOException{
        //for error about zip bombs
        ZipSecureFile.setMinInflateRatio(0);

        FileInputStream fileInputStream = new FileInputStream(excelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
        String deptName = workbook.getSheetName(sheetNumber);

        //pivot
        try{
            XSSFSheet pivotSheet = workbook.createSheet("PivotSheetFor" + deptName);
            AreaReference areaReference = new AreaReference("B:D", SpreadsheetVersion.EXCEL2007);
            XSSFPivotTable pivotTable = pivotSheet.createPivotTable(areaReference, new CellReference("A1"), sheet);

            pivotTable.addRowLabel(0);
            pivotTable.addRowLabel(1);
            pivotTable.addRowLabel(2);
        }catch(Exception e){
            System.out.println(e);
            return;
        }

        fileInputStream.close();

        FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();
    }

    public int howManySheets() throws IOException {
        ZipSecureFile.setMinInflateRatio(0);

        FileInputStream fileInputStream = new FileInputStream(excelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

         return workbook.getNumberOfSheets();
    }

    public boolean checkIfNormalSheet(int sheetNumber) throws IOException {
        //for error about zip bombs
        ZipSecureFile.setMinInflateRatio(0);

        FileInputStream fileInputStream = new FileInputStream(excelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(sheetNumber);

        Row row = sheet.getRow(0);

        if(row != null){
            Cell cell = row.getCell(0);

            if(cell != null && cell.getStringCellValue().equals("Name")){
                System.out.println("Normal sheet");
                return true;
            }else{
                System.out.println("Not a normal sheet");
                return false;
            }
        }else{
            return false;
        }

    }
}
