package com.excel.programs;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashSet;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UniqueValuesXLSX {
  public static void initializeSheet(XSSFSheet sheetFrom, XSSFSheet sheetTo) {
    for(int rowIndex = 0; rowIndex <= sheetFrom.getLastRowNum(); rowIndex++) {
      XSSFRow rowFrom = sheetFrom.getRow(rowIndex);
      XSSFRow rowTo = sheetTo.createRow(rowIndex);
      for(int columnIndex = 0; columnIndex < rowFrom.getLastCellNum(); columnIndex++) {
        rowTo.createCell(columnIndex);
      }
    }
  }

  // copies the header of the XLSX sheet
  public static void copyHeader(XSSFSheet sheetFrom, XSSFSheet sheetTo) {
    XSSFRow rowFrom = sheetFrom.getRow(0);
    XSSFCell cellFrom = rowFrom.getCell(0);
    XSSFRow rowTo = sheetTo.getRow(0);

    for(int columnIndex = 0; columnIndex < rowFrom.getLastCellNum(); columnIndex++) {
      cellFrom = rowFrom.getCell(columnIndex);
      XSSFCell cellTo = rowTo.getCell(columnIndex);
      cellTo.setCellValue(cellFrom.toString());
    }
  }

  public static void copyOneColumn(XSSFSheet sheetFrom, XSSFSheet sheetTo, XSSFCell cellFrom, int columnIndex) {
    LinkedHashSet<String> linkedSet = new LinkedHashSet<String>();
    for(int rowIndex = 1; rowIndex <= sheetFrom.getLastRowNum(); rowIndex++) {
      cellFrom = sheetFrom.getRow(rowIndex).getCell(columnIndex);
      if(cellFrom != null) {
        linkedSet.add(cellFrom.toString());
      }
    }

    Iterator<String> it = linkedSet.iterator();
    int rowIndex = 1;
    XSSFRow rowTo = sheetTo.getRow(rowIndex);
    XSSFCell cellTo = rowTo.getCell(columnIndex);
    while(it.hasNext()) {
      rowTo = sheetTo.getRow(rowIndex++);
      cellTo = rowTo.getCell(columnIndex);
      String s = it.next().toString();
      cellTo.setCellValue(s);
    }
  }

  public static void copyAllColumns(XSSFSheet sheetFrom, XSSFSheet sheetTo) {
    XSSFRow rowFrom = sheetFrom.getRow(0);
    XSSFCell cellFrom = rowFrom.getCell(0);

    for(int columnIndex = 0; columnIndex < rowFrom.getLastCellNum(); columnIndex++) {
      copyOneColumn(sheetFrom, sheetTo, cellFrom, columnIndex);
    }
  }

  public static void main(String[] args) throws FileNotFoundException, IOException {

    // ask the user for the file path of the input file
    Scanner in = new Scanner(System.in);
    System.out.println("Enter the file path");
    String filePath = in.nextLine();

    // open the input file we're copying from
    InputStream fileToCopy = new FileInputStream(filePath);
    XSSFWorkbook workbookFrom = new XSSFWorkbook(fileToCopy);
    XSSFSheet sheetFrom = workbookFrom.getSheetAt(0);

    // create a file that we're copying to
    XSSFWorkbook workbookTo = new XSSFWorkbook();
    XSSFSheet sheetTo = workbookTo.createSheet("Sheet 1");

    // initialize the sheet
    initializeSheet(sheetFrom, sheetTo);

    // copy the header
    copyHeader(sheetFrom, sheetTo);

    // then copy all the columns
    copyAllColumns(sheetFrom, sheetTo);

    // finally, write the workbook to an output file
    System.out.println("Enter the destination file path: ");
    String destPath = in.nextLine();
    workbookTo.write(new FileOutputStream(destPath));
    workbookTo.close();
  }
}
