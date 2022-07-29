import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Writer {
    static String[] cols = {"END ITEM", "ASSEMBLY", "COMPONENT", "COMP NOMENCLATURE",
                    "PART TYPE", "QUANTITY", "BOM LEVEL", "CAGE CODE",
                    "BOM SEQUENCE NUMBER"};
  public static void main(String[] args) throws IOException {
      Scanner input = new Scanner(System.in);
      System.out.println("Enter filepath for BOM (ex. D:\\BOMs\\SomeBOM.xlsx):");
      String filepath = input.nextLine();
      System.out.println("Enter output directory (ex. C:\\Documents\\BOMOutput):");
      String outDir = input.nextLine();
      System.out.println("Enter effectivity range to generate files for.");
      System.out.println("Minimum effectivity (ie. 1):");
      int minEff = input.nextInt();
      System.out.println("Maximum effectivity:");
      int maxEff = input.nextInt();
      System.out.println("Generating files for " + filepath + " from effectivities " + minEff + "-" + maxEff + "...");
      input.close();
      for (int eff = minEff; eff <= maxEff; eff++) {
          System.out.println("Processing effectivity " + eff + "...");
          int cellnum = 0;
          XSSFWorkbook workbook = new XSSFWorkbook();
          XSSFSheet sheet = workbook.createSheet("Sheet1");
          FileInputStream file = new FileInputStream(new File(filepath));
          XSSFWorkbook oldBook = new XSSFWorkbook(file);
          XSSFSheet oldSheet = oldBook.getSheetAt(0);

          {
            Row firstRow = sheet.createRow(0);
            for (int i = 0; i < 9; i++) {
              firstRow.createCell(i);
            }

            Iterator<Cell> cellIterator = firstRow.iterator();
            while (cellIterator.hasNext()) {
              Cell cell = cellIterator.next();
              cell.setCellValue(cols[cellnum]);
              cellnum++;
            }
          }
          for (int i = 1; i <= oldBook.getSheetAt(0).getLastRowNum();i++) {
              Row row = sheet.createRow(i);
              for (int j = 0; j < 9; j++) {
                  Cell cell = row.createCell(j);
                  cell.setCellValue("");
              }
          }

          try {
              int low = -1, high = -1;
              boolean isBetween = false;
              int usableCell = 49;
              int installedCell = 48;
              Iterator<Row> rowIterator = oldSheet.iterator();
              Row row = rowIterator.next(); // set row to index 1
              Pattern pattern1 = Pattern.compile(".*([(].[)])?E\\s((\\d{1,4})-(\\d{1,4})).*"); //identifies if relevant
            Pattern pattern2 =
              Pattern.compile(
                  "(?:E\\h+|\\G(?!^),?((\\d{1,4})-(\\d{1,4})))"); // identifies if multiple ranges
              int newRowNum = 1;
              String data = "NULL";
              while (rowIterator.hasNext()) {
                  row = rowIterator.next();

                  if (row.getCell(usableCell) != null && (row.getCell(installedCell) != null)){
                      data = row.getCell(usableCell).getStringCellValue();
                  } else if ((row.getCell(usableCell) == null) && (row.getCell(installedCell) != null)) {
                      data = row.getCell(installedCell).getStringCellValue();
                  } else if ((row.getCell(installedCell) == null) && (row.getCell(usableCell) != null)){
                      data = row.getCell(usableCell).getStringCellValue();
                  } else if ((row.getCell(usableCell) == null) && (row.getCell(installedCell) == null)) {
                      continue;
                  }

                  // Now, data is stored as the value of the effectivity we may want to put into the new
                  // sheet.
                  Matcher matcher1 = pattern1.matcher(data);
                  Matcher matcher2;
                  if (matcher1.matches()) {
                      matcher2 = pattern2.matcher(data);
                      while(matcher2.find()) {
                          // extract lows/highs here
                          if (matcher2.group(2) != null) {
                              if (eff >= Integer.parseInt(matcher2.group(2)) && eff <= Integer.parseInt(matcher2.group(3))) {
                                  isBetween = true;
                              }
                          }

                      }

                      if (isBetween) {
                          //Match found! Write row into the sheet at index newRowNum
                          Row newRow = sheet.getRow(newRowNum);
                          Cell firstCell = newRow.getCell(0);
                          firstCell.setCellValue("FA-18E");
                          for (int i = 1; i < 9; i++) {
                              Cell newCell = newRow.getCell(i);
                              switch (i) {
                                  case 1:
                                      if (row.getCell(40) != null) {
                                          newCell.setCellValue(row.getCell(40).getStringCellValue());
                                      }
                                      break;
                                  case 2:
                                      if (row.getCell(2) != null) {
                                          newCell.setCellValue(row.getCell(2).getStringCellValue());
                                      }
                                      break;
                                  case 3:
                                      if (row.getCell(29) != null) {
                                          newCell.setCellValue(row.getCell(29).getStringCellValue());
                                      }
                                      break;
                                  case 4:
                                      newCell.setCellValue("-");
                                      break;
                                  case 5:
                                      if (row.getCell(30) != null) {
                                          newCell.setCellValue(row.getCell(30).getStringCellValue());
                                      }
                                      break;
                                  case 6:
                                      break;
                                  case 7:
                                      if (row.getCell(0) != null) {
                                          newCell.setCellValue(row.getCell(0).getStringCellValue());
                                      }
                                      break;
                                  case 8:
                                      newCell.setCellValue(newRowNum++);
                                      break;
                              }
                          }

                      }
                      isBetween = false;
                  }
              }
          } catch (Exception e) {
              e.printStackTrace();
          }

          FileOutputStream out = new FileOutputStream(outDir + "\\" + eff + ".xlsx");
          workbook.write(out);
          out.close();
          System.out.println("Done effectivity " + eff + "!");
      }

  }
}