package read_data;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class RetrieveFromFile {

    @Test
    public void readFileTest() throws IOException {

        File excelFile = new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream = new FileInputStream(excelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet page1 = workbook.getSheet("Sheet1");
        XSSFRow row1 = page1.getRow(0);
        XSSFCell cell1 = row1.getCell(0);
        System.out.println(cell1);

        XSSFRow row2 = page1.getRow(1);
        XSSFCell cell2 = row2.getCell(0);
        System.out.println(cell2);
    }

    @Test
    public void getRowValuesTest() throws IOException {

        File file = new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1 = workbook.getSheetAt(0);
        XSSFRow row1 = sheet1.getRow(0);

        for (int i = row1.getFirstCellNum(); i < row1.getLastCellNum(); i++) {
            XSSFCell cell = row1.getCell(i);
            System.out.print(cell + " | ");
        }
    }

    @Test
    public void getAllDataTest() throws IOException {

        //GET all data from excel document
        File file = new File("src/test/resources/TestSetup.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1 = workbook.getSheetAt(0);

        for (int i = sheet1.getFirstRowNum(); i < sheet1.getLastRowNum(); i++) {
            XSSFRow tempRow = sheet1.getRow(i);
            System.out.print("| ");
            for (int j = tempRow.getFirstCellNum(); j < tempRow.getLastCellNum(); j++) {
                System.out.print(tempRow.getCell(j) + " | ");
            }
            System.out.println();
        }


    }


}
