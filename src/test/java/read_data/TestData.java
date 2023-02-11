package read_data;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class TestData {

    File excelFile = new File("src/test/resources/TestData.xlsx");
    FileInputStream fileInputStream;
    XSSFWorkbook workbook;
    XSSFSheet sheet1;

    @Before
    public void setup() throws IOException {
        fileInputStream = new FileInputStream(excelFile);
        workbook = new XSSFWorkbook(fileInputStream);
        sheet1 = workbook.getSheet("Sheet1");
    }

    @Test
    public void getAllDataTest() {

        for (int i = sheet1.getFirstRowNum(); i <= sheet1.getLastRowNum(); i++) {
            XSSFRow tempRow = sheet1.getRow(i);
            System.out.print("| ");
            for (int k = tempRow.getFirstCellNum(); k < tempRow.getLastCellNum(); k++) {
                System.out.print(tempRow.getCell(k) + " | ");
            }
            System.out.println();
        }
    }

    @Test
    public void getSpecificColumnValuesTest() {
        String columnName = "Expire";
        XSSFRow row1 = sheet1.getRow(0);
        int index = -1;
        System.out.println("hello");
        System.out.println("hello");

        for (int i = row1.getFirstCellNum(); i < row1.getLastCellNum(); i++) {
            XSSFCell tempCell = row1.getCell(i);
            if (tempCell.getStringCellValue().equalsIgnoreCase(columnName)) {
                index = i;
            }
        }

        if (index < 0) {
            throw new RuntimeException();
        }

        for (int i = sheet1.getFirstRowNum(); i <= sheet1.getLastRowNum(); i++) {
            XSSFRow tempRow = sheet1.getRow(i);
            System.out.println(tempRow.getCell(index));
        }
    }


}
