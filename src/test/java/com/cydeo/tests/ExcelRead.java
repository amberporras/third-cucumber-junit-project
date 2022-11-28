package com.cydeo.tests;

import org.apache.commons.io.filefilter.FileFileFilter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ExcelRead {

    @Test
    public void read_from_excel_file() throws IOException {

        String path= "SampleData.xlsx";

        File file = new File(path);

        //to read from excel we need to load it to FileInput Stream which is coming from java

        FileInputStream fileInputStream = new FileInputStream(file);

        // workbook>sheet>row>cell

        //<1> Create a workbook
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

        //<2> We need to get specific shhet from currently opened workbook
        XSSFSheet sheet = workbook.getSheet("Employees");

        //<3> Select row and cell
        // print out Mary's cell
        // indexs start from 0

        System.out.println(sheet.getRow(1).getCell(0));

        // print out developer
        System.out.println(sheet.getRow(3).getCell(2));

        //get physicalNumberOfRows() <method counts from 1 not 0>
        //example
        //row with data
        //row with data
        //row with empty
        //row with data
        //result would be 3 because 1 row is empty
        int usedRows = sheet.getPhysicalNumberOfRows();
        System.out.println(usedRows);


        //getLastRowNum() method
        //int usedRowsCount = datasheet.getLastRowNum():
        // starts counting from 0, counts empty rows
        //example
        //row with data
        //row with empty data
        // row with empty data
        // row with data
        //It will return 3, since it starts counting from 0

        int lastUsedRow = sheet.getLastRowNum();
        System.out.println(lastUsedRow);
        //ToDo: create a logic to print Vinods name from our excel data
        for (int rowNum=0; rowNum<usedRows; rowNum++){

            if(sheet.getRow(rowNum).getCell(0).toString().equals("Vinod")){
                System.out.println(sheet.getRow(rowNum).getCell(0));
            }
        }

        //TODO: Create a logic to print out Linda's job ID
        // check if name is Linda--> print out job ID of Linda

        for (int rowNum = 0; rowNum<usedRows; rowNum++){
            if(sheet.getRow(rowNum).getCell(0).toString().equals("Linda")){
                System.out.println("Linda's Job ID is: " + sheet.getRow(rowNum).getCell(2));
            }
        }




    }
}
