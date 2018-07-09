/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package read_write;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.TreeMap;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author S530671
 */
public class Testing_input {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, InvalidFormatException {

        TreeMap<String, ArrayList<String>> tm1 = new TreeMap<>();
        ArrayList<String> values = new ArrayList<>();
        ArrayList<Album> finalList = new ArrayList<Album>();
        File inputFile = new File("Agarwal_Input.xlsx");
        XSSFWorkbook input_workbook = new XSSFWorkbook(inputFile);
        XSSFSheet input_sheet = input_workbook.getSheetAt(0);
        for (int i = 1; i < input_sheet.getPhysicalNumberOfRows(); i++) {
            Row input_row = input_sheet.getRow(i);
            String input[] = {"" + input_row.getCell(4).getNumericCellValue()
            };
            // TODO code application logic here
        }

    }
}
