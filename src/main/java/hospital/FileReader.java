package hospital;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import  org.apache.poi.xssf.usermodel.XSSFColor;

public class FileReader {

    public void ReadFile()throws IOException {
        //obtaining input bytes from a file
        File file = new File("C:\\Users\\nuhyi\\OneDrive\\Masaüstü\\Nöbet\\nöbet.xlsx");   //creating a new file instance
        FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
        //creating Workbook instance that refers to .xlsx file
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet("EYLÜL 2023");     //creating a Sheet object to retrieve object
    //evaluating cell type
        FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
        for(Row row: sheet)     //iteration over row using for each loop
        {
            for(Cell cell: row)    //iteration over cell using for each loop
            {
                switch(cell.getCellType())
                {
                    case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type
                        //getting the value of the cell as a number
                        System.out.print(cell.getNumericCellValue()+ "\t\t");
                        break;
                    case Cell.CELL_TYPE_STRING:    //field that represents string cell type
                        //getting the value of the cell as a string
                        System.out.print(cell.getStringCellValue()+ "\t\t");
                        break;
                }
            }
            System.out.println();
        }
    }
}

