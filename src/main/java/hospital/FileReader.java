package hospital;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileReader {


    List<Cell> availableCells = new ArrayList<>();
    List<Integer> unavailableDoctors = new ArrayList<>();
    List<Integer> offDoctors = new ArrayList<>();
    List<Integer> watchingDoctors = new ArrayList<>();
    public void readFile()throws IOException {
        //obtaining input bytes from a file
        File file = new File("C:\\Users\\nuhyi\\OneDrive\\Masaüstü\\Nöbet\\nöbet.xlsx");   //creating a new file instance
        FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
        //creating Workbook instance that refers to .xlsx file
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet("EYLÜL 2023");     //creating a Sheet object to retrieve object
        FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
        for (int j = 3; j < 32; j++) {
            for (int i = 14; i < 30; i++) {
                Cell cell = sheet.getRow(i).getCell(j);
                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.STRING)) {
                    checkCellString(cell, i);
                } else if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                    System.out.println("Numara olmaması gereken '" + cell.getAddress() + "' hücresinde numara var. Lütfen excel dosyasını düzelt.");
                } else {
                    availableCells.add(cell);
                }
            }
            fillCells(availableCells, wb);
            resetLists(availableCells, unavailableDoctors, offDoctors, watchingDoctors);
        }
        fis.close();
        FileOutputStream os = new FileOutputStream("C:\\Users\\nuhyi\\OneDrive\\Masaüstü\\Nöbet\\nöbettest.xlsx");
        wb.write(os);
        //Close the workbook and output stream
        wb.close();
        os.close();

    }

    private void fillCells(List<Cell> cellList, XSSFWorkbook wb)
    {
        Random random = new Random();

        Font font = wb.createFont();
        editFont(font);

        CellStyle cellStyle = wb.createCellStyle();
        editCellStyle(cellStyle, font);
        
        while (4 > watchingDoctors.size()){
            int j = random.nextInt(0,cellList.size());
            if (!watchingDoctors.contains(j))
                if (!unavailableDoctors.contains(j))
                    if (!offDoctors.contains(j)) {
                        cellList.get(j).setCellValue("N");
                        cellList.get(j).setCellStyle(cellStyle);
                        watchingDoctors.add(j);
                        System.out.println(cellList.get(j).getAddress());
                    }
        }
    }

    private void checkCellString(Cell cell, int i)
    {
        if(cell.getStringCellValue().equals("N"))
        {
            watchingDoctors.add(i);
            unavailableDoctors.add(i);
        }
        else if (cell.getStringCellValue().equals("X")){
            offDoctors.add(i);
        }
    }

    
    private void editCellStyle(CellStyle cellStyle, Font font){
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFont(font);
    }

    private void editFont(Font font){
        font.setFontName("Arimo");
        font.setFontHeightInPoints((short) 10);
        font.setBold(true);
    }

    @SafeVarargs
    private void resetLists(List<Cell> cellList, List<Integer>... lists){
        cellList.clear();
        for (List<Integer> list: lists)
        {
            list.clear();
        }
    }
}

