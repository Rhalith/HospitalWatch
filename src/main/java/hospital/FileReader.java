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

    private final String adress;
    private final String excelName;
    private final String sheetName;
    private final String newExcelName;
    private final int days;
    private int columnNum;
    private boolean isPreviousClear = true;
    private boolean isNextClear = true;

    public FileReader(String adress, String excelName, String sheetName, int days) {
        this.adress = adress;
        this.excelName = excelName;
        this.sheetName = sheetName;
        this.newExcelName = excelName+" yeni";
        this.days = days;
    }
    public FileReader(String adress, String excelName, String sheetName, int days, String newExcelName) {
        this.adress = adress;
        this.excelName = excelName;
        this.sheetName = sheetName;
        this.newExcelName = newExcelName;
        this.days = days;
    }

    List<Cell> availableCells = new ArrayList<>();
    List<Integer> unavailableDoctors = new ArrayList<>();
    List<Integer> offDoctors = new ArrayList<>();
    List<Integer> watchingDoctors = new ArrayList<>();
    public void readFile()throws IOException {
        CheckDays(days);
        //obtaining input bytes from a file
        File file = new File(adress+"\\"+excelName+".xlsx");   //creating a new file instance
        FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
        //creating Workbook instance that refers to .xlsx file
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet(sheetName);     //creating a Sheet object to retrieve object
        FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
        checkFile(formulaEvaluator, sheet, wb);
        fis.close();
        FileOutputStream os = new FileOutputStream(adress+"\\"+newExcelName+".xlsx");
        wb.write(os);
        //Close the workbook and output stream
        wb.close();
        os.close();

    }

    private void checkFile(FormulaEvaluator formulaEvaluator, XSSFSheet sheet, XSSFWorkbook wb){
        for (int j = 3; j < columnNum; j++) {
            for (int i = 14; i < 30; i++) {
                Cell cell = sheet.getRow(i).getCell(j);
                Cell previousCell = null;
                Cell nextCell = null;
                if(j != 3) {
                    previousCell = sheet.getRow(i).getCell(j-1);
                }
                if(j != columnNum){
                    nextCell = sheet.getRow(i).getCell(j+1);
                }
                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.STRING)) {
                    checkCellString(cell, i);
                } else if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
                    System.out.println("Numara olmaması gereken '" + cell.getAddress() + "' hücresinde numara var. Lütfen excel dosyasını düzelt.");
                } else {
                    if(previousCell != null && previousCell.getCellTypeEnum().equals(CellType.STRING)){
                        if (previousCell.getStringCellValue().equals("N")) {
                            isPreviousClear = false;
                        }
                    }
                    if(nextCell != null && nextCell.getCellTypeEnum().equals(CellType.STRING)){
                        if(nextCell.getStringCellValue().equals("N")){
                            isNextClear = false;
                        }
                    }
                    if(isPreviousClear && isNextClear) availableCells.add(cell);
                    isPreviousClear = true; isNextClear = true;
                }
            }
            fillCells(availableCells, wb);
            resetLists(availableCells, offDoctors, watchingDoctors);
        }
    }

    private void CheckDays(int days){
        columnNum = days+2;
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

    public int getDays() {
        return days;
    }
}

