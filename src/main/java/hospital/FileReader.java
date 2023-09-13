package hospital;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.Timer;
import java.util.concurrent.TimeUnit;

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
    private int zortCount;
    List<Integer> skippingIndexes;
    List<Integer> filledRows = new ArrayList<>();

    public FileReader(String adress, String excelName, String sheetName, int days, List<Integer> skippingIndexes) {
        this.adress = adress;
        this.excelName = excelName;
        this.sheetName = sheetName;
        this.newExcelName = excelName+" yeni";
        this.days = days;
        this.skippingIndexes = skippingIndexes;
    }
    public FileReader(String adress, String excelName, String sheetName, int days, List<Integer> skippingIndexes, String newExcelName) {
        this.adress = adress;
        this.excelName = excelName;
        this.sheetName = sheetName;
        this.newExcelName = newExcelName;
        this.days = days;
        this.skippingIndexes = skippingIndexes;
    }

    List<Cell> filledCells = new ArrayList<>();
    List<Cell> availableCells = new ArrayList<>();
    List<Integer> unavailableDoctors = new ArrayList<>();
    List<Integer> offDoctors = new ArrayList<>();
    List<Integer> watchingDoctors = new ArrayList<>();
    public void readFile()throws IOException {
        CheckDays(days);
        //obtaining input bytes from a file
        File file = new File(adress+"\\"+excelName+".xlsx");   //creating a new file instance
        FileInputStream fis = new FileInputStream(file);
        Random random = new Random();//obtaining bytes from the file
        //creating Workbook instance that refers to .xlsx file
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet(sheetName);     //creating a Sheet object to retrieve object
        FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
        //checkFile(formulaEvaluator, sheet, wb);
        checkAvailableCells(formulaEvaluator, sheet, wb, random);
        fis.close();
        FileOutputStream os = new FileOutputStream(adress+"\\"+newExcelName+".xlsx");
        wb.write(os);
        //Close the workbook and output stream
        wb.close();
        os.close();

    }

    private void checkAvailableCells(FormulaEvaluator formulaEvaluator, XSSFSheet sheet, XSSFWorkbook wb, Random random){
        for (int i = 3; i < days+2; i++)
        {
            int rowNum = random.nextInt(14,30);
            while(skippingIndexes.contains(rowNum) || !checkNCountInRow(rowNum, i, sheet, formulaEvaluator) || filledRows.contains(rowNum))
            {
                rowNum = random.nextInt(14,30);
            }
            int nNum = 4;
            int nCount = 0;
            for (int j = 14; j < 30; j++) {
                Cell cell = sheet.getRow(j).getCell(i);
                if(checkCellHasN(formulaEvaluator, cell)){
                    nCount++;
                }
            }
            while (nCount < nNum){
                Cell previousCell = null;
                Cell nextCell = null;

                if(i != 3) previousCell = sheet.getRow(rowNum).getCell(i-1);
                if(i != days+2) nextCell = sheet.getRow(rowNum).getCell(i+1);

                Cell cell = sheet.getRow(rowNum).getCell(i);

                if(checkCellHasN(formulaEvaluator, cell)){
                    rowNum = random.nextInt(14,30);
                    continue;
                }
                if(checkNeighbourCells(previousCell, nextCell) && checkCellAvailability(formulaEvaluator, cell))
                {
                    fillCell(wb, cell);
                    nCount++;
                }
                else {
                    rowNum = random.nextInt(14,30);
                }
            }
            filledRows.add(rowNum);
            System.out.println(rowNum);
            }
            filledRows.clear();

    }

    private boolean checkNeighbourCells(Cell previousCell, Cell nextCell) {
        if(previousCell != null && previousCell.getCellTypeEnum().equals(CellType.STRING)){
            if (previousCell.getStringCellValue().equals("N")) {
                return false;
            }
        }
        if(nextCell != null && nextCell.getCellTypeEnum().equals(CellType.STRING)){
            if(nextCell.getStringCellValue().equals("N")){
                return false;
            }
        }
        return true;
    }

    private boolean checkCellAvailability(FormulaEvaluator formulaEvaluator, Cell cell){
        if(formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.STRING)){
            return !cell.getStringCellValue().equals("X");
        }
        return true;
    }

    private boolean checkCellHasN(FormulaEvaluator formulaEvaluator, Cell cell){
        if(formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.STRING)){
            return cell.getStringCellValue().equals("N");
        }
        return false;
    }

    private boolean checkNCountInRow(int rowNum, int columnCount, Sheet sheet, FormulaEvaluator formulaEvaluator){
        int nCount = 0;
        for (int i = 3; i < columnCount; i++) {
            Cell cell = sheet.getRow(rowNum).getCell(i);
            if(formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.STRING) && cell.getStringCellValue().equals("N")){
                nCount++;
            }
        }
        return nCount != 4;
    }

    private int countEmptyCells(FormulaEvaluator formulaEvaluator, int columnNum, int rowNum, Sheet sheet){
        int emptyCount = 0;
        Cell previousCell = null;
        Cell nextCell = null;

        if(columnNum != 3) previousCell = sheet.getRow(rowNum).getCell(columnNum-1);
        if(columnNum != 32) nextCell = sheet.getRow(rowNum).getCell(columnNum+1);
        for (int i = 0; i < columnNum; i++) {
            Cell cell = sheet.getRow(rowNum).getCell(columnNum);
            if(!checkNeighbourCells(previousCell, nextCell) && !checkCellAvailability(formulaEvaluator, cell)) emptyCount++;
        }
        return emptyCount;
    }
    private void fillCell(XSSFWorkbook wb, Cell cell){
        Font font = wb.createFont();
        editFont(font);

        CellStyle cellStyle = wb.createCellStyle();
        editCellStyle(cellStyle, font);
        cell.setCellValue("N");
        cell.setCellStyle(cellStyle);
    }

//    private void checkFile(FormulaEvaluator formulaEvaluator, XSSFSheet sheet, XSSFWorkbook wb){
//        for (int j = 3; j < columnNum; j++) {
//            for (int i = 14; i < 30; i++) {
//                if(skippingIndexes.contains(i)) continue;
//                Cell cell = sheet.getRow(i).getCell(j);
//                Cell previousCell = null;
//                Cell nextCell = null;
//                if(j != 3) {
//                    previousCell = sheet.getRow(i).getCell(j-1);
//                }
//                if(j != columnNum){
//                    nextCell = sheet.getRow(i).getCell(j+1);
//                }
//                if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.STRING)) {
//                    checkCellString(cell, i);
//                } else if (formulaEvaluator.evaluateInCell(cell).getCellTypeEnum().equals(CellType.NUMERIC)) {
//                    System.out.println("Numara olmaması gereken '" + cell.getAddress() + "' hücresinde numara var. Lütfen excel dosyasını düzelt.");
//                } else {
//                    checkNeighbourCells(previousCell, nextCell);
//                    if(isPreviousClear && isNextClear) availableCells.add(cell);
//                    isPreviousClear = true; isNextClear = true;
//                }
//            }
//            fillCells(availableCells, wb);
//            resetLists(availableCells, offDoctors, watchingDoctors);
//        }


//        if (checkWatchings(formulaEvaluator, sheet)) {
//            resetLists(availableCells, offDoctors, watchingDoctors);
//            clearCells(filledCells);
//            checkFile(formulaEvaluator, sheet, wb);
//        } else {
//            fillCells(availableCells, wb);
//        }
//    }
//
//
//    private boolean checkWatchings(FormulaEvaluator formulaEvaluator, XSSFSheet sheet) {
//        for (int i = 14; i < 30; i++) {
//            if (skippingIndexes.contains(i)) {
//                continue;
//            }
//
//            int nCount = 0;
//            int emptyCount = 0;
//
//            for (int j = 3; j < columnNum; j++) {
//                Row row = sheet.getRow(i);
//                Cell cell = row.getCell(j);
//                Cell previousCell = (j != 3) ? row.getCell(j - 1) : null;
//                Cell nextCell = (j != columnNum - 1) ? row.getCell(j + 1) : null;
//
//                CellType cellType = formulaEvaluator.evaluateInCell(cell).getCellTypeEnum();
//
//                if (cellType == CellType.STRING) {
//                    String cellValue = cell.getStringCellValue();
//                    if ("N".equals(cellValue)) {
//                        nCount++;
//                    }
//                } else if (cellType == CellType.NUMERIC) {
//                    System.out.println("Numara olmaması gereken '" + cell.getAddress() + "' hücresinde numara var. Lütfen excel dosyasını düzelt.");
//                } else {
//                    if (previousCell != null && previousCell.getCellTypeEnum() == CellType.STRING && "N".equals(previousCell.getStringCellValue())) {
//                        isPreviousClear = false;
//                    }
//                    if (nextCell != null && nextCell.getCellTypeEnum() == CellType.STRING && "N".equals(nextCell.getStringCellValue())) {
//                        isNextClear = false;
//                    }
//                    if (isPreviousClear && isNextClear) {
//                        emptyCount++;
//                    }
//                    isPreviousClear = true;
//                    isNextClear = true;
//                }
//            }
//            if (emptyCount + nCount >= 7 || nCount < 7 || nCount > 12) {
//                return true;
//            }
//        }
//        return false;
//    }



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
                        filledCells.add(cellList.get(j));
                        watchingDoctors.add(j);
                    }
        }
    }
    private void clearCells(List<Cell> cellList){
        System.out.println("cells cleared");
        for (int i = 0; i < cellList.size(); i++) {
            cellList.get(i).setCellValue("");
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

