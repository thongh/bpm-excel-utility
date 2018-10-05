package excel;

import java.io.File;
import java.util.logging.Logger;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {
    
    // assumes the current class is called MyLogger
    private final static Logger LOGGER = Logger.getLogger(ExcelUtility.class.getName());
    
    // public static final String FOLDER = "D:\\LibJar\\";
    public static final Integer AMOUNT = 1000;

    // If file merge of Kwese and Bringg is created, just load the merge file
    public static final boolean isCreateFileMerge = false;

    public static void main(String[] args) {
//        System.out.println("Start App");
        mergeExcel("C:\\temp\\files\\", "Kwese Fields .xlsx", "Bringg fields.xlsx",
                "Super Technites Tax clearance certificates.xlsx",
                "KweseBringg.xlsx",
                "Amount.xlsx");
//        System.out.println("----END APP----");
    }

    
    /**
     * 
     * Method to create an Excel file with the following rules:
     *  - Read KweseBringg.xlsx
     *  - Read Super Technites Tax clearance certificates.xlsx
     *  - For each row in KweseBringg.xlsx add a cell AMOUNT 
     *  - AMOUNT = 1000 if not have tax clearance
     *  - AMOUNT = AMOUNT - AMOUNT * 0.1 if have tax clearance
     *  - Produce final excel file with AMOUNT
     * 
     * @param folderPath: directory for the files
     * @param kwesePath: kwese file name
     * @param bringgPath: bringg file name
     * @param superPath: super file name
     * @param kweseBringgPath: kwese bringg file name
     * @param amountPath: final excel file name
     */
    public static void mergeExcel(String folderPath, String kwesePath,
            String bringgPath, String superPath, String kweseBringgPath,
            String amountPath) {
        LOGGER.info("Start merge Excel: ");
        // Folder
        if (folderPath == "") {
            folderPath = "/tmp/files/";
        }

        // superT
        if (superPath == "") {
            superPath = "Super Technites Tax clearance certificates.xlsx";
        }
        superPath = folderPath + superPath;
        Workbook wbT = readExcelFile(superPath);

        // Kwese Bringg
        if (kweseBringgPath == "") {
            kweseBringgPath = "KweseBringg.xlsx";
        }
        kweseBringgPath = folderPath + kweseBringgPath;

        /**
         * Merge 2 file Kwese Fields, Bringg fields by JobNum in Kwese Fields =
         * CUMII ORDER ID in Bringg fields This run take long time.
         */
        Workbook kb = null;
        if (isCreateFileMerge) {
            // Kwese
            if (kwesePath == "") {
                kwesePath = "Kwese Fields .xlsx";
            }
            kwesePath = folderPath + kwesePath;
            Workbook wbK = readExcelFile(kwesePath);

            // Bringg
            if (bringgPath == "") {
                bringgPath = "Bringg fields.xlsx";
            }
            bringgPath = folderPath + bringgPath;
            Workbook wbB = readExcelFile(bringgPath);

            kb = mergerKweseBringg(kweseBringgPath, wbK, wbB);
        } else {
            kb = readExcelFile(kweseBringgPath);
        }

        /**
         * Compare merge file with "Super Technites Tax clearance certificates"
         * by column ASSIGNED TEAM = Company Name No logic involve with column
         * Valid Tax clearance
         */
        if (amountPath == "") {
            amountPath = "Amount.xlsx";
        }
        amountPath = folderPath + amountPath;
        addColumnAmount(amountPath, kb, wbT);
        LOGGER.info("Done merge Excel: ");
    }

    private static void addColumnAmount(String amountFileName, Workbook kb,
            Workbook wbT) {
        LOGGER.info("Start create AMOUNT cell: " + amountFileName);
        
        double amount = AMOUNT - AMOUNT * 0.1;
        List<Cell[]> tData = getAllFromTaxClearance(wbT);
        Iterator<Row> kbIter = kb.getSheetAt(0).iterator();
        String name = "";
        Row currentRow = null;
        boolean isExistInTaxFile = false;
        
        int counter = 0;
        while (kbIter.hasNext()) {
            LOGGER.info("loop through excel at: " + counter);
            // Get current excel row
            currentRow = kbIter.next();
            
            // Set header, this only happens once
            if (currentRow.getRowNum() == 0) {
                Cell cell = kbIter.next().createCell(currentRow.getLastCellNum(),
                        CellType.STRING);
                cell.setCellValue("AMOUNT");
            }
            
            // Create numertic cell to store value
            Cell cell = currentRow.createCell(currentRow.getLastCellNum(),
                    CellType.NUMERIC);
            // Get assigned team
            if (currentRow.getCell(16) != null && CellType.NUMERIC
                    .equals(currentRow.getCell(16).getCellType())) {
                name = new String(
                        "" + currentRow.getCell(16).getNumericCellValue());
            } else {
                name = currentRow.getCell(16).getStringCellValue().trim();
            }
            // Check if this name is in the tax clearnance list
            isExistInTaxFile = findName(tData, name);
            // Set amount value
            if (isExistInTaxFile) {
                amount = AMOUNT;
            }
            cell.setCellValue(amount);
            counter++;
        }
        
        // After loop is done, create excel
        try {
            FileOutputStream outputStream = new FileOutputStream(
                    amountFileName);
            kb.write(outputStream);
            wbT.close();
            kb.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        LOGGER.info("Done create AMOUNT cell: " + amountFileName);
    }

    private static boolean findName(List<Cell[]> tData, String name) {
        LOGGER.info("Start checking tax clearnance for: " + name);
        Cell[] row;
        boolean isTaxClearnance = false;
        for (int i = 0; i < tData.size(); i++) {
            // ignore header
            if (i < 4) {
                continue;
            }
            row = tData.get(i);

            Cell companyCell = row[0];// Name
            String companyName = companyCell.getStringCellValue().trim();
            if (companyName.equalsIgnoreCase(name)) {
                isTaxClearnance = true;
                break;
            }
        
        }
        LOGGER.info("Done checking tax clearnance for: " + name);
        return isTaxClearnance;
    }

    private static List<Cell[]> getAllFromTaxClearance(Workbook wbT) {
        LOGGER.info("Start getAllFromTaxClearance: ");
        List<Cell[]> tData = new ArrayList<Cell[]>();
        Iterator<Row> iterator = wbT.getSheetAt(0).iterator();
        Row currentRow = null;
        Cell currentCell = null;
        Cell[] rowData;
        
        while (iterator.hasNext()) {
            currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            List<Cell> rwDataLst = new ArrayList<Cell>();
            while (cellIterator.hasNext()) {
                currentCell = cellIterator.next();
                rwDataLst.add(currentCell);
            }
            rowData = new Cell[rwDataLst.size()];
            rowData = rwDataLst.toArray(rowData);
            tData.add(rowData);
        }
        LOGGER.info("Done getAllFromTaxClearance: ");
        return tData;
    }

    private static void createHeaderKB(XSSFSheet outS, Workbook k, Workbook b) {
        LOGGER.info("Start createHeaderKB: ");
        int colNum = 0;
        Row row = outS.createRow(0);

        Iterator<Row> kIter = k.getSheetAt(0).iterator();
        Row currentRow = null;
        Cell writeCell = null;
        Cell currentCell = null;
        while (kIter.hasNext()) {
            currentRow = kIter.next();
            if (currentRow.getRowNum() == 0) {
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    writeCell = row.createCell(colNum++);
                    currentCell = cellIterator.next();
                    writeCell.setCellValue(currentCell.getStringCellValue());
                }
            } else {
                break;
            }
        }
        Iterator<Row> bIter = b.getSheetAt(0).iterator();
        while (bIter.hasNext()) {
            currentRow = bIter.next();
            if (currentRow.getRowNum() == 0) {
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    writeCell = row.createCell(colNum++);
                    currentCell = cellIterator.next();
                    writeCell.setCellValue(currentCell.getStringCellValue());
                }
            } else {
                break;
            }
        }
        LOGGER.info("Done createHeaderKB: ");
    }

    private static List<Cell[]> getDataFromWorkBook(Workbook wb) {
        LOGGER.info("Start getDataFromWorkBook: ");
        // Kwese
        List<Cell[]> kData = new ArrayList<Cell[]>();
        Iterator<Row> iterator = wb.getSheetAt(0).iterator();
        Row currentRow = null;
        Cell currentCell = null;
        List<Cell> rwDataLst = null;
        Cell[] rowData;
        
        while (iterator.hasNext()) {
            currentRow = iterator.next();
            // Remove Header
            if (currentRow.getRowNum() == 0) {
                continue;
            }
            rwDataLst = new ArrayList<Cell>();
            for (int i = 0; i < currentRow.getLastCellNum(); i++) {
                currentCell = currentRow.getCell(i,
                        Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                rwDataLst.add(currentCell);
            }
            rowData = new Cell[rwDataLst.size()];
            rowData = rwDataLst.toArray(rowData);
            kData.add(rowData);
        }
        LOGGER.info("Start getDataFromWorkBook: ");
        return kData;
    }

    private static Workbook mergerKweseBringg(String kweseBringgFileName,
            Workbook wbK, Workbook wbB) {
        List<Cell[]> kData = getDataFromWorkBook(wbK);
        List<Cell[]> bData = getDataFromWorkBook(wbB);
//        System.out.println("Read Data Finish");

//        System.out.println("Creating Merge Kwese-Bringg excel");
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet outS = workbook.createSheet("Result");

        createHeaderKB(outS, wbK, wbB);

        Cell jobNumCell = null;
        String jobNum = "";
        Cell[] bRow;
        int rowNum = 1;
        for (Cell[] kRow : kData) {
            jobNumCell = kRow[3];// JOBNUM
            jobNum = jobNumCell.getStringCellValue();
//            System.out.print("Search: " + jobNum);
            bRow = findDataByJobNum(bData, jobNum);
            if (bRow != null) {
//                System.out.println(",Found: " + bRow[1].getStringCellValue());
                createKBRow(outS, rowNum, kRow, bRow);
                rowNum++;
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(
                    kweseBringgFileName);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
//        System.out.println("Done create Merge Kwese-Bringg excel");
        return workbook;
    }

    private static void createKBRow(XSSFSheet outS, int rowNum, Cell[] kRow,
            Cell[] bRow) {
        LOGGER.info("Start create final excel row: " + rowNum);
        int colNum = 0;
        Row row = outS.createRow(rowNum);

        Cell writeCell = null;
        for (Cell inCell : kRow) {
            writeCell = row.createCell(colNum++);
            if (inCell.getCellType().equals(CellType.STRING)) {
                writeCell.setCellValue(inCell.getStringCellValue());
            } else if (inCell.getCellType().equals(CellType.NUMERIC)) {
                writeCell.setCellValue(inCell.getNumericCellValue());
            } else if (inCell.getCellType().equals(CellType.BOOLEAN)) {
                writeCell.setCellValue(inCell.getBooleanCellValue());
            } else if (inCell.getCellType().equals(CellType.BLANK)) {
                writeCell.setCellValue("");
            }
        }
        for (Cell inCell : bRow) {
            writeCell = row.createCell(colNum++);
            if (inCell.getCellType().equals(CellType.STRING)) {
                writeCell.setCellValue(inCell.getStringCellValue());
            } else if (inCell.getCellType().equals(CellType.NUMERIC)) {
                writeCell.setCellValue(inCell.getNumericCellValue());
            } else if (inCell.getCellType().equals(CellType.BOOLEAN)) {
                writeCell.setCellValue(inCell.getBooleanCellValue());
            } else if (inCell.getCellType().equals(CellType.BLANK)) {
                writeCell.setCellType(CellType.BLANK);
            }
        }
        LOGGER.info("Done create final excel row: " + rowNum);
    }

    private static Cell[] findDataByJobNum(List<Cell[]> bData, String jobNum) {
        LOGGER.info("Start findDataByJobNum: " + jobNum);
        Cell cell = null;
        for (Cell[] row : bData) {
            // JobNum in Kwese Fields = CUMII ORDER ID in Bringg fields
            //for (int i = 0; i < row.length; i++) {
                cell = row[1];// CUMII ORDER ID
                if (cell.getStringCellValue().trim().equals(jobNum)) {
                    return row;
                }
            //}
        }
        LOGGER.info("Done findDataByJobNum: " + jobNum);
        return null;
    }

    public static Workbook readExcelFile(String fileName) {
        LOGGER.info("Start reading excel: " + fileName);
        Workbook workbook = null; 
        try {
            FileInputStream excelFile = new FileInputStream(new File(fileName));
            workbook = new XSSFWorkbook(excelFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        LOGGER.info("Done reading excel: " + fileName);
        return workbook;
    }
}
