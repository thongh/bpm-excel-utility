package excel;

import java.io.File;

//import java.util.logging.Logger;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {
    
    // assumes the current class is called MyLogger
//    private final static Logger LOGGER = Logger.getLogger(ExcelUtility.class.getName());
    
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
        Workbook kb = readExcelFile(kweseBringgPath);

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
    }

    private static void addColumnAmount(String amountFileName, Workbook kb,
            Workbook wbT) {
   
        double amount = AMOUNT - AMOUNT * 0.1;
        Iterator<Row> kbIter = kb.getSheetAt(0).iterator();
        
        // Set header, this only happens once
        Row headerRow = kb.getSheetAt(0).getRow(0);
        Cell amountHeader = headerRow.createCell(headerRow.getLastCellNum());
        amountHeader.setCellType(CellType.STRING);
        amountHeader.setCellValue("AMOUNT");

        while (kbIter.hasNext()) {
            // Get current excel row
            Row currentRow = kbIter.next();
            // Create a cell to store value
            if (currentRow.getRowNum() != 0) {
                // Create numertic cell to store value
                Cell cell = currentRow.createCell(headerRow.getLastCellNum() - 1,
                        CellType.NUMERIC);
                
                // Get assigned team
                String assignedTeam = "";
                if (currentRow.getCell(16) != null && CellType.NUMERIC
                        .equals(currentRow.getCell(16).getCellType())) {
                    assignedTeam = new String(
                            "" + currentRow.getCell(16).getNumericCellValue());// ASSIGNED
                                                                               // TEAM
                } else {
                    assignedTeam = currentRow.getCell(16).getStringCellValue().trim();// ASSIGNED
                                                                              // TEAM
                }
                
                // Check if this name is in the tax clearnance list
                boolean isTaxClearnance = isTaxClearnance(wbT, assignedTeam.trim());
                
                // Set amount value
                if (isTaxClearnance) {
                    amount = AMOUNT;
                }
                cell.setCellValue(amount);              
            }
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
    }

    private static boolean isTaxClearnance(Workbook wbT, String name) {
        Iterator<Row> wbTIter = wbT.getSheetAt(0).iterator();
        boolean validTaxClear = false;        
        while (wbTIter.hasNext()) {
            Row currentRow = wbTIter.next();
            if (currentRow != null && currentRow.getCell(2) != null ) {
                StringBuilder sb = new StringBuilder();
                sb.append(currentRow.getCell(2).getStringCellValue());
                sb.append(" ");
                sb.append(currentRow.getCell(3).getStringCellValue());
                if (sb.toString().equals(name)) {                 
//                    LOGGER.info("Looking for name: " + name + ". Found: ");
//                    LOGGER.info("valid? " + currentRow.getCell(5).getStringCellValue());              
                    validTaxClear = currentRow.getCell(5).getStringCellValue().trim().equals("YES") ? true : false;
                    break;  // Stop this loop as we have found what we need                               
                }
            }
        }
        return validTaxClear;
    }
    
    private static Workbook readExcelFile(String fileName) {
        Workbook workbook = null; 
        try {
            FileInputStream excelFile = new FileInputStream(new File(fileName));
            workbook = new XSSFWorkbook(excelFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }
}
