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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {

    // assumes the current class is called MyLogger
    // private final static Logger LOGGER =
    // Logger.getLogger(ExcelUtility.class.getName());

    // public static final String FOLDER = "D:\\LibJar\\";
    public static final Integer AMOUNT = 1000;

    // If file merge of Kwese and Bringg is created, just load the merge file
    public static final boolean isCreateFileMerge = false;

    public static void main(String[] args) {
//        System.out.println("Start App");

        mergeExcel("C:\\temp\\files\\", "SMBS.xlsx", "Bringg Info.xlsx",
                "Super Technites Tax clearance certificates.xlsx",
                "KweseBringg.xlsx",
                "Amount.xlsx");
        
//        System.out.println("----END APP----");
    }

    /**
     * 
     * Method to create an Excel file with the following rules: - Read
     * KweseBringg.xlsx - Read Super Technites Tax clearance certificates.xlsx -
     * For each row in KweseBringg.xlsx add a cell AMOUNT - AMOUNT = 1000 if not
     * have tax clearance - AMOUNT = AMOUNT - AMOUNT * 0.1 if have tax clearance
     * - Produce final excel file with AMOUNT
     * 
     * @param folderPath:
     *            directory for the files
     * @param kwesePath:
     *            kwese file name
     * @param bringgPath:
     *            bringg file name
     * @param superPath:
     *            super file name
     * @param kweseBringgPath:
     *            kwese bringg file name
     * @param amountPath:
     *            final excel file name
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

        String kwese = folderPath + "SMBS.xlsx";
        Workbook wbK = readExcelFile(kwese);

        String bringg = folderPath + "Bringg Info.xlsx";
        Workbook wbB = readExcelFile(bringg);

        String kweseBringgFileName = folderPath + "KweseBringg.xlsx";
        kweseBringgPath = folderPath + kweseBringgPath;

        /**
         * Merge 2 file Kwese Fields, Bringg fields by JobNum in Kwese Fields =
         * CUMII ORDER ID in Bringg fields This run take long time.
         */

        Workbook kb = mergerKweseBringg(kweseBringgFileName, wbK, wbB);
        // Workbook kb = readExcelFile(kweseBringgPath);

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

    /**
     * 
     * Method to do the following: - Read merged KweseBringg file - Add 'AMOUNT'
     * header - For each row in KweseBringg, check if team has tax clearance -
     * Set 'AMOUNT' value accordingly - Once loop finish, create final Excel
     * file Amount.xlsx
     * 
     * @param amountFileName:
     *            name of final xlsx file
     * @param kb:
     *            kweseBringg file name
     * @param wbT:
     *            tax file name
     */
    private static void addColumnAmount(String amountFileName, Workbook kb,
            Workbook wbT) {
      
        double amount = AMOUNT - AMOUNT * 0.1;
        Iterator<Row> kbIter = kb.getSheetAt(0).iterator();

        // Set header, this only happens once
        Row headerRow = kb.getSheetAt(0).getRow(0);
        Cell amountHeader = headerRow.createCell(headerRow.getLastCellNum());
        amountHeader.setCellType(CellType.STRING);
        amountHeader.setCellValue("AMOUNT");
        
        // Get ASSIGNED TEAM col index
        int assignedTeamColIndex = 0;
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().trim().toUpperCase()
                    .equals("ASSIGNED TEAM")) {
                assignedTeamColIndex = cell.getColumnIndex();
                break;
            }
        }

        while (kbIter.hasNext()) {
            // Get current excel row
            Row currentRow = kbIter.next();
            // Create a cell to store value
            if (currentRow.getRowNum() != 0) {
                // Create numertic cell to store value
                Cell cell = currentRow.createCell(
                        headerRow.getLastCellNum() - 1, CellType.NUMERIC);            
                
                // Get assigned team
                String assignedTeam = "";
                if (currentRow.getCell(assignedTeamColIndex) != null) {
                    if (CellType.NUMERIC
                            .equals(currentRow.getCell(assignedTeamColIndex).getCellType())) {
                        assignedTeam = new String(
                                "" + currentRow.getCell(assignedTeamColIndex).getNumericCellValue());
                    } else {
                        assignedTeam = currentRow.getCell(assignedTeamColIndex).getStringCellValue()
                                .trim();
                    }
                }             

                // Check if this name is in the tax clearnance list
                boolean isTaxClearnance = isTaxClearnance(wbT,
                        assignedTeam.trim());

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

    /**
     * 
     * Check if a team has tax clearance
     * 
     * @param wbT:
     *            name of tax file
     * @param name:
     *            team name
     */
    private static boolean isTaxClearnance(Workbook wbT, String name) {
        Iterator<Row> wbTIter = wbT.getSheetAt(0).iterator();
        boolean validTaxClear = false;
        while (wbTIter.hasNext()) {
            Row currentRow = wbTIter.next();
            if (currentRow != null && currentRow.getCell(2) != null) {
                StringBuilder sb = new StringBuilder();
                sb.append(currentRow.getCell(2).getStringCellValue());
                sb.append(" ");
                sb.append(currentRow.getCell(3).getStringCellValue());
                if (sb.toString().equals(name)) {
                    validTaxClear = currentRow.getCell(5).getStringCellValue()
                            .trim().equals("YES") ? true : false;
                    break; // Stop this loop as we have found what we need
                }
            }
        }
        return validTaxClear;
    }

    /**
     * 
     * Read an xlsx file into POI's Workbook
     * 
     * @param fileName:
     *            xlsx file name
     * @return: Workbook object of the xlsx file
     */
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

    /**
     * 
     * Merge Kwese and Bringg files together following the rules: - Read
     * SMBS.xlsx file into Workbook object - Read Bringg Info.xlsx into Workbook
     * object - For each sheet/month, merge SMBS and Bringg Info by job no. -
     * Produce final xlsx file with correct sheets/months and merged data
     * 
     * @param kweseBringgFileName:
     *            final output xlsx file name
     * @param smbs:
     *            SMSB Workbook
     * @param bringgInfo:
     *            Bringg Info Workbook
     * @return: Workbook object of the xlsx file
     */
    private static Workbook mergerKweseBringg(String kweseBringgFileName,
            Workbook smbs, Workbook bringgInfo) {

        // For each sheet, iterate through each row and get the job number
        // For each row, search for job number in Bringg Info xlsx, if found,
        // add cell values to SMBS file
        // Continue the next rows in SMSB file until complete all
        // Continue the next sheets until complete all
        // Produce the final xlsx file

        // Iterate through each sheet in SMBS file
        for (int i = 0; i < smbs.getNumberOfSheets(); i++) {
            // Row iterator
            Iterator<Row> smbsIter = smbs.getSheetAt(i).iterator();
            // Get Bringg Info header
            Row bringgHeaderRow = bringgInfo.getSheetAt(i).getRow(0);
            // Get SMBS Header
            Row smbsHeaderRow = smbs.getSheetAt(i).getRow(0);
            // Loop through each Bringg header and add to SMSB header
            for (Cell cell : bringgHeaderRow) {
                Cell newHeaderCell = smbsHeaderRow.createCell(
                        smbsHeaderRow.getLastCellNum(), cell.getCellType());
                newHeaderCell.setCellValue(cell.getStringCellValue());
            }
            // Get jobnum column index
            int jobNumColIndex = 0;
            for (Cell cell : smbsHeaderRow) {
                if (cell.getStringCellValue().trim().toUpperCase()
                        .equals("JOBNUM")) {
                    jobNumColIndex = cell.getColumnIndex();
                    break;
                }
            }
            // Iterate through each row in SMBS and get the job number
            while (smbsIter.hasNext()) {
                // Get current row
                Row currentRow = smbsIter.next();
                // Get job number
                String jobNum = "";
                if (currentRow.getCell(jobNumColIndex) != null
                        && CellType.NUMERIC.equals(currentRow
                                .getCell(jobNumColIndex).getCellType())) {
                    jobNum = new String("" + currentRow.getCell(jobNumColIndex)
                            .getNumericCellValue());
                } else {
                    jobNum = currentRow.getCell(jobNumColIndex)
                            .getStringCellValue().trim();
                }
         
                // Seach in Bringg Info in this sheet for the same job num and
                // return the whole row
                Row bringgInfoRow = findBringgInfoByJobNum(
                        bringgInfo.getSheetAt(i), jobNum);
                if (bringgInfoRow != null) {
                    // Copy over the whole row into smbs row
                    for (int j = 0; j < bringgInfoRow.getLastCellNum(); j++) {
                        Cell cell = bringgInfoRow.getCell(j);
                        Cell newCell = null;
                        if (cell != null) {
                            if (cell.getCellType().equals(CellType.NUMERIC)) {
                                double numericVal = cell.getNumericCellValue();
                                newCell = currentRow.createCell(
                                        currentRow.getLastCellNum(),
                                        cell.getCellType());
                                newCell.setCellValue(numericVal);
                            } else if (cell.getCellType().equals(CellType.STRING)) {
                                String strVal = cell.getStringCellValue();
                                if (strVal.equals("")) {
                                    newCell = currentRow.createCell(
                                            currentRow.getLastCellNum(),
                                            CellType.STRING);
                                    newCell.setCellValue("");
                                } else {
                                    newCell = currentRow.createCell(
                                            currentRow.getLastCellNum(),
                                            cell.getCellType());
                                    newCell.setCellValue(strVal);
                                }                          
                            }
                        } else {
                            newCell = currentRow.createCell(
                                    currentRow.getLastCellNum(),
                                    CellType.BLANK);
                        }
                    }
                }
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(
                    kweseBringgFileName);
            smbs.write(outputStream);
            // workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return smbs;
    }

    /**
     * 
     * Find a row in Bringg Info by job number
     * 
     * @param sheet:
     *            the sheet to look into
     * @param jobNum:
     *            the job number to compare against
     * @return: found row
     */
    private static Row findBringgInfoByJobNum(Sheet sheet, String jobNum) {
        Row matchedRow = null;
        for (Row row : sheet) {
            if (row.getCell(1).getStringCellValue().trim().equals(jobNum)) {
                matchedRow = row;
                break;
            }
        }
        return matchedRow;
    }
}
