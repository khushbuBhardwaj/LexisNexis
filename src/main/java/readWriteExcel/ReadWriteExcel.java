package readWriteExcel;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import utils.Constants;
import utils.ReadWriteHelpers;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

public class ReadWriteExcel {
    Logger logger = LoggerFactory.getLogger(ReadWriteExcel.class);

    private Sheet outputSheet;
    private Workbook writerWorkbook;
    private Workbook readerWorkBook;
    private Date currentDateReceived;
    private Date currentDateOfDecision;
    private Map<String, Integer> vciCodesMap = new LinkedHashMap<>();

    public void readWriteExcel() throws IOException {
        ClassLoader classLoader = getClass().getClassLoader();
        File file = new File(classLoader.getResource(Constants.INPUT_FILE_PATH).getFile());
        readerWorkBook = WorkbookFactory.create(file);
        writerWorkbook = new XSSFWorkbook();

        logger.info("Workbook has total sheets {}", readerWorkBook.getNumberOfSheets());

        /*
           =============================================================
           Iterating over all the sheets in the workbook
           =============================================================
        */

        Sheet sheet = readerWorkBook.getSheetAt(0);
        AtomicInteger i = new AtomicInteger();

        sheet.forEach(row -> {
            if (i.get() == 0) {
                outputSheet = createSheetAndHeader(row);
                i.getAndIncrement();
            } else {
                if (writeSheetRows(outputSheet, row, i.get())) {
                    i.getAndIncrement();
                }
            }
        });
        writeDateToSheet();
    }

    private void writeDateToSheet() throws IOException {
        logger.info("Write data to output sheet and close file IO");

        File testResultsFolder = new File(Constants.OUTPUT_FILE_PATH + Constants.OUTPUT_FILE_NAME);
        FileOutputStream fileOut = new FileOutputStream(testResultsFolder);
        writerWorkbook.write(fileOut);
        fileOut.close();
        ReadWriteHelpers.closeWorkbook(readerWorkBook);
        ReadWriteHelpers.closeWorkbook(writerWorkbook);
    }


    private boolean writeSheetRows(Sheet outputSheet, Row inputRow, int rowNumber) {
        logger.info("Write data rows to output sheet rowNumber:::{}", rowNumber);

        AtomicInteger i = new AtomicInteger(0);
        String vciCodeVale = inputRow.getCell(0).getRichStringCellValue().toString().trim();
        boolean writeThisEntry = true;

        if (vciCodesMap.containsKey(vciCodeVale)) {
            writeThisEntry = false;
        } else {
            vciCodesMap.put(vciCodeVale, 1);
        }
        if (writeThisEntry) {
            Row outputRow = outputSheet.createRow(rowNumber);

            inputRow.forEach(inputCell -> {
                Cell outputCell = outputRow.createCell(i.get());
                writeCellByCellType(inputCell, outputCell, i.get());
                i.getAndIncrement();
            });
            Cell timeTakenCell = outputRow.createCell(i.get());
            timeTakenCell.setCellValue(ReadWriteHelpers.getDaysFromDates(currentDateReceived, currentDateOfDecision));
            ReadWriteHelpers.resizeSheetColumns(i.intValue() + 1, outputSheet);

        }
        return writeThisEntry;
    }


    private void writeCellByCellType(Cell inputCell, Cell outputCell, int columnNumber) {

        String cellColumnName = inputCell.getSheet().getRow(0).getCell(columnNumber).getRichStringCellValue().toString().trim();
        logger.info("Write data cell::{}", cellColumnName);

        switch (inputCell.getCellTypeEnum()) {
            case BOOLEAN:
                outputCell.setCellValue(inputCell.getBooleanCellValue());
                break;
            case STRING:
                outputCell.setCellValue(inputCell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(inputCell)) {
                    dateCellStyle(outputCell, inputCell, cellColumnName);
                } else {
                    outputCell.setCellValue(inputCell.getNumericCellValue());
                    dateCellStyle(outputCell, inputCell, cellColumnName);
                }
                break;
            default:
                break;
        }
    }

    private void dateCellStyle(Cell outputCell, Cell inputCell, String cellName) {
        logger.info("Formatting date type fields::{}", cellName);

        CellStyle dateCellStyle = writerWorkbook.createCellStyle();
        CreationHelper createHelper = writerWorkbook.getCreationHelper();
        switch (cellName) {
            case "Date Recv'd":
                currentDateReceived = inputCell.getDateCellValue();
                outputCell.setCellValue(inputCell.getDateCellValue());
                dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(Constants.DATE_FORMATTER_RECEIVED_DATE));
                break;

            case "DATE OF DECISION":
                currentDateOfDecision = inputCell.getDateCellValue();
                outputCell.setCellValue(inputCell.getDateCellValue());
                dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(Constants.DATE_FORMATTER_DECISION_DATE));
                break;

            case "Time Recv'd":
                outputCell.setCellValue(ReadWriteHelpers.timeFormatter(inputCell.getDateCellValue()));
                break;

            case "Time Ingested (MNL Time)":
                outputCell.setCellValue(ReadWriteHelpers.aestTimeFormatter(inputCell.getDateCellValue()));
                break;
            default:
                break;
        }
        outputCell.setCellStyle(dateCellStyle);
    }

    /**
     * Method to create header cells
     *
     * @return sheet
     */
    private Sheet createSheetAndHeader(Row inputHeaderRow) {
        logger.info("Creating sheet header");

        Sheet sheet = writerWorkbook.getSheet("Output");
        if (sheet == null) {
            logger.info("Steep already exists in workbook");
            sheet = writerWorkbook.createSheet("Output");
        }
        Font headerFont = writerWorkbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);

        CellStyle headerCellStyle = writerWorkbook.createCellStyle();
        headerCellStyle.setFillBackgroundColor(IndexedColors.GREY_80_PERCENT.getIndex());

        headerCellStyle.setFont(headerFont);
        Row outputHeaderRow = sheet.createRow(0);

        AtomicInteger i = new AtomicInteger(0);
        inputHeaderRow.forEach(inputCell -> {
            Cell outputCell = outputHeaderRow.createCell(i.get());
            outputCell.setCellValue(inputCell.getRichStringCellValue().getString());
            outputCell.setCellStyle(headerCellStyle);
            i.getAndIncrement();
        });

        //Create New header cell for time taken days
        Cell outputCell = outputHeaderRow.createCell(i.get());
        outputCell.setCellValue(Constants.TIME_TAKEN_COLUMN_NAME);
        outputCell.setCellStyle(headerCellStyle);

        //Resize column
        ReadWriteHelpers.resizeSheetColumns(i.intValue() + 1, sheet);

        return sheet;
    }

}