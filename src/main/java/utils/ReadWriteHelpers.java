package utils;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import readWriteExcel.ReadWriteExcel;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;
import java.util.concurrent.TimeUnit;

/**
 * Helper class
 */
public class ReadWriteHelpers {

    private static  Logger logger = LoggerFactory.getLogger(ReadWriteHelpers.class);

    /**
     * Close workbook instance
     *
     * @param workbook
     * @throws IOException
     */
    public static void closeWorkbook(Workbook workbook) throws IOException {
        logger.info("Closing workbook");
        workbook.close();
    }

    /**
     * Method to calculate no of days between to dates
     *
     * @param received
     * @param decision
     * @return long
     */
    public static long getDaysFromDates(Date received, Date decision){
        long diff = received.getTime() - decision.getTime();
        logger.info("Calculated days:::"+diff);

        return  TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
    }

    /**
     * Method to resize sheet column to wrap content
     *
     * @param length
     * @param sheet
     */
    public static void resizeSheetColumns(int length, Sheet sheet){
        logger.info("Resizing column");
        for(int j=0;j<length;j++){
            sheet.autoSizeColumn(j);
        }
    }

    /**
     * Method to format local time  to AEST time Zone
     *
     * @param date
     * @return
     */
    public static String aestTimeFormatter(Date date){
        SimpleDateFormat formatter = new SimpleDateFormat(Constants.TIME_FORMATTER);
        formatter.setTimeZone(TimeZone.getTimeZone(Constants.AEST_TIME_ZONE));
        return formatter.format(date);
    }

    /**
     * Formatter to format time
     *
     * @param date
     * @return string
     */
    public static String timeFormatter(Date date){
        SimpleDateFormat formatTime = new SimpleDateFormat(Constants.TIME_FORMATTER);
        return formatTime.format(date);
    }
}