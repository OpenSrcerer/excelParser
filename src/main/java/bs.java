import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.jetbrains.annotations.NotNull;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

public class bs
{
    private static int successful, failed;
    private static XSSFWorkbook xssfSheets;
    private static Logger logger;
    private static CellStyle style;

    static {
        logger = LoggerFactory.getLogger(bs.class);
    }

    public static int parseTime(@NotNull String mins, @NotNull String secs)
    {
        int outputInSeconds = 0;
        try {
            outputInSeconds = (Integer.parseInt(mins) * 60) + Integer.parseInt(secs);
        } catch (NumberFormatException ex) {
            ++failed;
        }
        return outputInSeconds;
    }

    public static void main (String[] args)
    {
        List<String> fileList;
        {
            File dir = new File("./resources");
            String[] dirlist = dir.list();

            if (dirlist == null) {
                logger.error("Directory is empty.");
                return;
            }

            fileList = Arrays.asList(dirlist);
        }


        for (String file : fileList) {

            File inputFile = new File("./resources/" + file);

            try {
                FileInputStream inputStream = new FileInputStream(inputFile);
                xssfSheets = new XSSFWorkbook(inputStream);
                style = xssfSheets.createCellStyle();
                style.setBorderTop(CellStyle.BORDER_THIN);
                style.setBorderLeft(CellStyle.BORDER_THIN);
                style.setBorderRight(CellStyle.BORDER_THIN);
                style.setBorderBottom(CellStyle.BORDER_THIN);
                style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                style.setTopBorderColor(IndexedColors.BLACK.getIndex());
                style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                style.setRightBorderColor(IndexedColors.BLACK.getIndex());
                inputStream.close();
            } catch (IOException ex) {
                logger.error("I/O Operations Failure!", ex);
            }

            for (XSSFSheet currentSheet : xssfSheets) {
                for (Row row : currentSheet) {
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();

                        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                            if (Pattern.matches("\\d{2}m \\d{2}s", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}' \\d{2}\"", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}' \\d{2}''", cell.getStringCellValue())) {

                                String minuteValue = cell.getStringCellValue().substring(0, 2);
                                String secValue = cell.getStringCellValue().substring(4, 6);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minuteValue, secValue));
                                ++successful;

                            } else if (Pattern.matches("\\d{2}m\\d{2}s", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}'\\d{2}\"", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}'\\d{2}''", cell.getStringCellValue())) {

                                String minuteValue = cell.getStringCellValue().substring(0, 2);
                                String secValue = cell.getStringCellValue().substring(3, 5);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minuteValue, secValue));
                                ++successful;

                            } else if (Pattern.matches("\\dm \\d{2}s", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d' \\d{2}\"", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d' \\d{2}''", cell.getStringCellValue())) {

                                String minuteValue = cell.getStringCellValue().substring(0, 1);
                                String secValue = cell.getStringCellValue().substring(3, 5);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minuteValue, secValue));
                                ++successful;

                            } else if (Pattern.matches("\\dm\\d{2}s", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d'\\d{2}\"", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d'\\d{2}''", cell.getStringCellValue())) {
                                String minuteValue = cell.getStringCellValue().substring(0, 1);
                                String secValue = cell.getStringCellValue().substring(2, 4);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minuteValue, secValue));
                                ++successful;

                            } else if (Pattern.matches("\\d{2}m \\ds", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}' \\d\"", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}' \\d''", cell.getStringCellValue())) {

                                String minuteValue = cell.getStringCellValue().substring(0, 2);
                                String secValue = cell.getStringCellValue().substring(4, 5);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minuteValue, secValue));
                                ++successful;

                            } else if (Pattern.matches("\\d{2}m\\ds", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}'\\d\"", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}'\\d''", cell.getStringCellValue())) {
                                String minuteValue = cell.getStringCellValue().substring(0, 2);
                                String secValue = cell.getStringCellValue().substring(3, 4);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minuteValue, secValue));
                                ++successful;

                            } else if (Pattern.matches("\\dm \\ds", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d' \\d\"", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d' \\d''", cell.getStringCellValue())) {

                                String minuteValue = cell.getStringCellValue().substring(0, 1);
                                String secValue = cell.getStringCellValue().substring(3, 4);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minuteValue, secValue));
                                ++successful;

                            } else if (Pattern.matches("\\dm\\ds", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d'\\d\"", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d'\\d''", cell.getStringCellValue())) {
                                String minuteValue = cell.getStringCellValue().substring(0, 1);
                                String secValue = cell.getStringCellValue().substring(2, 3);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minuteValue, secValue));
                                ++successful;

                            } else if (Pattern.matches("\\ds", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d\"", cell.getStringCellValue())) {

                                String secValue = cell.getStringCellValue().substring(0, 1);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime("0", secValue));
                                ++successful;

                            } else if (Pattern.matches("\\d{2}s", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}\"", cell.getStringCellValue())) {
                                String secValue = cell.getStringCellValue().substring(0, 2);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime("0", secValue));
                                ++successful;
                            } else if (Pattern.matches("\\d{2}m", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d{2}'", cell.getStringCellValue())) {
                                String minValue = cell.getStringCellValue().substring(0, 2);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minValue, "0"));
                                ++successful;
                            } else if (Pattern.matches("\\dm", cell.getStringCellValue()) ||
                                    Pattern.matches("\\d'", cell.getStringCellValue())) {
                                String minValue = cell.getStringCellValue().substring(0, 1);

                                cell.setCellStyle(style);
                                cell.setCellValue(parseTime(minValue, "0"));
                                ++successful;
                            }
                        }
                    }
                }
            }


            try {
                FileOutputStream fOS = new FileOutputStream(inputFile);
                xssfSheets.write(fOS);
                fOS.close();
                logger.info("[" + file + "] " + successful + " successful and " + failed + " failed cell manipulations.");
                successful = 0;
            } catch (IOException ex) {
                logger.error("I/O Operations Failure!", ex);
            }
        }
        logger.info("Parsing Excel file(s) complete!");
    }
}
