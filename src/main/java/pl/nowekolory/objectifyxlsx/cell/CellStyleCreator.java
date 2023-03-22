package pl.nowekolory.objectifyxlsx.cell;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyleCreator{

    private static final String defaultDateFormat = "dd.MM.yyyy";
    private static final String defaultDateTimeFormat = "dd.MM.yyyy h:mm";

    public static CellStyle createDefaultCellStyle(Workbook workbook){
        return workbook.createCellStyle();
    }
    public static CellStyle createDateTimeCellStyle(Workbook workbook, String format){
        var dateTimeCellStyle = workbook.createCellStyle();
        var stringFormat = getFormat(format, defaultDateTimeFormat);
        var dateTimeFormat = workbook.getCreationHelper().createDataFormat().getFormat(stringFormat);
        dateTimeCellStyle.setDataFormat(dateTimeFormat);
        return dateTimeCellStyle;
    }

    public static CellStyle createDateCellStyle(Workbook workbook, String format){
        var dateCellStyle = workbook.createCellStyle();
        var stringFormat = getFormat(format, defaultDateFormat);
        var dateFormat = workbook.getCreationHelper().createDataFormat().getFormat(stringFormat);
        dateCellStyle.setDataFormat(dateFormat);
        return dateCellStyle;
    }

    public static CellStyle createHeaderCellStyle(Workbook workbook){
        var cellStyle = workbook.createCellStyle();
        var font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        return cellStyle;
    }

    private static String getFormat(String format, String defaultFormat){
        if(format == null || format.isBlank()){
            return defaultFormat;
        }
        return format;
    }
}
