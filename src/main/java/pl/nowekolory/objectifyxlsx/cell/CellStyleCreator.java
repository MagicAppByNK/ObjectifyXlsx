package pl.nowekolory.objectifyxlsx.cell;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.DateFormatConverter;

import java.util.Locale;

public class CellStyleCreator{

    /**
     * Data format should be compliant with the Excel standard
     */
    private static final String defaultDateFormat = "d.M.yy";
    private static final String defaultDateTimeFormat = "d.M.yy h:mm";

    public static CellStyle createDefaultCellStyle(Workbook workbook){
        return workbook.createCellStyle();
    }
    public static CellStyle createFormatedCellStyle(Workbook workbook, Boolean roundDouble, HorizontalAlignment horizontalAlignment, Boolean boldFont){
        var cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(horizontalAlignment);
        if(boldFont){
            var font = workbook.createFont();
            font.setBold(true);
            cellStyle.setFont(font);
        }
        if(roundDouble){
            var format = workbook.getCreationHelper().createDataFormat().getFormat("0.00");
            cellStyle.setDataFormat(format);
        }
        return cellStyle;
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
        var stringFormat = getFormat(format, DateFormatConverter.convert(new Locale("pl"), defaultDateFormat));
        var dateFormat = workbook.getCreationHelper().createDataFormat().getFormat(stringFormat);
        dateCellStyle.setDataFormat(dateFormat);
        return dateCellStyle;
    }

    public static CellStyle createHeaderCellStyle(Workbook workbook){
        var cellStyle = workbook.createCellStyle();
        var font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        return cellStyle;
    }

    private static String getFormat(String format, String defaultFormat){
        if(format == null || format.isBlank()){
            return defaultFormat;
        }
        return format;
    }
}
