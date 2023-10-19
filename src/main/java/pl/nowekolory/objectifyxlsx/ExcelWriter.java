package pl.nowekolory.objectifyxlsx;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import pl.nowekolory.objectifyxlsx.cell.CellParameters;
import pl.nowekolory.objectifyxlsx.cell.CellStyleCreator;
import pl.nowekolory.objectifyxlsx.header.HeaderCreator;
import pl.nowekolory.objectifyxlsx.row.RowCreator;


import java.util.List;

public class ExcelWriter{

    private final Workbook workbook;
    private final CellStyle dateStyle;
    private final CellStyle dateTimeStyle;

    public ExcelWriter(Workbook workbook){
        this(workbook, CellStyleCreator.createDateCellStyle(workbook, null),
             CellStyleCreator.createDateTimeCellStyle(workbook, null));
    }

    public ExcelWriter(Workbook workbook, CellStyle dateStyle, CellStyle dateTimeStyle){
        this.workbook = workbook;
        this.dateStyle = dateStyle;
        this.dateTimeStyle = dateTimeStyle;
    }

    /**
     * default sheet creator with sheet named by class name
     */
    public void createSheet(List<?> objectsToWrite){
        createSheet(objectsToWrite, null);
    }

    public void createSheet(List<?> objectsToWrite, String name){
        if(objectsToWriteNotExist(objectsToWrite)){
            return;
        }

        var objectsToWriteClazz = objectsToWrite.get(0).getClass();
        var sheetName = getSheetName(objectsToWriteClazz, name);
        var sheet = workbook.createSheet(sheetName);

        HeaderCreator.createHeader(sheet, workbook, objectsToWriteClazz);
        createRows(sheet, objectsToWrite);
    }
    public void resizeColumns(Workbook wb){
        if(wb == null){
            return;
        }
        var numberOfSheets = wb.getNumberOfSheets();
        if(numberOfSheets == 0){
            return;
        }
        for(var i = 0; i < numberOfSheets; i++){
            var sheet = wb.getSheetAt(i);
            var numberOfColumns = getNumberOfColumns(sheet);
            for(var j = 0; j < numberOfColumns; j++){
                if (sheet instanceof SXSSFSheet) {
                    ((SXSSFSheet) sheet).trackColumnForAutoSizing(j);
                }
                sheet.autoSizeColumn(j);
            }
        }
    }
    private short getNumberOfColumns(Sheet sheet){
        var row = sheet.getRow(0);
        if(row == null){
            return getNumberOfColumnsInLongestRow(sheet);
        }
        var numberOfColumns = row.getLastCellNum();
        if(numberOfColumns <= 1){
            return getNumberOfColumnsInLongestRow(sheet);
        }
        return row.getLastCellNum();
    }
    private short getNumberOfColumnsInLongestRow(Sheet sheet){
        var lastRow = sheet.getLastRowNum();
        var numberOfColumns = (short)0;
        for(var i = 0; i < lastRow; i++){
            var row = sheet.getRow(i);
            if(row != null){
                var currentRowColumnsNumber = row.getLastCellNum();
                if(currentRowColumnsNumber > numberOfColumns){
                    numberOfColumns = currentRowColumnsNumber;
                }
            }
        }
        return numberOfColumns;
    }

    private String getSheetName(Class<?> clazz, String name){
        if(name == null || name.isBlank()){
            return clazz.getName();
        }
        return name;
    }

    private void createRows(Sheet sheet, List<?> objectsToWrite){
        var cellParameters = CellParameters.builder()
                .boldFont(false)
                .horizontalAlignment(HorizontalAlignment.RIGHT)
                .roundDouble(true)
                .build();
        var cellStyle = CellStyleCreator.createFormatedCellStyle(workbook, cellParameters.getRoundDouble(), cellParameters.getHorizontalAlignment(), cellParameters.getBoldFont());
        createRows(sheet, objectsToWrite, cellStyle);
    }
    private void createRows(Sheet sheet, List<?> objectsToWrite, CellStyle cellStyle){
        Class<?> objectsToWriteClazz;
        if(objectsToWrite != null && !objectsToWrite.isEmpty()){
            objectsToWriteClazz = objectsToWrite.get(0).getClass();
        }else{
            return;
        }

        var rowCreator = new RowCreator(objectsToWriteClazz, dateStyle,
                                        dateTimeStyle);
        var index = 1;
        for(var objectToWrite : objectsToWrite){
            var row = sheet.createRow(index);
            rowCreator.createRow(row, objectToWrite, cellStyle);
            index++;
        }
    }

    private boolean objectsToWriteNotExist(List<?> objectsToWrite){
        return objectsToWrite == null || objectsToWrite.isEmpty();
    }
}
