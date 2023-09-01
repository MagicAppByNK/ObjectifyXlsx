package pl.nowekolory.objectifyxlsx;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import pl.nowekolory.objectifyxlsx.cell.CellStyleCreator;
import pl.nowekolory.objectifyxlsx.header.HeaderCreator;
import pl.nowekolory.objectifyxlsx.row.RowCreator;


import java.util.List;

public class ExcelWriter{

    private static final Logger logger = LogManager.getLogger(ExcelWriter.class);

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

    private String getSheetName(Class<?> clazz, String name){
        if(name == null || name.isBlank()){
            return clazz.getName();
        }
        return name;
    }

    private void createRows(Sheet sheet, List<?> objectsToWrite){
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
            rowCreator.createRow(row, objectToWrite);
            index++;
        }
    }

    private boolean objectsToWriteNotExist(List<?> objectsToWrite){
        return objectsToWrite == null || objectsToWrite.isEmpty();
    }
}
