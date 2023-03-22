package pl.nowekolory.objectifyxlsx.row;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import pl.nowekolory.objectifyxlsx.cell.CellCreator;
import pl.nowekolory.objectifyxlsx.header.ExternalObject;
import pl.nowekolory.objectifyxlsx.header.ReportHeader;

import java.lang.reflect.Field;
import java.util.Arrays;

public class RowCreator{

    private static final Logger logger = LogManager.getLogger(RowCreator.class);

    private final Field[] fields;
    private final CellCreator cellCreator;
    private int cellIndex;

    public RowCreator(Class<?> clazz, CellStyle dateStyle, CellStyle dateTimeStyle){
        this.fields = getReportFields(clazz);
        this.cellCreator = new CellCreator(dateTimeStyle, dateStyle);
    }

    public void createRow(Row row, Object objectToWrite){
        cellIndex = 0;
        createCellsFromFields(fields, row, objectToWrite);
    }

    private void createCellsFromFields(Field[] fields, Row row, Object objectToWrite){
        for(var field : fields){
            field.setAccessible(true);
            try{
                var value = getValue(objectToWrite, field);
                checkForIterableAndCreateCells(field, value, row);
            }catch(IllegalAccessException e){
                logger.error(e.getMessage());
            }
        }
    }

    private Object getValue(Object object, Field field) throws IllegalAccessException{
        if(object == null){
            return null;
        }
        return field.get(object);
    }

    private void checkForIterableAndCreateCells(Field field, Object value, Row row){
        if(Iterable.class.isAssignableFrom(field.getDeclaringClass()) && value != null){
            for(var element : (Iterable<?>) value){
                createCells(field, element, row);
            }
        }else{
            createCells(field, value, row);
        }
    }

    private void createCells(Field field, Object value, Row row){
        if(field.isAnnotationPresent(ExternalObject.class)){
            var externalObject = field.getAnnotation(ExternalObject.class);
            var externalObjectClass = externalObject.className();
            var externalObjectFields = getReportFields(externalObjectClass);
            createCellsFromFields(externalObjectFields, row, value);
        }else{
            createCell(row, value);
        }
    }

    private void createCell(Row row, Object value){
        if(value != null){
            cellCreator.addCell(row, value, cellIndex);
        }
        cellIndex++;
    }

    private Field[] getReportFields(Class<?> clazz){
        var allFields = getAllFields(clazz);
        var reportFields = Arrays.stream(allFields)
                .filter(field -> field.isAnnotationPresent(ReportHeader.class))
                .toArray(Field[]::new);

        if(reportFields.length == 0){
            return allFields;
        }
        return reportFields;
    }

    private Field[] getAllFields(Class<?> clazz){
        return clazz.getDeclaredFields();
    }

}
