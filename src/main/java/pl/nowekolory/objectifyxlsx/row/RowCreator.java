package pl.nowekolory.objectifyxlsx.row;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import pl.nowekolory.objectifyxlsx.cell.CellCreator;
import pl.nowekolory.objectifyxlsx.cell.CellParameters;
import pl.nowekolory.objectifyxlsx.cell.CellStyleCreator;
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
        var cellParameters = CellParameters.builder()
                .boldFont(false)
                .horizontalAlignment(HorizontalAlignment.RIGHT)
                .roundDouble(true)
                .build();
        createRowWithCellParameters(row, objectToWrite, cellParameters);
    }
    public void createRowWithCellParameters(Row row, Object objectToWrite, CellParameters cellParameters){
        cellIndex = 0;
        createCellsFromFields(fields, row, objectToWrite, cellParameters);
    }

    private void createCellsFromFields(Field[] fields, Row row, Object objectToWrite, CellParameters cellParameters){
        for(var field : fields){
            field.setAccessible(true);
            try{
                var value = getValue(objectToWrite, field);
                checkForIterableAndCreateCells(field, value, row, cellParameters);
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

    private void checkForIterableAndCreateCells(Field field, Object value, Row row, CellParameters cellParameters){
        if(Iterable.class.isAssignableFrom(field.getDeclaringClass()) && value != null){
            for(var element : (Iterable<?>) value){
                createCells(field, element, row, cellParameters);
            }
        }else{
            createCells(field, value, row, cellParameters);
        }
    }

    private void createCells(Field field, Object value, Row row, CellParameters cellParameters){
        if(field.isAnnotationPresent(ExternalObject.class)){
            var externalObject = field.getAnnotation(ExternalObject.class);
            var externalObjectClass = externalObject.className();
            var externalObjectFields = getReportFields(externalObjectClass);
            createCellsFromFields(externalObjectFields, row, value,cellParameters);
        }else{
            createCell(row, value, cellParameters);
        }
    }

    private void createCell(Row row, Object value, CellParameters cellParameters){
        if(value != null){
            var cellStyle = CellStyleCreator.createFormatedCellStyle(row.getSheet().getWorkbook(), cellParameters.getRoundDouble(), cellParameters.getHorizontalAlignment(), cellParameters.getBoldFont());
            cellCreator.addCell(row, value, cellIndex, cellStyle);
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
