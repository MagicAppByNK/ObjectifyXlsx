package pl.nowekolory.objectifyxlsx.row;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import pl.nowekolory.objectifyxlsx.cell.CellCreator;
import pl.nowekolory.objectifyxlsx.header.ReportHeader;

import java.lang.reflect.Field;
import java.util.Arrays;

public class RowCreator{

    private static final Logger logger = LogManager.getLogger(RowCreator.class);

    private final Field[] fields;
    private final CellStyle dateStyle;
    private final CellStyle dateTimeStyle;

    public RowCreator(Class<?> clazz, CellStyle dateStyle, CellStyle dateTimeStyle){
        this.dateStyle = dateStyle;
        this.dateTimeStyle = dateTimeStyle;
        var fields =  getReportFields(clazz);
        if(fields.length == 0){
            this.fields = getAllFields(clazz);
        }else {
            this.fields = fields;
        }

    }

    public void createRow(Row row, Object objectToWrite){
        var cellCreator = new CellCreator();
        var cellIndex = 0;
        for(var field : fields){
            field.setAccessible(true);
            try{
                var value = field.get(objectToWrite);
                if(value != null){
                    cellCreator.addCell(row, value, cellIndex);
                }
                cellIndex++;
            }catch(IllegalAccessException e){
                logger.error(e.getMessage());
            }
        }
    }


    private Field[] getReportFields(Class<?> clazz){
        var fields = clazz.getDeclaredFields();
        return Arrays.stream(fields).filter(field -> field.isAnnotationPresent(
                ReportHeader.class)).toArray(Field[]::new);
    }

    private Field[] getAllFields(Class<?> clazz){
        return clazz.getDeclaredFields();
    }

}
