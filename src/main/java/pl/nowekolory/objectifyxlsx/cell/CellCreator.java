package pl.nowekolory.objectifyxlsx.cell;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.time.LocalDate;
import java.time.LocalDateTime;

public class CellCreator{
    private static final Logger logger = LogManager.getLogger(CellCreator.class);

    CellStyle dateTimeStyle;
    CellStyle dateStyle;

    public CellCreator(CellStyle dateTimeStyle, CellStyle dateStyle){
        this.dateStyle = dateStyle;
        this.dateTimeStyle = dateTimeStyle;
    }
    public CellCreator(){
    }

    public void addCell(Row row, Object value, Integer cellIndex){
        if(value == null){
            addEmptyCell(row, cellIndex);
            return;
        }
        if(value instanceof String){
            addCell(row, (String) value, cellIndex);
        }else if(value instanceof Integer){
            addCell(row, (Integer) value, cellIndex);
        }else if(value instanceof Double){
            addCell(row, (Double) value, cellIndex);
        }else if(value instanceof Long){
            addCell(row, (Long) value, cellIndex);
        }else if(value instanceof LocalDateTime){
            addCell(dateTimeStyle, row, (LocalDateTime) value, cellIndex);
        }else if(value instanceof LocalDate){
            addCell(dateStyle, row, (LocalDate) value, cellIndex);
        }else if(value instanceof Boolean){
            addCell(row, (Boolean) value, cellIndex);
        }else if(value instanceof Float){
            addCell(row, (Float) value, cellIndex);
        }else{
            logger.error("No cell creator for given value class: {}", value.getClass());
        }
    }

    private void addEmptyCell(Row row, Integer cellIndex){
        row.createCell(cellIndex).setBlank();
    }
    private void addCell(Row row, String value, Integer cellIndex){
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCell(Row row, Integer value, Integer cellIndex){
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCell(Row row, Double value, Integer cellIndex){
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCell(Row row, Long value, Integer cellIndex){
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCell(Row row, Float value, Integer cellIndex){
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCell(CellStyle cellStyle, Row row, LocalDateTime value, Integer cellIndex){
        var cell = row.createCell(cellIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }

    private void addCell(Row row, Boolean value, Integer cellIndex){
        if(value == null){
            value = false;
        }
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCell(CellStyle cellStyle, Row row, LocalDate value, Integer cellIndex){
        var cell = row.createCell(cellIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }
}
