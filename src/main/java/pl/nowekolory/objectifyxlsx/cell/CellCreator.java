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
            addCellWithString(row, (String) value, cellIndex);
        }else if(value instanceof Integer){
            addCellWithInteger(row, (Integer) value, cellIndex);
        }else if(value instanceof Double){
            addCellWithDouble(row, (Double) value, cellIndex);
        }else if(value instanceof LocalDateTime){
            addCellWithLocalDateTime(dateTimeStyle, row, (LocalDateTime) value, cellIndex);
        }else if(value instanceof LocalDate){
            addCellWithLocalDate(dateStyle, row, (LocalDate) value, cellIndex);
        }else if(value instanceof Boolean){
            addCellWithBoolean(row, (Boolean) value, cellIndex);
        }else{
            logger.error("No cell creator for given value class: {}", value.getClass());
        }
    }

    private void addEmptyCell(Row row, Integer cellIndex){
        row.createCell(cellIndex).setBlank();
    }
    private void addCellWithString(Row row, String value, Integer cellIndex){
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCellWithInteger(Row row, Integer value, Integer cellIndex){
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCellWithDouble(Row row, Double value, Integer cellIndex){
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCellWithLocalDateTime(CellStyle cellStyle, Row row, LocalDateTime value, Integer cellIndex){
        var cell = row.createCell(cellIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }

    private void addCellWithBoolean(Row row, Boolean value, Integer cellIndex){
        if(value == null){
            value = false;
        }
        row.createCell(cellIndex).setCellValue(value);
    }

    private void addCellWithLocalDate(CellStyle cellStyle, Row row, LocalDate value, Integer cellIndex){
        var cell = row.createCell(cellIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }
}
