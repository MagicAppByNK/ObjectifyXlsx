package pl.nowekolory.objectifyxlsx.cell;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.time.LocalDate;
import java.time.LocalDateTime;

public class CellCreator {
    private static final Logger logger = LogManager.getLogger(CellCreator.class);

    CellStyle dateTimeStyle;
    CellStyle dateStyle;


    public CellCreator(CellStyle dateTimeStyle, CellStyle dateStyle) {
        this.dateStyle = dateStyle;
        this.dateTimeStyle = dateTimeStyle;

    }

    public CellCreator() {
    }

    public void addFormattedCell(Object value, Row row, Integer cellIndex, CellParameters cellParameters) {
        var cellStyle = CellStyleCreator.createFormatedCellStyle(row.getSheet().getWorkbook(), cellParameters.getRoundDouble(), cellParameters.getHorizontalAlignment(), cellParameters.getBoldFont());
        addCell(row, value, cellIndex, cellStyle);
    }

    public void addCell(Row row, Object value, Integer cellIndex) {
        var workbook = row.getSheet().getWorkbook();
        var cellStyle = CellStyleCreator.createDefaultCellStyle(workbook);
        addCell(row, value, cellIndex, cellStyle);
    }

    public void addCell(Row row, Object value, Integer cellIndex, CellStyle cellStyle) {
        if (value == null) {
            addEmptyCell(row, cellIndex, cellStyle);
            return;
        }
        if (value instanceof String) {
            addCell(row, (String) value, cellIndex, cellStyle);
        } else if (value instanceof Integer) {
            addCell(row, (Integer) value, cellIndex, cellStyle);
        } else if (value instanceof Double) {
            addCell(row, (Double) value, cellIndex, cellStyle);
        } else if (value instanceof Long) {
            addCell(row, (Long) value, cellIndex, cellStyle);
        } else if (value instanceof LocalDateTime) {
            addCell(dateTimeStyle, row, (LocalDateTime) value, cellIndex);
        } else if (value instanceof LocalDate) {
            addCell(dateStyle, row, (LocalDate) value, cellIndex);
        } else if (value instanceof Boolean) {
            addCell(row, (Boolean) value, cellIndex, cellStyle);
        } else if (value instanceof Float) {
            addCell(row, (Float) value, cellIndex, cellStyle);
        } else {
            logger.error("No cell creator for given value class: {}", value.getClass());
        }
    }

    private void addEmptyCell(Row row, Integer cellIndex, CellStyle cellStyle) {
        row.createCell(cellIndex).setBlank();
        row.getCell(cellIndex).setCellStyle(cellStyle);
    }

    private void addCell(Row row, String value, Integer cellIndex, CellStyle cellStyle) {
        row.createCell(cellIndex).setCellValue(value);
        row.getCell(cellIndex).setCellStyle(cellStyle);
    }

    private void addCell(Row row, Integer value, Integer cellIndex, CellStyle cellStyle) {
        row.createCell(cellIndex).setCellValue(value);
        row.getCell(cellIndex).setCellStyle(cellStyle);
    }

    private void addCell(Row row, Double value, Integer cellIndex, CellStyle cellStyle) {
        if (value == 0.0) {
            addEmptyCell(row, cellIndex, cellStyle);
        } else {
            var cell = row.createCell(cellIndex);
            cell.setCellValue(value);
            cell.setCellStyle(cellStyle);
        }
    }

    private void addCell(Row row, Long value, Integer cellIndex, CellStyle cellStyle) {
        row.createCell(cellIndex).setCellValue(value);
        row.getCell(cellIndex).setCellStyle(cellStyle);
    }

    private void addCell(Row row, Float value, Integer cellIndex, CellStyle cellStyle) {
        if (value == 0.0) {
            addEmptyCell(row, cellIndex, cellStyle);
        } else {
            row.createCell(cellIndex).setCellValue(value);
            row.getCell(cellIndex).setCellStyle(cellStyle);
        }
    }

    private void addCell(CellStyle cellStyle, Row row, LocalDateTime value, Integer cellIndex) {
        var cell = row.createCell(cellIndex);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }

    private void addCell(Row row, Boolean value, Integer cellIndex, CellStyle cellStyle) {
        if (value == null) {
            value = false;
        }
        row.createCell(cellIndex).setCellValue(value);
        row.getCell(cellIndex).setCellStyle(cellStyle);
    }

    private void addCell(CellStyle cellStyle, Row row, LocalDate value, Integer cellIndex) {
        var cell = row.createCell(cellIndex, CellType.NUMERIC);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }
}
