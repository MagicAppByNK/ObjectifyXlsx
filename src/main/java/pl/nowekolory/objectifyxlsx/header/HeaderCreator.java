package pl.nowekolory.objectifyxlsx.header;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import pl.nowekolory.objectifyxlsx.cell.CellCreator;
import pl.nowekolory.objectifyxlsx.cell.CellStyleCreator;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class HeaderCreator{

    public static void createHeader(Sheet sheet, Workbook workbook, Class<?> objectsToWriteClazz){
        var headerValues = getHeaderValues(objectsToWriteClazz);
        createHeader(sheet, workbook, headerValues);
    }

    public static void createHeader(Sheet sheet, Workbook workbook, List<String> headersValues){
        var cellCreator = new CellCreator();
        var row = sheet.createRow(0);
        var cellIndex = 0;
        var headerCellStyle = CellStyleCreator.createHeaderCellStyle(workbook);
        for(var title : headersValues){
            cellCreator.addCell(row, title, cellIndex);
            cellIndex++;
        }
        for(var i = 0; i < headersValues.size(); i++){
            row.getCell(i).setCellStyle(headerCellStyle);
        }
    }

    private static List<String> getHeaderValues(Class<?> clazz){
        var headerClassFields = clazz.getDeclaredFields();
        var headerList = new ArrayList<String>();
        for(var field : headerClassFields){
            if(field.isAnnotationPresent(ReportHeader.class)){
                headerList.add(field.getAnnotation(ReportHeader.class).name());
            }
        }
        if(headerList.isEmpty()){
            headerList = new ArrayList<>(getDefaultHeadersValues(clazz));
        }
        return headerList;
    }

    private static List<String> getDefaultHeadersValues(Class<?> clazz){
        var declaredFields = clazz.getDeclaredFields();
        return Arrays.stream(declaredFields).map(Field::getName).collect(Collectors.toList());
    }
}
