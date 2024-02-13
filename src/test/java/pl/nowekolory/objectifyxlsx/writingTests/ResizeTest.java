package pl.nowekolory.objectifyxlsx.writingTests;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import pl.nowekolory.objectifyxlsx.ExcelWriter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ResizeTest {
    @Test
    public void calculateIterateTime(){
        try{
            var file = new FileInputStream(new File("src/test/resources/iterateTimeTest.xlsx"));
            var workbook = new XSSFWorkbook(file);
            var startMilliseconds = System.currentTimeMillis();
            var excelWriter = new ExcelWriter(workbook);
            excelWriter.resizeColumns();
            var out = new FileOutputStream("src/test/resources/dateTest.xlsx");
            workbook.write(out);
            var endMilliseconds = System.currentTimeMillis();
            var duration = endMilliseconds - startMilliseconds;
            System.out.println("Duration: " + duration);

        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
