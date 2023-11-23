package pl.nowekolory.objectifyxlsx.writingTests;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;
import pl.nowekolory.objectifyxlsx.ExcelWriter;
import pl.nowekolory.objectifyxlsx.header.ReportHeader;

import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.ArrayList;

public class DateWritingTest {
    @Test
    public void writeDate(){
        try(var workbook = new SXSSFWorkbook()){
            var excelWriter = new ExcelWriter(workbook);
            var dateList = new ArrayList<TestDateClass>();
            var date = LocalDate.of(2012,9,12);
            var testDateClass = new TestDateClass(date);
            dateList.add(testDateClass);
            excelWriter.createSheet(dateList,"test");
            var out = new FileOutputStream("src/test/resources/dateTest.xlsx");
            excelWriter.resizeColumns();
            workbook.write(out);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    class TestDateClass{
        @ReportHeader(name = "Date")
        LocalDate date;
        public TestDateClass(LocalDate date){
            this.date = date;
        }
        public LocalDate getDate(){
            return date;
        }
    }
}
