package pl.nowekolory.objectifyxlsx.writingTests;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;
import pl.nowekolory.objectifyxlsx.ExcelWriter;
import pl.nowekolory.objectifyxlsx.header.ReportHeader;

import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;

public class DateWritingTest {
    @Test
    public void writeDate(){
        try(var workbook = new SXSSFWorkbook()){
            var excelWriter = new ExcelWriter(workbook);
            var dateList = new ArrayList<TestDateClass>();
            var date = LocalDate.now();
            var dateTime = LocalDateTime.now();
            var testDateClass = new TestDateClass(date,dateTime);
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
        @ReportHeader(name = "DateTime")
        LocalDateTime dateTime;
        public TestDateClass(LocalDate date, LocalDateTime dateTime){
            this.date = date;
            this.dateTime = dateTime;
        }
        public LocalDate getDate(){
            return date;
        }
        public LocalDateTime getDateTime(){
            return dateTime;
        }
    }
}
