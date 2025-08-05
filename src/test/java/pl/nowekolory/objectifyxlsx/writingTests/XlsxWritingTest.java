package pl.nowekolory.objectifyxlsx.writingTests;

import org.junit.jupiter.api.Test;
import pl.nowekolory.objectifyxlsx.ReportFileCreator;
import pl.nowekolory.objectifyxlsx.data.CmrPdfTestData;
import pl.nowekolory.objectifyxlsx.enums.FileType;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;

public class XlsxWritingTest{

    @Test
    public void writeData(){
        final var fileCreator = new ReportFileCreator(FileType.XLSX);
        final var dateList = new ArrayList<CmrPdfTestData>();
        dateList.add(createData("1"));
        dateList.add(createData("2"));
        fileCreator.addDataToFile(dateList, "test");
        final var bytes = fileCreator.getFileBytes();
        try (var stream = new FileOutputStream("src/test/java/test.xlsx")) {
            stream.write(bytes);
        }catch(IOException e){
            throw new RuntimeException(e);
        }
    }

    @Test
    public void addDataTwice(){
        final var fileCreator = new ReportFileCreator(FileType.XLSX);
        final var dateList = new ArrayList<CmrPdfTestData>();
        dateList.add(createData("1"));
        dateList.add(createData("2"));
        fileCreator.addDataToFile(dateList, "test1");
        fileCreator.addDataToFile(dateList, "test2");
        final var bytes = fileCreator.getFileBytes();
        try (var stream = new FileOutputStream("src/test/java/test.xlsx")) {
            stream.write(bytes);
        }catch(IOException e){
            throw new RuntimeException(e);
        }
    }

    @Test
    public void writeManyPagesData(){
        final var fileCreator = new ReportFileCreator(FileType.PDF);
        final var dateList = new ArrayList<CmrPdfTestData>();
        for(var i = 0; i < 200; i++){
            dateList.add(createData(String.valueOf(i)));
        }
        fileCreator.addDataToFile(dateList, "test");
        final var bytes = fileCreator.getFileBytes();
        try (var stream = new FileOutputStream("src/test/java/test.xlsx")) {
            stream.write(bytes);
        }catch(IOException e){
            throw new RuntimeException(e);
        }
    }

    private CmrPdfTestData createData(String index) {
        return new CmrPdfTestData(index, "UHNIKOŃ", "CLN10113285000", "12EX1",
                                  "Sameday", "5EHULN1873000", "DELIVERED", LocalDate.now(), LocalDate.now(),
                                  "Marek Mostowiak", "HU", "Törökszentmiklós,\nSurjány ",
                                  "Almássy út 46-48 [EasyBox #14797]");
    }

}


