package pl.nowekolory.objectifyxlsx;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class ExcelWriterTest {

    @Test
    void testCreateSheetWithPolishCharactersInName() throws Exception {
        try (var workbook = WorkbookFactory.create(false)) {
            var excelWriter = new ExcelWriter(workbook);
            var polishSheetName = "Arkusz Żółty";

            var objectsToWrite = Arrays.asList(
                    new TestObject("Element 1"),
                    new TestObject("Element 2")
            );

            excelWriter.createSheet(objectsToWrite, polishSheetName);

            assertEquals(polishSheetName, workbook.getSheetAt(0).getSheetName(),
                    "Nazwa arkusza powinna zostać poprawnie zakodowana");
        }
    }

    @Test
    void testCreateSheetWithDefaultClassNameWhenNameIsNull() throws Exception {
        try (var workbook = WorkbookFactory.create(false)) {
            ExcelWriter excelWriter = new ExcelWriter(workbook);

            var objectsToWrite = List.of(
                    new TestObject("Element 1")
            );

            excelWriter.createSheet(objectsToWrite, null);

            var expectedSheetName = TestObject.class.getName().substring(0, 31); // nazwy arkuszy excel mają limitację 31 znaków
            assertEquals(expectedSheetName, workbook.getSheetAt(0).getSheetName(),
                    "Nazwą arkusza powinna być skrócona nazwa klasy obiektu do 31 znaków");
        }
    }


    static class TestObject {
        private final String value;

        TestObject(String value) {
            this.value = value;
        }

        public String getValue() {
            return value;
        }
    }
}