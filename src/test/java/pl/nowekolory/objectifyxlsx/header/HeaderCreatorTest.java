package pl.nowekolory.objectifyxlsx.header;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class HeaderCreatorTest {

    @Test
    void testCreateHeaderWithPolishCharacters() throws Exception {
        try (var workbook = WorkbookFactory.create(false)) {
            var sheet = workbook.createSheet("TestSheet");

            var headersWithPolishCharacters = Arrays.asList("Nagłówek 1", "Kolumna Łącząca", "Wartość Żródłowa");

            HeaderCreator.createHeader(sheet, workbook, headersWithPolishCharacters);

            for (var i = 0; i < headersWithPolishCharacters.size(); i++) {
                var expectedHeader = headersWithPolishCharacters.get(i);
                var actualHeader = sheet.getRow(0).getCell(i).getStringCellValue();
                assertEquals(expectedHeader, actualHeader,
                        "Sprawdzenie, czy nagłówek został poprawnie zakodowany dla indeksu: " + i);
            }
        }
    }

    @Test
    void shouldHandleNullTitleInHeadersList() throws Exception {
        try (var workbook = WorkbookFactory.create(false)) {
            var sheet = workbook.createSheet();
            List<String> headersValues = new ArrayList<>();
            headersValues.add("Column1");
            headersValues.add(null);
            headersValues.add("Column3");

            assertDoesNotThrow(() ->
                            HeaderCreator.createHeader(sheet, workbook, headersValues),
                    "Metoda powinna obsłużyć przypadek, gdy jeden z nagłówków jest null");
        }
    }


}