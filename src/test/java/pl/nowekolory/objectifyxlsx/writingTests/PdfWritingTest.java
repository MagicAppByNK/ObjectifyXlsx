package pl.nowekolory.objectifyxlsx.writingTests;

import org.junit.jupiter.api.Test;
import org.openpdf.text.Element;
import pl.nowekolory.objectifyxlsx.ReportFileCreator;
import pl.nowekolory.objectifyxlsx.data.CmrPdfTestData;
import pl.nowekolory.objectifyxlsx.enums.FileType;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.ArrayList;

public class PdfWritingTest{

    @Test
    public void writeData(){
        final var fileCreator = new ReportFileCreator(FileType.PDF, 5);
        final var dateList = new ArrayList<CmrPdfTestData>();
        dateList.add(createData("1"));
        dateList.add(createData("2"));
        fileCreator.addIssueDateToPDF("Data wystawienia: 02.07.2025", 8);
        fileCreator.addHeaderToPDF("Zestawienie przesyłek zrealizowanych w ramach dostawy wewnątrzwspólnotowej w okresie od 01.06.2025 do 30.06.2025\n", 8);
        final var partner = "Zleceniodawca\n UNIKOŃ MOSTY GROUP SPÓŁKA Z OGRANICZONĄ ODPOWIEDZIALNOŚCIĄ \n 04-040 Warszawa, Polska\n ul. Rembrantów 9999/99\n NIP: 1234567890";
        final var nkCompany = "Firma logistyczna\nNowe Paczki Nazwa, Nazwa Sp.K.\nRembrantów 12b\n04-040 Warszawa, Polska\nNIP 1234567891\n";
        fileCreator.addCompaniesToPDF(partner, nkCompany, 8);
        fileCreator.addDataToFile(dateList, "");
        final var img = loadImageFromFile("src/test/java/pl/nowekolory/objectifyxlsx/data/pepe.jpg");
        fileCreator.addImageToPDF(img, 150, 85, Element.ALIGN_RIGHT, 100);
        final var bytes = fileCreator.getFileBytes();
        try (var stream = new FileOutputStream("src/test/java/test.pdf")) {
            stream.write(bytes);
        } catch(IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    public void writeEmptyData(){
        final var fileCreator = new ReportFileCreator(FileType.PDF);
        final var dateList = new ArrayList<CmrPdfTestData>();
        fileCreator.addDataToFile(dateList, "test");
        final var bytes = fileCreator.getFileBytes();
        try (var stream = new FileOutputStream("src/test/java/test.pdf")) {
            stream.write(bytes);
        }catch(IOException e){
            throw new RuntimeException(e);
        }
    }

    @Test
    public void addDataTwice(){
        final var fileCreator = new ReportFileCreator(FileType.PDF);
        final var dateList = new ArrayList<CmrPdfTestData>();
        dateList.add(createData("1"));
        dateList.add(createData("2"));
        fileCreator.addDataToFile(dateList, "test1");
        fileCreator.newPagePdf();
        fileCreator.addDataToFile(dateList, "test2");
        fileCreator.newPagePdf();
        fileCreator.addDataToFile(dateList, "test3");
        final var bytes = fileCreator.getFileBytes();
        try (var stream = new FileOutputStream("src/test/java/test.pdf")) {
            stream.write(bytes);
        }catch(IOException e){
            throw new RuntimeException(e);
        }
    }

    @Test
    public void writeManyPagesData(){
        final var fileCreator = new ReportFileCreator(FileType.PDF, 5);
        fileCreator.addIssueDateToPDF("Data wystawienia: 02.07.2025", 8);
        fileCreator.addHeaderToPDF("Zestawienie przesyłek zrealizowanych w ramach dostawy wewnątrzwspólnotowej w okresie od 01.06.2025 do 30.06.2025\n", 8);
        final var partner = "Zleceniodawca\n UNIKOŃ MOSTY GROUP SPÓŁKA Z OGRANICZONĄ ODPOWIEDZIALNOŚCIĄ \n 04-040 Warszawa, Polska\n ul. Rembrantów 9999/99\n NIP: 1234567890";
        final var nkCompany = "Firma logistyczna\nNowe Paczki Nazwa, Nazwa Sp.K.\nRembrantów 12b\n04-040 Warszawa, Polska\nNIP 1234567891\n";
        fileCreator.addCompaniesToPDF(partner, nkCompany, 8);
        final var dateList = new ArrayList<CmrPdfTestData>();
        for(var i = 1; i < 101; i++){
            dateList.add(createData(String.valueOf(i)));
        }
        fileCreator.addDataToFile(dateList, "", true);
        final var img = loadImageFromFile("src/test/java/pl/nowekolory/objectifyxlsx/data/pepe.jpg");
        fileCreator.addImageToPDF(img, 150, 85, Element.ALIGN_RIGHT, 100);
        final var bytes = fileCreator.getFileBytes();
        try (var stream = new FileOutputStream("src/test/java/test.pdf")) {
            stream.write(bytes);
        }catch(IOException e){
            throw new RuntimeException(e);
        }
    }

    private CmrPdfTestData createData(String index) {
        return new CmrPdfTestData(index, "SQ1", "CV911431320000", "WZ/00001/ISKROW/01/2025",
                                  "Urgent Cargus", "7000074455105001B0000", "Zwrot przyjęty w oddziale Bielsko-Biała",
                                  LocalDate.now(), LocalDate.now(),
                                  "Marcel - Wenecjowa Paczkowa", "RO", "Sighetu Marmatiei, Maramures",
                                  "Str Calea Buc 999 Lic Mihalache Str Calea Buc 9999 Li Str Calea Buc 999 Li");
    }

    private byte[] loadImageFromFile(String path) {
        try {
            final var imagePath = Paths.get(path);
            return Files.readAllBytes(imagePath);
        } catch (IOException e) {
            throw new RuntimeException("Błąd podczas wczytywania obrazu z pliku: " + e.getMessage(), e);
        }
    }


}


