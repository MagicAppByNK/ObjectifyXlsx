package pl.nowekolory.objectifyxlsx;


import lombok.Data;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openpdf.text.Document;
import org.openpdf.text.PageSize;
import pl.nowekolory.objectifyxlsx.enums.FileType;
import pl.nowekolory.objectifyxlsx.pdf.ToPdfWriter;

import java.io.IOException;
import java.util.List;

@Data
public class ReportFileCreator{

    private static final Logger logger = LogManager.getLogger(ReportFileCreator.class);

    private Document pdfDocument;
    private Workbook workbook ;
    private FileType fileType;
    private ByteArrayOutputStream outputStream;
    private int fontSize;
    private boolean newPage = true;

    public ReportFileCreator(FileType fileType) {
        this(fileType, 6);
    }

    public ReportFileCreator(FileType fileType, int fontSize) {
        this.fileType = fileType;
        this.outputStream = new ByteArrayOutputStream();
        if (FileType.PDF == fileType) {
            this.pdfDocument = new Document(PageSize.A4.rotate());
            this.workbook = null;
            this.fontSize = fontSize;
            new ToPdfWriter(outputStream, fontSize).initPDF(pdfDocument);
        } else {
            this.pdfDocument = null;
            this.workbook = new XSSFWorkbook();
        }
    }

    public void addDataToFile(List<?> objectsToWrite, String sheetName) {
        addDataToFile(objectsToWrite, sheetName, false);
    }

    public void addDataToFile(List<?> objectsToWrite, String sheetName, boolean tableHeaders) {
        if (FileType.PDF == fileType) {
            addDataToPdf(objectsToWrite, sheetName, tableHeaders);
        } else {
            addDataToExcel(objectsToWrite, sheetName);
        }
    }

    public void newPagePdf() {
        if (FileType.XLSX == fileType) {
            return;
        }
        final var outputStream = new ByteArrayOutputStream();
        new ToPdfWriter(outputStream, fontSize).newPagePdf(pdfDocument);
    }

    public void addHeaderToPDF(String headerText, int fontSize) {
        if (FileType.XLSX == fileType) {
            return;
        }
        final var outputStream = new ByteArrayOutputStream();
        new ToPdfWriter(outputStream, fontSize).addHeaderToPDF(headerText, pdfDocument);
    }

    public void addCompaniesToPDF(String leftCompany, String rightCompany, int fontSize) {
        if (FileType.XLSX == fileType) {
            return;
        }
        final var outputStream = new ByteArrayOutputStream();
        new ToPdfWriter(outputStream, fontSize).addCompaniesToPDF(leftCompany, rightCompany, pdfDocument);
    }

    public void addIssueDateToPDF(String issueDateText, int fontSize) {
        if (FileType.XLSX == fileType) {
            return;
        }
        final var outputStream = new ByteArrayOutputStream();
        new ToPdfWriter(outputStream, fontSize).addIssueDateToPDF(issueDateText, pdfDocument);
    }

    public void addImageToPDF(byte[] imageBytes, int width, int height, int alignment, int margin) {
        if (FileType.XLSX == fileType) {
            return;
        }
        final var outputStream = new ByteArrayOutputStream();
        new ToPdfWriter(outputStream, fontSize).addImageToPDF(imageBytes, width, height, alignment, margin, pdfDocument);
    }

    public byte[] getFileBytes() {
        try {
            if (FileType.PDF == fileType) {
                new ToPdfWriter(outputStream, fontSize).closePdf(pdfDocument);
            } else {
                workbook.write(outputStream);
            }
            outputStream.flush();
            outputStream.close();
            return outputStream.toByteArray();
        } catch (IOException e) {
            logger.error("Error during getting file bytes", e);
            return null;
        }
    }

    private void addDataToExcel(List<?> objectsToWrite, String sheetName) {
        final var excelWriter = new ExcelWriter(workbook);
        excelWriter.createSheet(objectsToWrite, sheetName);
        excelWriter.resizeColumns();
    }

    private void addDataToPdf(List<?> objectsToWrite, String sheetName, boolean tableHeaders) {
        final var outputStream = new ByteArrayOutputStream();
        new ToPdfWriter(outputStream, fontSize).addDataToPDF(objectsToWrite, sheetName, pdfDocument, tableHeaders);
    }

}
