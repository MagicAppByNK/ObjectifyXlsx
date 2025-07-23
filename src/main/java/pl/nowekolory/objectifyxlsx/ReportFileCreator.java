package pl.nowekolory.objectifyxlsx;

import com.lowagie.text.Document;
import com.lowagie.text.PageSize;
import lombok.Data;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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

    public ReportFileCreator(FileType fileType) {
        this.fileType = fileType;
        this.outputStream = new ByteArrayOutputStream();
        if (FileType.PDF == fileType) {
            this.pdfDocument = new Document(PageSize.A4.rotate());
            this.workbook = null;
            new ToPdfWriter(outputStream).initPDF(pdfDocument);
        } else {
            this.pdfDocument = null;
            this.workbook = new XSSFWorkbook();
        }
    }

    public void addDataToFile(List<?> objectsToWrite, String sheetName) {
        if (FileType.PDF == fileType) {
            addDataToPdf(objectsToWrite, sheetName);
        } else {
            addDataToExcel(objectsToWrite, sheetName);
        }
    }

    public byte[] getFileBytes() {
        try {
            if (FileType.PDF == fileType) {
                new ToPdfWriter(outputStream).closePdf(pdfDocument);
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

    private void addDataToPdf(List<?> objectsToWrite, String sheetName) {
        final var outputStream = new ByteArrayOutputStream();
        final var pdfWriter = new ToPdfWriter(outputStream);
        pdfWriter.addDataToPDF(objectsToWrite, sheetName, pdfDocument);
    }

}
