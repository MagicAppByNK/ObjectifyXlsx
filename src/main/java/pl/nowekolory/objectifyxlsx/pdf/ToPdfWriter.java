package pl.nowekolory.objectifyxlsx.pdf;

import com.lowagie.text.*;
import com.lowagie.text.Font;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.util.StringUtil;
import pl.nowekolory.objectifyxlsx.header.ExternalObject;
import pl.nowekolory.objectifyxlsx.header.ReportHeader;

import java.awt.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class ToPdfWriter{

    private static final String noDataMessage = "No data to show!";

    private final ByteArrayOutputStream outputStream;

    public ToPdfWriter(ByteArrayOutputStream outputStream){
        this.outputStream = outputStream;
    }

    public void initPDF(Document document) {
        final var writer = PdfWriter.getInstance(document, outputStream);
        writer.setPageEvent(new PageNumberEvent());
        document.open();
        document.setMargins(20, 20, 50, 50);
    }

    public void addDataToPDF(List<?> objectsToWrite, String name, Document document) {
        if(objectsToWrite.isEmpty()){
            addTitlePage(noDataMessage, document);
            return;
        }
        final var objectsToWriteClazz = objectsToWrite.get(0).getClass();
        try {
            if(StringUtil.isNotBlank(name)) {
                addTitlePage(name, document);
            }
            final var headers = getHeaderValues(objectsToWriteClazz);
            final var table = new PdfPTable(headers.size());

            table.setWidthPercentage(100);

            final var columnWidths = getColumnWidths(objectsToWriteClazz, headers.size());
            table.setWidths(columnWidths);

            for (var header : headers) {
                final var headerCell = new PdfPCell(new Phrase(header, new Font(Font.HELVETICA, 8, Font.NORMAL)));
                headerCell.setHorizontalAlignment(Element.ALIGN_LEFT);
                headerCell.setBackgroundColor(Color.LIGHT_GRAY);
                headerCell.setBorder(Rectangle.NO_BORDER);
                table.addCell(headerCell);
            }

            for (var obj : objectsToWrite) {
                final var fields = obj.getClass().getDeclaredFields();
                for (var field : fields) {
                    field.setAccessible(true);
                    try {
                        if (field.isAnnotationPresent(ReportHeader.class)
                                && !field.getAnnotation(ReportHeader.class).inPDF()) {
                            continue;
                        }
                        table.addCell(getCell(obj, field));
                    } catch (IllegalAccessException e) {
                        table.addCell("");
                    }
                }
            }
            document.add(table);
            document.newPage();
        } catch (DocumentException e) {
            throw new RuntimeException("Błąd podczas tworzenia pliku PDF: " + e.getMessage(), e);
        }
    }

    public void closePdf(Document document) {
        document.close();
    }

    private static PdfPCell getCell(Object obj, Field field) throws IllegalAccessException {
        final var value = field.get(obj);
        final var cell = new PdfPCell(new Phrase(value != null ? value.toString() : "",
                                                new Font(Font.HELVETICA, 8, Font.NORMAL)));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorder(Rectangle.NO_BORDER);
        if (field.isAnnotationPresent(ReportHeader.class) && !field.getAnnotation(ReportHeader.class).multiline()) {
            cell.setNoWrap(true);
        }
        return cell;
    }

    private static float[] getColumnWidths(Class<?> headers, int size) {
        final var headerClassFields = headers.getDeclaredFields();
        float[] columnWidths = new float[size];
        var position = 0;
        for (var field : headerClassFields) {
            if (field.isAnnotationPresent(ReportHeader.class)
                    && field.getAnnotation(ReportHeader.class).inPDF()){
                    columnWidths[position] = field.getAnnotation(ReportHeader.class).columnWidthPdf();
                    position++;
            }
        }
        return columnWidths;
    }

    private static List<String> getHeaderValues(Class<?> clazz) {
        final var headerClassFields = clazz.getDeclaredFields();
        var headerList = new ArrayList<String>();
        for (var field : headerClassFields) {
            if (field.isAnnotationPresent(ExternalObject.class)) {
                var externalObject = field.getAnnotation(ExternalObject.class);
                headerList.addAll(getHeaderValues(externalObject.className()));
            } else if(field.isAnnotationPresent(ReportHeader.class)) {
                if (field.getAnnotation(ReportHeader.class).inPDF()) {
                    headerList.add(field.getAnnotation(ReportHeader.class).name());
                }
            }
        }
        if (headerList.isEmpty()) {
            headerList = new ArrayList<>(getDefaultHeadersValues(clazz));
        }
        return headerList;
    }

    private static void addTitlePage(String titleValue, Document document) {
        final var titleFont = new Font(Font.HELVETICA, 24, Font.BOLD);
        final var title = new Paragraph(titleValue, titleFont);
        title.setAlignment(Element.ALIGN_CENTER);
        document.add(title);
        document.newPage();
    }

    private static List<String> getDefaultHeadersValues(Class<?> clazz){
        final var declaredFields = clazz.getDeclaredFields();
        return Arrays.stream(declaredFields).map(Field::getName).collect(Collectors.toList());
    }
}
