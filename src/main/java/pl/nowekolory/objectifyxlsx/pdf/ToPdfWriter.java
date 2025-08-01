package pl.nowekolory.objectifyxlsx.pdf;

import com.lowagie.text.*;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
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
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class ToPdfWriter{

    private static final String noDataMessage = "No data to show!";

    private final ByteArrayOutputStream outputStream;
    private final int fontSize;

    public ToPdfWriter(ByteArrayOutputStream outputStream, int fontSize) {
        this.outputStream = outputStream;
        this.fontSize = fontSize;
    }

    public void initPDF(Document document) {
        final var writer = PdfWriter.getInstance(document, outputStream);
        writer.setPageEvent(new PageNumberEvent());
        document.open();
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
            table.setSpacingBefore(0);

            final var columnWidths = getColumnWidths(objectsToWriteClazz, headers.size());
            table.setWidths(columnWidths);

            for (var header : headers) {
                final var headerCell = new PdfPCell(new Phrase(header, new Font(Font.HELVETICA, fontSize, Font.NORMAL)));
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
        } catch (DocumentException e) {
            throw new RuntimeException("Error while creating PDF file: " + e.getMessage(), e);
        }
    }

    public void newPagePdf(Document document) {
        document.newPage();
    }

    public void addHeaderToPDF(String headerText, Document document) {
        final var font = new Font(Font.HELVETICA, fontSize, Font.NORMAL);
        final var table = new PdfPTable(1);
        table.setWidthPercentage(100);

        final var cell = new PdfPCell(new Phrase(headerText, font));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBackgroundColor(Color.LIGHT_GRAY);
        cell.setBorder(Rectangle.NO_BORDER);

        table.addCell(cell);
        table.setSpacingBefore(5);
        table.setSpacingAfter(5);
        document.add(table);
    }

    public void addCompaniesToPDF(String leftCompany, String rightCompany, Document document) {
        final var font = new Font(Font.HELVETICA, fontSize, Font.NORMAL);
        final var table = new PdfPTable(4);
        table.setWidthPercentage(100);

        final var leftCell = new PdfPCell(new Phrase(leftCompany, font));
        leftCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        leftCell.setBorder(Rectangle.NO_BORDER);
        leftCell.setLeading(fontSize * 1.2f, 0);
        table.addCell(leftCell);

        final var placeholderCell = new PdfPCell(new Phrase("", font));
        placeholderCell.setBorder(Rectangle.NO_BORDER);
        table.addCell(placeholderCell);
        table.addCell(placeholderCell);

        final var rightCell = new PdfPCell(new Phrase(rightCompany, font));
        rightCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        rightCell.setBorder(Rectangle.NO_BORDER);
        rightCell.setLeading(fontSize * 1.2f, 0);
        table.addCell(rightCell);

        table.setSpacingBefore(5);
        table.setSpacingAfter(5);
        document.add(table);
    }

    public void addIssueDateToPDF(String issueDateText, Document document) {
        final var font = new Font(Font.HELVETICA, fontSize, Font.NORMAL);
        final var paragraph = new Paragraph(issueDateText, font);
        paragraph.setAlignment(Element.ALIGN_RIGHT);
        paragraph.setSpacingBefore(5);
        paragraph.setSpacingAfter(5);
        document.add(paragraph);
    }

    public void addImageToPDF(byte[] imageBytes, int width, int height, int alignment, int margin, Document document) {
        if (imageBytes == null || imageBytes.length == 0) {
            return;
        }
        try {
            final var image = Image.getInstance(imageBytes);
            image.scaleAbsolute(width, height);
            image.setAlignment(alignment);
            if (alignment == 2) {
                float pageWidth = document.getPageSize().getWidth();
                image.setAbsolutePosition(pageWidth - width - margin, image.getAbsoluteY());
            }
            if (alignment == 0) {
                image.setAbsolutePosition(margin, image.getAbsoluteY());
            }

            final var imageParagraph = new Paragraph();
            imageParagraph.add(image);
            imageParagraph.setSpacingBefore(20);

            document.add(imageParagraph);
        } catch (Exception e) {
            throw new RuntimeException("Error while loading image from file:" + e.getMessage());
        }
    }


    public void closePdf(Document document) {
        document.close();
    }

    private PdfPCell getCell(Object obj, Field field) throws IllegalAccessException {
        final var value = field.get(obj);
        final var cell = new PdfPCell(new Phrase(getObjectStringValue(value),
                                                new Font(Font.HELVETICA, fontSize, Font.NORMAL)));
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        cell.setBorder(Rectangle.NO_BORDER);
        if (field.isAnnotationPresent(ReportHeader.class) && !field.getAnnotation(ReportHeader.class).multiline()) {
            cell.setNoWrap(true);
        }
        return cell;
    }

    private static String getObjectStringValue(Object value) {
        if (value == null) {
            return "";
        }
        if (value instanceof LocalDate) {
            return ((LocalDate) value).format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
        }
        return value.toString();
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
        final var font = new Font(Font.HELVETICA, 24, Font.NORMAL);
        final var table = new PdfPTable(1);
        table.setWidthPercentage(100);

        final var cell = new PdfPCell(new Phrase(titleValue, font));
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBackgroundColor(Color.LIGHT_GRAY);
        cell.setBorder(Rectangle.NO_BORDER);

        table.addCell(cell);
        table.setSpacingBefore(5);
        table.setSpacingAfter(5);
        document.add(table);
    }

    private static List<String> getDefaultHeadersValues(Class<?> clazz){
        final var declaredFields = clazz.getDeclaredFields();
        return Arrays.stream(declaredFields).map(Field::getName).collect(Collectors.toList());
    }
}
