package pl.nowekolory.objectifyxlsx.data;

import lombok.AllArgsConstructor;
import lombok.Data;
import pl.nowekolory.objectifyxlsx.header.ReportHeader;

import java.time.LocalDate;

@Data
@AllArgsConstructor
public class CmrPdfTestData{
    @ReportHeader(name = "Lp.", columnWidthPdf = 0.175F)
    private String index;
    @ReportHeader(name = "Kod partnera", inPDF = false)
    private String partnerCode;
    @ReportHeader(name = "Numer zamówienia", columnWidthPdf = 0.6F)
    private String orderNumber;
    @ReportHeader(name = "Zewnętrzny numer zamówienia", columnWidthPdf = 0.9F)
    private String externalOrderNumber;
    @ReportHeader(name = "Kurier", columnWidthPdf = 0.5F)
    private String courier;
    @ReportHeader(name = "AWB", columnWidthPdf = 0.8F)
    private String awb;
    @ReportHeader(name = "Status paczki")
    private String parcelStatus;
    @ReportHeader(name = "Data wysłania", columnWidthPdf = 0.4F)
    private LocalDate sentDate;
    @ReportHeader(name = "Data dostarczenia", inPDF = false)
    private LocalDate deliveredDate;
    @ReportHeader(name = "Odbiorca")
    private String recipient;
    @ReportHeader(name = "Kraj", columnWidthPdf = 0.15F)
    private String country;
    @ReportHeader(name = "Miasto")
    private String city;
    @ReportHeader(name = "Ulica", columnWidthPdf = 2F)
    private String street;
}
