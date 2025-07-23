package pl.nowekolory.objectifyxlsx.data;

import lombok.AllArgsConstructor;
import lombok.Data;
import pl.nowekolory.objectifyxlsx.header.ReportHeader;

import java.time.LocalDate;

@Data
@AllArgsConstructor
public class CmrPdfTestData{
    @ReportHeader(name = "Lp.", columnWidthPdf = 0.3F)
    private String index;
    @ReportHeader(name = "Kod partnera", columnWidthPdf = 0.8F)
    private String partnerCode;
    @ReportHeader(name = "Numer zamówienia")
    private String orderNumber;
    @ReportHeader(name = "Zewnętrzny numer zamówienia", inPDF = false)
    private String externalOrderNumber;
    @ReportHeader(name = "Kurier")
    private String courier;
    @ReportHeader(name = "AWB")
    private String awb;
    @ReportHeader(name = "Status paczki", inPDF = false)
    private String parcelStatus;
    @ReportHeader(name = "Odbiorca")
    private String recipient;
    @ReportHeader(name = "Kraj", columnWidthPdf = 0.3F)
    private String country;
    @ReportHeader(name = "Miasto")
    private String city;
    @ReportHeader(name = "Ulica", columnWidthPdf = 1.5F)
    private String street;
    @ReportHeader(name = "Data wysłania")
    private LocalDate sentDate;
    @ReportHeader(name = "Data dostarczenia", inPDF = false)
    private LocalDate deliveredDate;
}
