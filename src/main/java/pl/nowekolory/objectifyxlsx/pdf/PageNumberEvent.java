package pl.nowekolory.objectifyxlsx.pdf;


import org.openpdf.text.Document;
import org.openpdf.text.Element;
import org.openpdf.text.pdf.BaseFont;
import org.openpdf.text.pdf.PdfPageEventHelper;
import org.openpdf.text.pdf.PdfWriter;

public class PageNumberEvent extends PdfPageEventHelper{

    private static final float pageNumberTextSize = 8;

    @Override
    public void onEndPage(PdfWriter writer, Document document) {
        try {
            final var cb = writer.getDirectContent();
            cb.saveState();

            final var textBase = document.bottom() - 20;

            cb.beginText();
            cb.setFontAndSize(BaseFont.createFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.NOT_EMBEDDED), pageNumberTextSize);
            cb.showTextAligned(
                    Element.ALIGN_CENTER,
                    String.format("%d ", writer.getPageNumber()),
                    (document.right() + document.left()) / 2,
                    textBase,
                    0
            );
            cb.endText();

            cb.restoreState();
        } catch (Exception e) {
            throw new RuntimeException("Błąd podczas dodawania numeracji stron: " + e.getMessage(), e);
        }
    }

}
