package pl.nowekolory.objectifyxlsx.cell;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class CellParameters {
    @Builder.Default
    private Boolean roundDouble = true;
    @Builder.Default
    private HorizontalAlignment horizontalAlignment = HorizontalAlignment.RIGHT;
    @Builder.Default
    private Boolean boldFont = false;
}