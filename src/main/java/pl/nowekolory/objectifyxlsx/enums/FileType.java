package pl.nowekolory.objectifyxlsx.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum FileType{

    XLSX(".xlsx"),
    PDF(".pdf");

    private final String extension;

}
