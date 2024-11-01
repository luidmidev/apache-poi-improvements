package io.github.luidmidev.apache.poi.exceptions;

import lombok.Getter;

@Getter
public class NotFoundSheetWorkbookException extends WorkbookException {

    private final String sheet;

    public NotFoundSheetWorkbookException(String sheetName) {
        super("Sheet not found: " + sheetName);
        this.sheet = sheetName;
    }

    public NotFoundSheetWorkbookException(int sheetIndex) {
        super("Sheet not found index: " + sheetIndex);
        this.sheet = "Sheet index: " + sheetIndex;
    }

}
