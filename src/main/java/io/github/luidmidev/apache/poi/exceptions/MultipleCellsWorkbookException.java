package io.github.luidmidev.apache.poi.exceptions;

import lombok.Getter;

@Getter
public class MultipleCellsWorkbookException extends WorkbookException {

    private final String reference;

    public MultipleCellsWorkbookException(String reference) {
        super("Multiple cells found for reference: " + reference + ", expected only one cell");
        this.reference = reference;
    }

}
