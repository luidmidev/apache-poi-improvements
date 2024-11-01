package io.github.luidmidev.apache.poi.exceptions;

import lombok.Getter;

@Getter
public class NotFoundCellWorkbookException extends WorkbookException {

    private final int index;

    public NotFoundCellWorkbookException(int index) {
        super("Cell not found index: " + index);
        this.index = index;
    }
}
