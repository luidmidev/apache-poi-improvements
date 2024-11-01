package io.github.luidmidev.apache.poi.exceptions;

import lombok.Getter;

@Getter
public class NotFoundRowWorkbookException extends WorkbookException {

    private final int index;

    public NotFoundRowWorkbookException(int index) {
        super("Row not found index: " + index);
        this.index = index;
    }
}
