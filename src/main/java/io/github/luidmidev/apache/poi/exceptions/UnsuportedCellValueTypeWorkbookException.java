package io.github.luidmidev.apache.poi.exceptions;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;

@Getter
public class UnsuportedCellValueTypeWorkbookException extends WorkbookException {

    private final Cell cellValue;
    private final Class<?> castType;

    public UnsuportedCellValueTypeWorkbookException(Cell cellValue, Class<?> castType) {
        super("Unsupported cell value: " + cellValue + " for type: " + castType);
        this.cellValue = cellValue;
        this.castType = castType;
    }
}
