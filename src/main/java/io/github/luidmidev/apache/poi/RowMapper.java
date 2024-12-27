package io.github.luidmidev.apache.poi;

import org.apache.poi.ss.usermodel.Cell;

import java.util.function.Consumer;

/**
 * Represents a row mapper that associates a cell value with a getter method of an object
 * and allows applying styling actions to the cell.
 */
public record RowMapper<T>(String column, Getter<T> getter, Consumer<Cell> action) {

    /**
     * Retrieves the value of the specified object based on the getter function.
     *
     * @param object The object from which to extract the value.
     * @return The value retrieved by the getter function.
     */
    Object get(T object, int rowNum) {
        return getter.get(object, rowNum);
    }

    @FunctionalInterface
    public interface Getter<T> {
        Object get(T object, int rowNum);
    }
}
