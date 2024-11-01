package io.github.luidmidev.apache.poi;

import org.apache.poi.ss.usermodel.Cell;

import java.util.function.Consumer;
import java.util.function.Function;

/**
 * Represents a row mapper that associates a cell value with a getter method of an object
 * and allows applying styling actions to the cell.
 *
 * @param <T> The type of the object from which the cell value is extracted.
 * @param column The name of the column to which this cell mapping belongs.
 * @param getter A function to retrieve the value from the specified object.
 *               This value is then used to populate the cell.
 * @param action A {@link Consumer<Cell> } action applied to the cell
 *                 after the value has been set. This may involve setting font, color,
 *                 alignment, or other cell properties.
 *
 * <p> Example usage:
 * <pre>{@code
 * RowMapper<MyObject> mapper = new RowMapper<>("Name", MyObject::getName, cell -> {
 *     // Example styling action
 *     cell.setFont("Arial");
 *     cell.setBold(true);
 * });
 * }</pre>
 */
public record RowMapper<T>(String column, Function<T, Object> getter, Consumer<Cell> action) {

    /**
     * Retrieves the value of the specified object based on the getter function.
     *
     * @param object The object from which to extract the value.
     * @return The value retrieved by the getter function.
     */
    public Object get(T object) {
        return getter.apply(object);
    }
}
