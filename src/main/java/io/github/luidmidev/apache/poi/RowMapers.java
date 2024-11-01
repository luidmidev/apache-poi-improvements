package io.github.luidmidev.apache.poi;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * Represents a collection of RowMapper instances for a specific object type, allowing
 * easy access to column names and mappers, as well as iteration over them.
 *
 * @param <T> The type of the objects mapped by this collection of RowMapper instances.
 */
@Getter
public class RowMapers<T> implements Iterable<RowMapper<T>> {

    private final List<RowMapper<T>> mappers = new ArrayList<>();

    /**
     * Retrieves the column names associated with each RowMapper in this collection.
     *
     * @return A list of column names, one for each RowMapper in the collection.
     */
    public List<String> getColumnNames() {
        return mappers.stream()
                .map(RowMapper::column)
                .toList();
    }

    /**
     * Returns an iterator over the RowMapper instances in this collection.
     *
     * @return An iterator for iterating over RowMapper elements.
     */
    @Override
    public Iterator<RowMapper<T>> iterator() {
        return mappers.iterator();
    }

    /**
     * Adds a new RowMapper to this collection with the specified column name, getter,
     * and cell configurator action.
     *
     * @param column The name of the column to associate with this RowMapper.
     * @param getter A function to retrieve the value from an instance of T.
     * @param action A CellWorkbookConsumer action to style or format the cell.
     */
    public void add(String column, Function<T, Object> getter, Consumer<Cell> action) {
        mappers.add(new RowMapper<>(column, getter, action));
    }
}