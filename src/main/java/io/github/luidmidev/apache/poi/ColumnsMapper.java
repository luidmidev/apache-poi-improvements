package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.functions.CellWorkbookConsumer;
import lombok.Getter;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.function.Function;

@Getter
public class ColumnsMapper<T> implements Iterable<ColumnMapper<T>> {

    private final List<ColumnMapper<T>> mappers = new ArrayList<>();


    public List<String> getColumnNames() {
        return mappers.stream()
                .map(ColumnMapper::column)
                .toList();
    }

    @Override
    public Iterator<ColumnMapper<T>> iterator() {
        return mappers.iterator();
    }

    public void add(String column, Function<T, Object> getter, CellWorkbookConsumer cellConfigurator) {
        mappers.add(new ColumnMapper<>(column, getter, cellConfigurator));
    }
}
