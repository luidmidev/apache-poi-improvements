package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.functions.CellWorkbookConsumer;

import java.util.function.Function;

public record ColumnMapper<T>(String column, Function<T, Object> getter, CellWorkbookConsumer stylizer) {

    public Object get(T object) {
        return getter.apply(object);
    }
}
