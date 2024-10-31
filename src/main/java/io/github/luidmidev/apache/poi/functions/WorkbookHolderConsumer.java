package io.github.luidmidev.apache.poi.functions;

import io.github.luidmidev.apache.poi.CellStylizer;

import java.util.function.Consumer;

@FunctionalInterface
public interface WorkbookHolderConsumer extends Consumer<CellStylizer.WorkbookHolder> {
}
