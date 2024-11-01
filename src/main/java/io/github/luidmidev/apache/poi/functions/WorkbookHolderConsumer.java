package io.github.luidmidev.apache.poi.functions;

import io.github.luidmidev.apache.poi.CellStylizer;

import java.util.function.Consumer;

/**
 * Represents an operation that accepts a single {@link CellStylizer.WorkbookHolder} argument and returns no result.
 */
@FunctionalInterface
public interface WorkbookHolderConsumer extends Consumer<CellStylizer.WorkbookHolder> {
}
