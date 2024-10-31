package io.github.luidmidev.apache.poi.functions;

import org.apache.poi.ss.usermodel.Cell;

import java.util.function.Consumer;

@FunctionalInterface
public interface CellConsumer extends Consumer<Cell> {
}
