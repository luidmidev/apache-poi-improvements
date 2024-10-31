package io.github.luidmidev.apache.poi.functions;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.function.BiConsumer;

@FunctionalInterface
public interface CellWorkbookConsumer extends BiConsumer<Cell, Workbook> {
}
