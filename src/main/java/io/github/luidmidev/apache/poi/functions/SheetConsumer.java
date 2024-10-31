package io.github.luidmidev.apache.poi.functions;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.function.Consumer;

@FunctionalInterface
public interface SheetConsumer extends Consumer<Sheet> {
}
