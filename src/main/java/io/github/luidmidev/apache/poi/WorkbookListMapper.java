package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.exceptions.WorkbookException;
import io.github.luidmidev.apache.poi.functions.Functionals;
import io.github.luidmidev.apache.poi.model.WorkbookType;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;


/**
 * Maps a list of objects to a workbook with a given configuration
 * @param <T>
 */
@Log4j2
public class WorkbookListMapper<T> {

    private final WorkbookManager workbookManager;
    private final ListMapperConfiguration<T> configuration;

    public WorkbookListMapper(List<T> data) {
        this(data, WorkbookType.XLSX);
    }

    public WorkbookListMapper(List<T> data, WorkbookType type) {
        this(data, new WorkbookManager(type), 0);
    }

    public WorkbookListMapper(List<T> data, WorkbookManager existingWorkbook, int startRow) {
        this.workbookManager = existingWorkbook;
        this.configuration = new ListMapperConfiguration<>(data, workbookManager.getWorkbook(), startRow);
    }

    public WorkbookManager map(ListMapperConfigurator<T> configurator) throws WorkbookException {

        configurator.apply(workbookManager, this.configuration);

        this.configuration.build();
        return workbookManager;
    }

    public static <T> WorkbookListMapper<T> from(List<T> models) {
        return new WorkbookListMapper<>(models);
    }

    public static <T> WorkbookListMapper<T> from(List<T> models, WorkbookManager existingWorkbook, int startRow) {
        return new WorkbookListMapper<>(models, existingWorkbook, startRow);
    }

    @FunctionalInterface
    public interface ListMapperConfigurator<T> {
        void apply(WorkbookManager manager, ListMapperConfiguration<T> configuration) throws WorkbookException;
    }

    public static class ListMapperConfiguration<T> {

        private final List<T> data;
        private final Workbook workbook;
        private final int startRow;

        private CellStyle headerStyle;
        private final RowMapers<T> rowMapers = new RowMapers<>();
        private Consumer<Sheet> sheetConsumer = Functionals.consumerNoAction();
        private BiConsumer<Integer, Integer> onProgress = Functionals.biConsumerNoAction();
        private BiConsumer<Row, T> rowConsumer = Functionals.biConsumerNoAction();

        private final Map<CellStylizer, CellStyle> computedStyles = new HashMap<>();

        private ListMapperConfiguration(List<T> data, Workbook workbook, int startRow) {
            this.data = data;
            this.startRow = startRow;
            this.workbook = workbook;
        }


        public ListMapperConfiguration<T> withColumn(String column, Function<T, Object> getter) {
            return withColumn(column, getter, Functionals.consumerNoAction());
        }

        public ListMapperConfiguration<T> withColumn(String column, Function<T, Object> getter, Consumer<Cell> cellConfigurator) {
            rowMapers.add(column, getter, cellConfigurator);
            return this;
        }

        public ListMapperConfiguration<T> withColumn(String column, Function<T, Object> getter, CellStylizer stylizer) {
            var style = computedStyles.computeIfAbsent(stylizer, (key) -> stylizer.build(workbook));
            return withColumn(column, getter, (cell) -> cell.setCellStyle(style));
        }

        public ListMapperConfiguration<T> withHeaderStyle(CellStylizer stylizer) {
            this.headerStyle = computedStyles.computeIfAbsent(stylizer, (key) -> stylizer.build(workbook));
            return this;
        }

        public ListMapperConfiguration<T> configureSheet(Consumer<Sheet> sheetConsumer) {
            this.sheetConsumer = sheetConsumer;
            return this;
        }

        public ListMapperConfiguration<T> onProgress(BiConsumer<Integer, Integer> onProgress) {
            this.onProgress = onProgress;
            return this;
        }

        private void build() throws WorkbookException {

            var sheet = getFirstSheet(workbook);

            if (sheet.getLastRowNum() > startRow) {
                sheet.shiftRows(startRow, sheet.getLastRowNum(), data.size());
            }

            int rowCounter = startRow;
            List<RowMapper<T>> mappers = rowMapers.getMappers();

            createRows(sheet, rowCounter, rowCounter + data.size() + 1);

            var rowHeader = sheet.getRow(rowCounter);
            var columns = rowMapers.getColumnNames();

            for (int i = 0; i < columns.size(); i++) {
                var cellHeader = rowHeader.createCell(i);
                cellHeader.setCellValue(columns.get(i));
                if (headerStyle != null) cellHeader.setCellStyle(headerStyle);
            }

            rowCounter++;
            final var size = data.size();
            for (int i = 0; i < size; i++) {
                writeRow(sheet, rowCounter + i, mappers, data.get(i));
                onProgress.accept(i + 1, size);
            }

            sheetConsumer.accept(sheet);
        }

        private void writeRow(Sheet sheet, int rowNum, List<RowMapper<T>> mappers, T model) throws WorkbookException {
            var row = sheet.getRow(rowNum);

            for (var j = 0; j < mappers.size(); j++) {
                RowMapper<T> mapper = mappers.get(j);

                var cell = row.createCell(j);
                var value = mapper.get(model);
                WorkbookManagerUtils.setCellValue(cell, value);
                mapper.action().accept(cell);
            }
            rowConsumer.accept(row, model);
        }

        private void createRows(Sheet sheet, int startRow, int endRow) {
            for (int i = startRow; i < endRow; i++) createRow(sheet, i);
        }

        private void createRow(Sheet sheet, int num) {
            var newRow = sheet.getRow(num);
            if (newRow != null) {
                log.trace("Row {} already exists, shifting rows, this could be a performance issue", num);
                sheet.shiftRows(num, sheet.getLastRowNum(), 1);
                return;
            }
            sheet.createRow(num);
        }

        public static Sheet getFirstSheet(Workbook workbook) {
            if (workbook.getNumberOfSheets() == 0) return workbook.createSheet();
            return workbook.getSheetAt(0);
        }

        public ListMapperConfiguration<T> forEachRow(BiConsumer<Row, T> rowConsumer) {
            this.rowConsumer = rowConsumer;
            return this;
        }
    }
}
