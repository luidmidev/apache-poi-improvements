package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.exceptions.WorkbookException;
import io.github.luidmidev.apache.poi.functions.CellConsumer;
import io.github.luidmidev.apache.poi.functions.CellWorkbookConsumer;
import io.github.luidmidev.apache.poi.functions.SheetConsumer;
import io.github.luidmidev.apache.poi.model.WorkbookType;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Function;

import static io.github.luidmidev.apache.poi.WorkbookManager.setCellValue;

@Log4j2
public class WorkbookModelBuilder<T> {

    private final List<T> data;

    private final ColumnsMapper<T> columnsMapper;

    private final Map<CellStylizer, CellStyle> computedStyles = new HashMap<>();

    private CellStyle headerStyle;

    private SheetConsumer sheetConsumer = sheet -> {
    };

    private BiConsumer<Integer, Integer> onProgress = (current, total) -> {
    };

    private BiConsumer<Row, T> rowConsumer = (row, model) -> {
    };


    private final WorkbookManager manager;
    private final int startRow;

    public WorkbookModelBuilder(List<T> data) {
        this(data, WorkbookType.XLSX);
    }

    public WorkbookModelBuilder(List<T> data, WorkbookType type) {
        this.data = data;
        this.columnsMapper = new ColumnsMapper<>();
        this.manager = new WorkbookManager(type == null ? WorkbookType.XLSX : type);
        startRow = 0;
    }

    public WorkbookModelBuilder(List<T> data, WorkbookManager existingWorkbook, int startRow) {
        this.data = data;
        this.columnsMapper = new ColumnsMapper<>();
        this.manager = existingWorkbook;
        this.startRow = startRow;
    }

    public Workbook getWorkbook() {
        return manager.getWorkbook();
    }


    public WorkbookModelBuilder<T> withColumn(String column, Function<T, Object> getter) {
        return withColumn(column, getter, cell -> {
        });
    }

    public WorkbookModelBuilder<T> withColumn(String column, Function<T, Object> getter, CellConsumer cellConfigurator) {
        columnsMapper.add(column, getter, (cell, workbook) -> cellConfigurator.accept(cell));
        return this;
    }

    public WorkbookModelBuilder<T> withColumn(String column, Function<T, Object> getter, CellWorkbookConsumer cellConfigurator) {
        columnsMapper.add(column, getter, cellConfigurator);
        return this;
    }

    public WorkbookModelBuilder<T> withColumn(String column, Function<T, Object> getter, CellStylizer stylizer) {
        var style = computedStyles.computeIfAbsent(stylizer, (key) -> stylizer.build(getWorkbook()));
        return withColumn(column, getter, (cell) -> cell.setCellStyle(style));
    }

    public WorkbookModelBuilder<T> withHeaderStyle(CellStylizer stylizer) {
        this.headerStyle = computedStyles.computeIfAbsent(stylizer, (key) -> stylizer.build(getWorkbook()));
        return this;
    }

    public WorkbookModelBuilder<T> configureSheet(SheetConsumer sheetConsumer) {
        this.sheetConsumer = sheetConsumer;
        return this;
    }

    public WorkbookModelBuilder<T> onProgress(BiConsumer<Integer, Integer> onProgress) {
        this.onProgress = onProgress;
        return this;
    }

    public WorkbookManager build() throws WorkbookException {

        var sheet = getFirstSheet(manager.getWorkbook());

        if (sheet.getLastRowNum() > startRow) {
            sheet.shiftRows(startRow, sheet.getLastRowNum(), data.size());
        }

        var rowCounter = startRow;
        List<ColumnMapper<T>> mappers = columnsMapper.getMappers();

        createRows(sheet, rowCounter, rowCounter + data.size() + 1);

        var rowHeader = sheet.getRow(rowCounter);
        var columns = columnsMapper.getColumnNames();

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

        return manager;
    }

    private void writeRow(Sheet sheet, int rowNum, List<ColumnMapper<T>> mappers, T model) throws WorkbookException {
        var row = sheet.getRow(rowNum);
        for (var j = 0; j < mappers.size(); j++) {
            ColumnMapper<T> mapper = mappers.get(j);
            var cell = row.createCell(j);
            var value = mapper.get(model);
            setCellValue(cell, value);
            mapper.stylizer().accept(cell, manager.getWorkbook());
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

    public WorkbookModelBuilder<T> forEachRow(BiConsumer<Row, T> rowConsumer) {
        this.rowConsumer = rowConsumer;
        return this;
    }
}
