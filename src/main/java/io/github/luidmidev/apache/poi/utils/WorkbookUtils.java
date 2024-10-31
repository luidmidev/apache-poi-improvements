package io.github.luidmidev.apache.poi.utils;


import io.github.luidmidev.apache.poi.WorkbookManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public final class WorkbookUtils {

    private WorkbookUtils() {
        throw new UnsupportedOperationException("This is a utility class and cannot be instantiated");
    }

    public static void adjustRowHeightByLines(Row row) {
        final Cell cell = getCellWithMostLines(row);
        final int fontIndex = cell.getCellStyle().getFontIndex();
        final float point = cell.getSheet().getWorkbook().getFontAt(fontIndex).getFontHeightInPoints();
        final float pointsByLines = calculatePointsByLines(cell.getStringCellValue(), point);
        row.setHeightInPoints(pointsByLines * 1.2f);
    }

    private static Cell getCellWithMostLines(Row row) {
        Cell cellWithMostLines = null;
        for (var cell : row) {
            var cellValue = WorkbookManager.getCellValue(cell);
            if (cellWithMostLines == null || cellValue.split("\n").length > cellWithMostLines.getStringCellValue().split("\n").length) {
                cellWithMostLines = cell;
            }
        }
        if (cellWithMostLines == null) {
            throw new IllegalArgumentException("Row must have at least one cell");
        }
        return cellWithMostLines;
    }

    private static float calculatePointsByLines(String content, float pointsPerLine) {
        var lines = content.split("\n").length;
        return pointsPerLine * lines;
    }

}