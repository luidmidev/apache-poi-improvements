package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.exceptions.NotFoundCellWorkbookException;
import io.github.luidmidev.apache.poi.exceptions.NotFoundRowWorkbookException;
import io.github.luidmidev.apache.poi.exceptions.UnsuportedCellValueTypeWorkbookException;
import io.github.luidmidev.apache.poi.model.WorkbookType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.Optional;

public final class WorkbookManagerUtils {

    private WorkbookManagerUtils() {
        throw new UnsupportedOperationException("Utility class");
    }


    /**
     * Devuelve una cadena que representa la referencia completa de una celda, incluyendo el nombre de la hoja.
     *
     * @param cell La celda de la cual se desea obtener la referencia completa.
     * @return Una cadena que representa la referencia completa de la celda en el formato "nombreHoja!referenciaCelda".
     */
    public static String getReference(Cell cell) {
        var rowIndex = cell.getRowIndex();
        var colIntex = cell.getColumnIndex();
        var sheetName = "'" + cell.getSheet().getSheetName() + "'";
        var cellReference = new CellReference(rowIndex, colIntex, true, true).formatAsString();
        return sheetName + "!" + cellReference;
    }

    /**
     * Adjusts the row height based on the cell with the most lines
     * @param row the row to adjust
     */
    public static void adjustRowHeightByLines(Row row, FormulaEvaluator evaluator) {
        final Cell cell = getCellWithMostLines(row, evaluator);
        final int fontIndex = cell.getCellStyle().getFontIndex();
        final float point = cell.getSheet().getWorkbook().getFontAt(fontIndex).getFontHeightInPoints();
        final float pointsByLines = calculatePointsByLines(cell.getStringCellValue(), point);
        row.setHeightInPoints(pointsByLines * 1.2f);
    }

    /**
     * Gets the cell with the most lines in a row
     * @param row the row to search
     * @return the cell with the most lines
     */
    public static Cell getCellWithMostLines(Row row, FormulaEvaluator evaluator) {
        Cell cellWithMostLines = null;
        for (var cell : row) {
            if (cell.getCellType() == CellType.STRING) {
                var cellValue = (String) WorkbookManagerUtils.getCellValue(cell, evaluator);
                if (cellWithMostLines == null || cellValue.split("\n").length > cellWithMostLines.getStringCellValue().split("\n").length) {
                    cellWithMostLines = cell;
                }
            }
        }
        if (cellWithMostLines == null) {
            throw new IllegalArgumentException("Row must have at least one cell");
        }
        return cellWithMostLines;
    }

    /**
     * Calculates the points based on the number of lines
     * @param content the content
     * @param pointsPerLine the points per line
     * @return the points
     */
    public static float calculatePointsByLines(String content, float pointsPerLine) {
        var lines = content.split("\n").length;
        return pointsPerLine * lines;
    }

    /**
     * Get a single cell by indexes and sheet
     * @param rowIndex the index of the row
     * @param cellIndex the index of the cell
     * @param sxxfSheet the sheet
     * @return the cell
     * @throws NotFoundRowWorkbookException if the row does not exist
     * @throws NotFoundCellWorkbookException if the cell does not exist
     */
    public static Cell getCell(int rowIndex, int cellIndex, Sheet sxxfSheet) throws NotFoundRowWorkbookException, NotFoundCellWorkbookException {
        var row = sxxfSheet.getRow(rowIndex);
        if (row == null) throw new NotFoundRowWorkbookException(rowIndex);
        var cell = row.getCell(cellIndex);
        if (cell == null) throw new NotFoundCellWorkbookException(cellIndex);
        return cell;
    }

    /**
     * Get a single cell by indexes and sheet
     * @param rowIndex the index of the row
     * @param cellIndex the index of the cell
     * @param sxxfSheet the sheet
     * @return the cell
     * */
    public static Cell getCellSafe(int rowIndex, int cellIndex, Sheet sxxfSheet) {
        var row = Optional.ofNullable(sxxfSheet.getRow(rowIndex))
                .orElseGet(() -> sxxfSheet.createRow(rowIndex));

        return Optional.ofNullable(row.getCell(cellIndex))
                .orElseGet(() -> row.createCell(cellIndex));
    }

    /**
     * Set the value of a cell dynamically
     * @param cell the cell to set the value
     * @param value the value to set
     * @throws UnsuportedCellValueTypeWorkbookException if the value type is not supported
     */
    public static void setCellValue(Cell cell, Object value) throws UnsuportedCellValueTypeWorkbookException {
        switch (value) {
            case String casted -> cell.setCellValue(casted);
            case Number casted -> cell.setCellValue(casted.doubleValue());
            case Boolean casted -> cell.setCellValue(casted);
            case Date casted -> cell.setCellValue(casted);
            case LocalDate casted -> cell.setCellValue(casted);
            case LocalDateTime casted -> cell.setCellValue(casted);
            case Calendar casted -> cell.setCellValue(casted);
            case RichTextString casted -> cell.setCellValue(casted);
            case null -> cell.setCellValue("");
            default -> throw new UnsuportedCellValueTypeWorkbookException(cell, value.getClass());
        }
    }

    /**
     * Copy a row from a worksheet to another
     * @param worksheet the worksheet
     * @param sourceRowNum the source row number
     * @param destinationRowNum the destination row number
     * @return the new row
     */
    public static Row copyRow(Sheet worksheet, int sourceRowNum, int destinationRowNum) {

        var sourceRow = worksheet.getRow(sourceRowNum);
        var newRow = worksheet.getRow(destinationRowNum);

        if (newRow != null) {
            worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
        }
        newRow = worksheet.createRow(destinationRowNum);

        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {

            var oldCell = sourceRow.getCell(i);
            var newCell = newRow.createCell(i);

            if (oldCell == null) continue;


            var newCellStyle = worksheet.getWorkbook().createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);

            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }


            switch (oldCell.getCellType()) {
                case BLANK -> newCell.setCellValue(oldCell.getStringCellValue());
                case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
                case FORMULA -> newCell.setCellFormula(oldCell.getCellFormula());
                case NUMERIC -> newCell.setCellValue(oldCell.getNumericCellValue());
                case STRING -> newCell.setCellValue(oldCell.getRichStringCellValue());
                case ERROR -> newCell.setCellErrorValue(oldCell.getErrorCellValue());
                case _NONE -> {
                    // Do nothing
                }
            }
        }

        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            var cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                var newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(), (newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())), cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }
        return newRow;
    }


    /**
     * Obtiene el valor de la celda especificada utilizando un evaluador de fórmulas y lo devuelve como cadena de caracteres.
     *
     * @param cell Celda de la que se desea obtener el valor.
     * @return Valor de la celda como cadena de caracteres.
     */
    private static Object getCellValue(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case _NONE, BLANK -> "";
            case NUMERIC -> cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            case FORMULA -> getCellValue(evaluator.evaluate(cell));
            case ERROR -> cell.getErrorCellValue();
        };
    }

    /**
     * Obtiene el valor de la celdas especificadas y lo delvuelve como cadena de caracteres
     *
     * @param cellValue CellValue de las que se desea obtener el valor
     * @return Valor de la celda como cadena de caracteres
     */
    public static Object getCellValue(CellValue cellValue) {
        var cellType = cellValue.getCellType();
        return switch (cellType) {
            case NUMERIC -> cellValue.getNumberValue();
            case STRING -> cellValue.getStringValue();
            case BOOLEAN -> cellValue.getBooleanValue();
            case ERROR -> "Error";
            default -> "<error unexpected cell type " + cellType + ">";
        };
    }


    public static WorkbookType resolveWorkbookType(Workbook workbook) {
        return switch (workbook) {
            case HSSFWorkbook ignored -> WorkbookType.XLS;
            case XSSFWorkbook xssfWorkbook -> resolveWorkbookType(xssfWorkbook);
            case SXSSFWorkbook sxssfWorkbook -> resolveWorkbookType(sxssfWorkbook.getXSSFWorkbook());
            default -> throw new UnsupportedOperationException("Unsupported workbook type");
        };
    }

    public static WorkbookType resolveWorkbookType(XSSFWorkbook workbook) {
        return workbook.getWorkbookType() == XSSFWorkbookType.XLSX ? WorkbookType.XLSX : WorkbookType.XLS;
    }
}
