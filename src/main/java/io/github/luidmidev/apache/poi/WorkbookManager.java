package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.exceptions.WorkbookException;
import io.github.luidmidev.apache.poi.model.ReportFile;
import io.github.luidmidev.apache.poi.model.WorkbookType;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.function.IntFunction;


public class WorkbookManager implements Closeable, AutoCloseable {

    @Getter
    protected Workbook workbook;

    protected FormulaEvaluator evaluator;

    @Getter
    protected WorkbookType type;

    /**
     * Constructor para crear un objeto ExcelModel a partir de un nuevo archivo de Excel.
     */
    public WorkbookManager() {
        this(new XSSFWorkbook());
        this.type = WorkbookType.XLSX;
    }

    /**
     * Constructor para crear un objeto ExcelModel a partir de un nuevo archivo de Excel.
     *
     * @param type el tipo de archivo de Excel
     */
    public WorkbookManager(WorkbookType type) {
        this(type == WorkbookType.XLSX || type == WorkbookType.XLSM ? new XSSFWorkbook() : new HSSFWorkbook());
        this.type = type;
    }

    /**
     * Constructor para crear un objeto ExcelModel a partir de una ruta de archivo.
     *
     * @param filePath la ruta del archivo de Excel
     * @throws IOException si el archivo no se encuentra o no se puede leer
     */
    public WorkbookManager(String filePath) throws IOException {
        this(new FileInputStream(filePath), WorkbookType.fromFilename(filePath));
    }

    /**
     * Constructor para crear un objeto ExcelModel a partir de un arreglo de bytes.
     *
     * @param file el arreglo de bytes que contiene los datos del archivo de Excel
     * @throws IOException si el arreglo de bytes no contiene datos de un archivo de Excel
     */
    public WorkbookManager(byte[] file, WorkbookType type) throws IOException {
        this(new ByteArrayInputStream(file), type);
    }

    /**
     * Constructor para crear un objeto ExcelModel a partir de un objeto XSSFWorkbook.
     */
    private WorkbookManager(Workbook workbook) {
        this.workbook = workbook;
        this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
    }

    /**
     * Constructor para crear un objeto ExcelModel a partir de un objeto InputStream.
     *
     * @param inputStream el objeto InputStream que contiene los datos del archivo de Excel
     * @throws IOException si el objeto InputStream no contiene datos de un archivo de Excel
     */
    public WorkbookManager(InputStream inputStream, WorkbookType type) throws IOException {
        this(type == WorkbookType.XLSX || type == WorkbookType.XLSM ? new XSSFWorkbook(inputStream) : new HSSFWorkbook(inputStream));
        this.type = type;
    }

    /**
     * Obtiene un array de celdas asociadas a un nombre de celda dado.
     *
     * @param reference referencia de la celda, ya sea nombre o formula
     * @return un array de celdas asociadas al nombre de celda dado
     */
    public Cell getCell(String reference) throws WorkbookException {
        var cells = getCells(reference);
        if (cells.length > 1) throw new WorkbookException("El nombre " + reference + " debe hacer referencia a una unica celda");
        return cells[0];

    }

    public Cell getCell(CellReference reference) throws WorkbookException {
        var sxxfSheet = workbook.getSheet(reference.getSheetName());
        if (sxxfSheet == null) throw new WorkbookException("No existe la hoja de calculo con nombre " + reference.getSheetName());
        return getCell(reference.getRow(), reference.getCol(), sxxfSheet);
    }

    public Cell getCell(int sheetIndex, int rowIndex, int cellIndex) throws WorkbookException {
        var sxxfSheet = workbook.getSheetAt(sheetIndex);
        if (sxxfSheet == null) throw new WorkbookException("No existe la hoja de calculo con indice " + sheetIndex);
        return getCell(rowIndex, cellIndex, sxxfSheet);
    }


    public Cell getCell(String sheetName, int rowIndex, int cellIndex) throws WorkbookException {
        var sxxfSheet = workbook.getSheet(sheetName);
        if (sxxfSheet == null) throw new WorkbookException("No existe la hoja de calculo con nombre " + sheetName);
        return getCell(rowIndex, cellIndex, sxxfSheet);
    }

    private static Cell getCell(int rowIndex, int cellIndex, Sheet sxxfSheet) {
        var row = sxxfSheet.getRow(rowIndex);
        if (row == null) row = sxxfSheet.createRow(rowIndex);
        var cell = row.getCell(cellIndex);
        if (cell == null) cell = row.createCell(cellIndex);
        return cell;
    }

    public <T> T getCellValue(String reference, Class<T> type) throws WorkbookException {
        return getCellValue(getCell(reference), type);
    }

    public <T> T getCellValue(CellReference reference, Class<T> type) throws WorkbookException {
        return getCellValue(getCell(reference), type);
    }

    public <T> T getCellValue(int sheetIndex, int rowIndex, int cellIndex, Class<T> type) throws WorkbookException {
        return getCellValue(getCell(sheetIndex, rowIndex, cellIndex), type);
    }

    public <T> T getCellValue(String sheetName, int rowIndex, int cellIndex, Class<T> type) throws WorkbookException {
        return getCellValue(getCell(sheetName, rowIndex, cellIndex), type);
    }

    public <T> T getCellValue(Cell cell, Class<T> type) throws WorkbookException, ClassCastException {
        final String value = getCellValueWithEvaluator(cell);

        if (type.equals(String.class)) {
            return type.cast(value);
        }

        if (type.equals(Double.class)) {
            return type.cast(Double.parseDouble(value));
        }

        if (type.equals(Integer.class)) {
            return type.cast(Integer.parseInt(value));
        }

        if (type.equals(Boolean.class)) {
            return type.cast(Boolean.parseBoolean(value));
        }

        if (type.equals(Date.class)) {
            return type.cast(cell.getDateCellValue());
        }

        if (type.equals(LocalDate.class)) {
            return type.cast(cell.getLocalDateTimeCellValue().toLocalDate());
        }

        if (type.equals(LocalDateTime.class)) {
            return type.cast(cell.getLocalDateTimeCellValue());
        }

        if (type.equals(Calendar.class)) {
            var calendar = Calendar.getInstance();
            calendar.setTime(cell.getDateCellValue());
            return type.cast(calendar);
        }

        if (type.equals(RichTextString.class)) {
            return type.cast(cell.getRichStringCellValue());
        }

        throw new WorkbookException("El tipo de dato no es soportado");


    }

    public void setCellValue(String reference, Object value) throws WorkbookException {
        setCellValue(getCell(reference), value);
    }

    public void setCellValue(CellReference reference, Object value) throws WorkbookException {
        setCellValue(getCell(reference), value);
    }

    public void setCellValue(int sheetIndex, int rowIndex, int cellIndex, Object value) throws WorkbookException {
        setCellValue(getCell(sheetIndex, rowIndex, cellIndex), value);
    }

    public void setCellValue(String sheetName, int rowIndex, int cellIndex, Object value) throws WorkbookException {
        setCellValue(getCell(sheetName, rowIndex, cellIndex), value);
    }

    static void setCellValue(Cell cell, Object value) throws WorkbookException {
        switch (value) {
            case String casted -> cell.setCellValue(casted);
            case Number casted -> cell.setCellValue(casted.doubleValue());
            case Boolean casted -> cell.setCellValue(casted);
            case Date casted -> cell.setCellValue(casted);
            case LocalDate casted -> cell.setCellValue(casted);
            case LocalDateTime casted -> cell.setCellValue(casted);
            case Calendar casted -> cell.setCellValue(casted);
            case RichTextString casted -> cell.setCellValue(casted);
            case null, default -> throw new WorkbookException("El tipo de dato no es soportado");
        }
    }


    public Cell[] getCells(String reference) throws WorkbookException {
        CellReference[] cellsReferences = workbook.getName(reference) != null ? getCellsReferencesFromName(reference) : getCellsReferencesFromFormula(reference);
        return getCells(cellsReferences);
    }

    public Cell[] getCells(CellReference... cellsReferences) throws WorkbookException {
        var cells = new Cell[cellsReferences.length];
        for (int i = 0; i < cellsReferences.length; i++) {
            cells[i] = getCell(cellsReferences[i]);
        }
        return cells;
    }


    public <T> T[] getCellsValues(String reference, Class<T> type, IntFunction<T[]> generator) throws WorkbookException {
        return getCellsValues(getCells(reference), type, generator);
    }

    public <T> T[] getCellsValues(CellReference[] cellsReferences, Class<T> type, IntFunction<T[]> generator) throws WorkbookException {
        return getCellsValues(getCells(cellsReferences), type, generator);
    }

    public <T> T[] getCellsValues(Cell[] cells, Class<T> type, IntFunction<T[]> generator) throws WorkbookException {
        var values = generator.apply(cells.length);
        for (int i = 0; i < cells.length; i++) {
            values[i] = getCellValue(cells[i], type);
        }
        return values;
    }


    /**
     * Obtiene un arreglo de objetos CellReference que representan las celdas incluidas en el
     * rango con nombre especificado en el libro de trabajo.
     *
     * @param name el nombre del rango de celdas para el cual se desea obtener los objetos CellReference
     * @return un arreglo de objetos CellReference que representan las celdas incluidas en el rango
     * con nombre especificado
     */
    private CellReference[] getCellsReferencesFromName(String name) {
        var xssfName = workbook.getName(name);
        var formula = xssfName.getRefersToFormula();
        return getCellsReferencesFromFormula(formula);
    }


    /**
     * Obtiene un arreglo de objetos CellReference que representan las celdas incluidas en el rango
     *
     * @param formula La formula que contiene el rango de celdas.
     * @return Un arreglo de objetos CellReference que representan las celdas incluidas en el rango.
     */
    private CellReference[] getCellsReferencesFromFormula(String formula) {
        var areaReference = new AreaReference(formula, workbook.getSpreadsheetVersion());
        return areaReference.getAllReferencedCells();
    }


    /**
     * Obtiene el valor de la celdas especificadas y lo delvuelve como cadena de caracteres
     *
     * @param cellValue CellValue de las que se desea obtener el valor
     * @return Valor de la celda como cadena de caracteres
     */
    public static String getCellValue(CellValue cellValue) throws WorkbookException {
        if (cellValue == null) return "";
        return switch (cellValue.getCellType()) {
            case STRING -> cellValue.getStringValue();
            case _NONE, BLANK -> "";
            case NUMERIC -> String.valueOf(cellValue.getNumberValue());
            case BOOLEAN -> String.valueOf(cellValue.getBooleanValue());
            case ERROR -> "Error";
            case FORMULA -> throw new WorkbookException("No se puede obtener el valor de una celda de tipo fórmula, en su lugar, utilice el metodo de instancia");
        };
    }

    public static String getCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case _NONE, BLANK -> "";
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case ERROR -> "ERROR";
            case FORMULA -> "FORMULA";
        };
    }

    /**
     * Obtiene el valor de la celda especificada utilizando un evaluador de fórmulas y lo devuelve como cadena de caracteres.
     *
     * @param cell Celda de la que se desea obtener el valor.
     * @return Valor de la celda como cadena de caracteres.
     */
    private String getCellValueWithEvaluator(Cell cell) throws WorkbookException {
        if (cell == null) return "";

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case _NONE -> throw new WorkbookException("No se puede obtener el valor de una celda vacía");
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> getCellValue(evaluator.evaluate(cell));
            case BLANK -> "";
            case ERROR -> "Error";
        };
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
     * Obtiene un objeto ByteArrayInputStream a partir del Workbook actual.
     *
     * @return Objeto ByteArrayInputStream que contiene los datos del Workbook.
     * @throws IOException Si ocurre un error al escribir en el ByteArrayOutputStream o al crear el ByteArrayInputStream.
     */
    public ByteArrayInputStream getInputStream() throws IOException {
        return new ByteArrayInputStream(getBytes());
    }


    private byte[] getBytes() throws IOException {
        evaluator.clearAllCachedResultValues();
        evaluator.evaluateAll();
        var bos = new ByteArrayOutputStream();
        workbook.write(bos);
        byte[] bytes = bos.toByteArray();
        bos.close();
        return bytes;
    }

    @Override
    public void close() throws IOException {
        workbook.close();
    }

    /**
     * Clone the current workbook.
     *
     * @return a new instance of WokbookManager with the same workbook
     */
    public WorkbookManager copy() throws IOException {
        InputStream inputStream = getInputStream();
        return new WorkbookManager(inputStream, type);

    }


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

    public static <T> WorkbookModelBuilder<T> fromItems(List<T> models) {
        return new WorkbookModelBuilder<>(models);
    }

    public static <T> WorkbookModelBuilder<T> fromItems(List<T> models, WorkbookManager existingWorkbook, int startRow) {
        return new WorkbookModelBuilder<>(models, existingWorkbook, startRow);
    }


    public ReportFile getSpreadsheet(String name) throws IOException {
        var report = new ReportFile();
        var filename = name + "." + type.getExtension();

        report.setFilename(filename);
        report.setContent(getBytes());
        report.setMediaType(type.getMediaType());

        return report;
    }
}
