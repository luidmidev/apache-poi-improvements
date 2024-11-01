package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.exceptions.*;
import io.github.luidmidev.apache.poi.model.SpreadSheetFile;
import io.github.luidmidev.apache.poi.model.WorkbookType;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;


/**
 * Represents a manager for a workbook that provides utility methods to work with it.
 */
@Getter
public class WorkbookManager implements Closeable, AutoCloseable {

    protected final Workbook workbook;
    protected final WorkbookType type;
    protected final FormulaEvaluator evaluator;

    /**
     * Creates a new instance of {@link WorkbookManager} with a new XSSFWorkbook.
     */
    public WorkbookManager() {
        this(new XSSFWorkbook());
    }

    /**
     * Creates a new instance of {@link WorkbookManager} with a new workbook of the specified type.
     *
     * @param type the type of the workbook
     */
    public WorkbookManager(WorkbookType type) {
        this(type == WorkbookType.XLSX || type == WorkbookType.XLSM
                ? new XSSFWorkbook(type == WorkbookType.XLSM ? XSSFWorkbookType.XLSM : XSSFWorkbookType.XLSX)
                : new HSSFWorkbook()
        );
    }

    /**
     * Creates a new instance of {@link WorkbookManager} with the specified spreadsheet file.
     *
     * @param filePath the path of the file
     * @throws IOException if an I/O error occurs
     */
    public WorkbookManager(String filePath) throws IOException {
        this(new FileInputStream(filePath));
    }


    /**
     * Creates a new instance of {@link WorkbookManager} with the specified content bytes.
     * @param file the file
     * @throws IOException if an I/O error occurs
     */
    public WorkbookManager(byte[] file) throws IOException {
        this(new ByteArrayInputStream(file));
    }


    /**
     * Creates a new instance of {@link WorkbookManager} with the specified workbook.
     * @param workbook the workbook
     */
    private WorkbookManager(Workbook workbook) {
        this.workbook = workbook;
        this.type = WorkbookManagerUtils.resolveWorkbookType(workbook);
        this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
    }


    /**
     * Creates a new instance of {@link WorkbookManager} with the specified input stream
     * @param inputStream the input stream
     * @throws IOException if an I/O error occurs
     */
    public WorkbookManager(InputStream inputStream) throws IOException {
        this(WorkbookFactory.create(inputStream));
    }


    /**
     * Creates a new instance of {@link WorkbookManager} with the specified file and type.
     * @param reference the file
     * @return a new instance of WorkbookManager
     * @throws NotFoundCellWorkbookException if the cell does not exist
     * @throws NotFoundSheetWorkbookException if the sheet does not exist
     * @throws NotFoundRowWorkbookException if the row does not exist
     * @throws MultipleCellsWorkbookException if multiple cells are found
     */
    public Cell getCell(String reference) throws NotFoundCellWorkbookException, NotFoundSheetWorkbookException, NotFoundRowWorkbookException, MultipleCellsWorkbookException {
        var cells = getCells(reference);
        if (cells.length > 1) throw new MultipleCellsWorkbookException(reference);
        return cells[0];

    }

    /**
     * Get a single cell by its reference.
     * @param reference the reference of the cell
     * @return the cell
     * @throws NotFoundSheetWorkbookException if the sheet does not exist
     */
    public Cell getCell(CellReference reference) throws NotFoundSheetWorkbookException, NotFoundCellWorkbookException, NotFoundRowWorkbookException {
        return getCell(reference.getSheetName(), reference.getRow(), reference.getCol());
    }

    /**
     * Get a single cell by indexes and sheet name.
     * @param sheetName the name of the sheet
     * @param rowIndex the index of the row
     * @param cellIndex the index of the cell
     * @return the cell
     * @throws NotFoundSheetWorkbookException if the sheet does not exist
     * @throws NotFoundCellWorkbookException if the cell does not exist
     * @throws NotFoundRowWorkbookException if the row does not exist
     */
    public Cell getCell(String sheetName, int rowIndex, int cellIndex) throws NotFoundSheetWorkbookException, NotFoundCellWorkbookException, NotFoundRowWorkbookException {
        var sxxfSheet = workbook.getSheet(sheetName);
        if (sxxfSheet == null) throw new NotFoundSheetWorkbookException(sheetName);
        return WorkbookManagerUtils.getCell(rowIndex, cellIndex, sxxfSheet);
    }

    /**
     * Get a single cell by indexes.
     * @param sheetIndex the index of the sheet
     * @param rowIndex the index of the row
     * @param cellIndex the index of the cell
     * @return the cell
     * @throws NotFoundSheetWorkbookException if the sheet does not exist
     */
    public Cell getCell(int sheetIndex, int rowIndex, int cellIndex) throws NotFoundSheetWorkbookException, NotFoundCellWorkbookException, NotFoundRowWorkbookException {
        var sxxfSheet = workbook.getSheetAt(sheetIndex);
        if (sxxfSheet == null) throw new NotFoundSheetWorkbookException(sheetIndex);
        return WorkbookManagerUtils.getCell(rowIndex, cellIndex, sxxfSheet);
    }

    public Cell[] getCells(String reference) throws NotFoundCellWorkbookException, NotFoundSheetWorkbookException, NotFoundRowWorkbookException {
        var cellsReferences = workbook.getName(reference) != null ? getCellsReferencesFromName(reference) : getCellsReferencesFromFormula(reference);
        return getCells(cellsReferences);
    }

    public Cell[] getCells(CellReference... cellsReferences) throws NotFoundCellWorkbookException, NotFoundSheetWorkbookException, NotFoundRowWorkbookException {
        var cells = new Cell[cellsReferences.length];
        for (int i = 0; i < cellsReferences.length; i++) {
            cells[i] = getCell(cellsReferences[i]);
        }
        return cells;
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
     * Get the input stream of the workbook of the current state.
     * @return the input stream of the workbook
     * @throws IOException if an I/O error occurs
     */
    public ByteArrayInputStream getInputStream() throws IOException {
        return new ByteArrayInputStream(getBytes());
    }


    /**
     * Get the bytes of the workbook of the current state.
     * @return the bytes of the workbook
     * @throws IOException if an I/O error occurs
     */
    private byte[] getBytes() throws IOException {
        var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
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
        return new WorkbookManager(getInputStream());

    }


    public SpreadSheetFile getSpreadsheet(String name) throws IOException {
        var report = new SpreadSheetFile();
        var filename = name + "." + type.getExtension();

        report.setFilename(filename);
        report.setContent(getBytes());
        report.setType(type);

        return report;
    }
}
