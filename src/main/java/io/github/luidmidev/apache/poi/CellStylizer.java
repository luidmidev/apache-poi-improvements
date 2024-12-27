package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.functions.WorkbookHolderConsumer;
import lombok.Data;
import lombok.RequiredArgsConstructor;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

/**
 * A fluent API for styling cells.
 */
@Log4j2
public class CellStylizer {

    private final List<WorkbookHolderConsumer> consumers = new ArrayList<>();

    /**
     * Initializes the cell stylizer.
     * @return a new instance of {@link CellStylizer}
     */
    public static CellStylizer init() {
        return new CellStylizer();
    }

    /**
     * Builds the cell style.
     * @param workbook the workbook to build the style for
     * @return the cell style
     */
    CellStyle build(Workbook workbook) {

        var holder = new WorkbookHolder(workbook);

        if (consumers.isEmpty()) {
            log.debug("No style was applied to the cell");
        }

        for (var consumer : consumers) {
            consumer.accept(holder);
        }

        return holder.getStyle();
    }

    /**
     * Sets the border style for all borders, if borderStyle is null, the border style will be set to NONE.
     * @param borderStyle the border style to set
     * @return the current cell stylizer
     */
    public CellStylizer allBorders(BorderStyle borderStyle) {
        consumers.add(holder -> setBorders(holder.getStyle(), borderStyle, borderStyle, borderStyle, borderStyle));
        return this;
    }

    /**
     * Sets the border style for the specified borders, if borderStyle is null, the border style will be set to NONE.
     * @param top the border style for the top border
     * @param right the border style for the right border
     * @param bottom the border style for the bottom border
     * @param left the border style for the left border
     * @return the current cell stylizer
     */
    public CellStylizer onlyBorders(BorderStyle top, BorderStyle right, BorderStyle bottom, BorderStyle left) {
        consumers.add(holder -> setBorders(holder.getStyle(), top, right, bottom, left));
        return this;
    }

    /**
     * Centers the cell content.
     * @return the current cell stylizer
     */
    public CellStylizer center() {
        consumers.add(holder -> setCenter(holder.getStyle()));
        return this;
    }

    /**
     * Sets the foreground color for the cell.
     * @param indexedColors the indexed color to set
     * @return the current cell stylizer
     */
    public CellStylizer foregroundColor(IndexedColors indexedColors) {
        return foregroundColor(indexedColors.getIndex());
    }

    /**
     * Sets the foreground color for the cell.
     * @param index the color index to set
     * @return the current cell stylizer
     */
    public CellStylizer foregroundColor(short index) {
        return foregroundColor(index, FillPatternType.SOLID_FOREGROUND);
    }

    /**
     * Sets the foreground color for the cell.
     * @param indexedColors the indexed color to set
     * @param fillPattern the fill pattern to set
     * @return the current cell stylizer
     */
    public CellStylizer foregroundColor(IndexedColors indexedColors, FillPatternType fillPattern) {
        return foregroundColor(indexedColors.getIndex(), fillPattern);
    }

    public CellStylizer foregroundColor(int red, int green, int blue, FillPatternType fillPattern) {
        consumers.add(holder -> setForeground(holder.getStyle(), resolveColor(red, green, blue, holder.getWorkbook()), fillPattern));
        return this;
    }

    /**
     * Sets the foreground color for the cell.
     * @param red the red component
     * @param green the green component
     * @param blue the blue component
     * @return the current cell stylizer
     */
    public CellStylizer foregroundColor(int red, int green, int blue) {
        return foregroundColor(red, green, blue, FillPatternType.SOLID_FOREGROUND);
    }

    /**
     * Sets the foreground color for the cell.
     * @param index the color index to set
     * @param fillPattern the fill pattern to set
     * @return the current cell stylizer
     */
    public CellStylizer foregroundColor(short index, FillPatternType fillPattern) {
        consumers.add(holder -> setForeground(holder.getStyle(), index, fillPattern));
        return this;
    }

    /**
     * Sets the foreground color for the cell.
     * @param color the color to set
     * @param fillPattern the fill pattern to set
     * @return the current cell stylizer
     */
    public CellStylizer foregroundColor(Color color, FillPatternType fillPattern) {
        consumers.add(holder -> setForeground(holder.getStyle(), color, fillPattern));
        return this;
    }

    /**
     * Sets the foreground color for the cell.
     * @param indexedColors the indexed color to set
     * @return the current cell stylizer
     */
    public CellStylizer fontColor(IndexedColors indexedColors) {
        return fontColor(indexedColors.getIndex());
    }

    /**
     * Sets the foreground color for the cell.
     * @param index the color index to set
     * @return the current cell stylizer
     */
    public CellStylizer fontColor(short index) {
        consumers.add(holder -> holder.getFont().setColor(index));
        return this;
    }

    /**
     * Sets the foreground color for the cell.
     * @return the current cell stylizer
     */
    public CellStylizer fontBold() {
        consumers.add(holder -> holder.getFont().setBold(true));
        return this;
    }

    /**
     * Sets the font size for the cell in points.
     * @param fontSize the font size to set
     * @return the current cell stylizer
     */
    public CellStylizer fontSize(int fontSize) {
        consumers.add(holder -> holder.getFont().setFontHeightInPoints((short) fontSize));
        return this;
    }

    /**
     * Sets the font name to apply to the cell.
     * @param fontName the font name to set
     * @return the current cell stylizer
     */
    public CellStylizer fontName(String fontName) {
        consumers.add(holder -> holder.getFont().setFontName(fontName));
        return this;
    }

    /**
     * Set horizontal alignment for the cell.
     * @param horizontalAlignment the horizontal alignment to set
     * @return the current cell stylizer
     */
    public CellStylizer alignment(HorizontalAlignment horizontalAlignment) {
        consumers.add(holder -> holder.getStyle().setAlignment(horizontalAlignment));
        return this;
    }

    /**
     * Set vertical alignment for the cell.
     * @param verticalAlignment the vertical alignment to set
     * @return the current cell stylizer
     */
    public CellStylizer alignment(VerticalAlignment verticalAlignment) {
        consumers.add(holder -> holder.getStyle().setVerticalAlignment(verticalAlignment));
        return this;
    }

    /**
     * Set the cell to wrap text.
     * @return the current cell stylizer
     */
    public CellStylizer wrapText() {
        consumers.add(holder -> holder.getStyle().setWrapText(true));
        return this;
    }

    /**
     * Set the cell to not wrap text.
     * @param style the cell style to set
     * @param top the border style for the top border
     * @param right the border style for the right border
     * @param bottom the border style for the bottom border
     * @param left the border style for the left border
     */
    private static void setBorders(CellStyle style, BorderStyle top, BorderStyle right, BorderStyle bottom, BorderStyle left) {
        style.setBorderTop(top == null ? BorderStyle.NONE : top);
        style.setBorderRight(right == null ? BorderStyle.NONE : right);
        style.setBorderBottom(bottom == null ? BorderStyle.NONE : bottom);
        style.setBorderLeft(left == null ? BorderStyle.NONE : left);
    }

    /**
     * Set the cell to center.
     * @param style the cell style to set
     */
    private static void setCenter(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    }


    /**
     * Set the foreground color for the cell.
     * @param style the cell style to set
     * @param color the color to set
     * @param fillPattern the fill pattern to set
     */
    private static void setForeground(CellStyle style, short color, FillPatternType fillPattern) {
        style.setFillForegroundColor(color);
        style.setFillPattern(fillPattern);
    }

    /**
     * Set the foreground color for the cell.
     * @param style the cell style to set
     * @param color the color to set
     * @param fillPattern the fill pattern to set
     */
    private static void setForeground(CellStyle style, Color color, FillPatternType fillPattern) {
        style.setFillForegroundColor(color);
        style.setFillPattern(fillPattern);
    }

    /**
     * Resolves the color based on the workbook type.
     * @param red the red component
     * @param green the green component
     * @param blue the blue component
     * @param workbook the workbook to resolve the color for
     * @return the color
     */
    private static Color resolveColor(int red, int green, int blue, Workbook workbook) {

        if (workbook instanceof HSSFWorkbook hssfWorkbook) {
            HSSFPalette colorMap = hssfWorkbook.getCustomPalette();
            return colorMap.findSimilarColor(red, green, blue);
        }

        if (workbook instanceof XSSFWorkbook xssfWorkbook) {
            var indexedColors = xssfWorkbook.getStylesSource().getIndexedColors();
            return new XSSFColor(new java.awt.Color(red, green, blue), indexedColors);
        }

        if (workbook instanceof SXSSFWorkbook sxssfWorkbook) {
            var indexedColors = sxssfWorkbook.getXSSFWorkbook().getStylesSource().getIndexedColors();
            return new XSSFColor(new java.awt.Color(red, green, blue), indexedColors);
        }

        throw new UnsupportedOperationException("Workbook not supported for color resolution: " + workbook);
    }

    /**
     * Represents a holder for the workbook and its style and font creation.
     */
    @Data
    @RequiredArgsConstructor
    public static class WorkbookHolder {

        private final Workbook workbook;
        private Font font;
        private CellStyle style;

        /**
         * Gets the font for the workbook, if it doesn't exist, it will be created.
         * @return the font
         */
        private Font getFont() {

            if (font == null) {
                font = workbook.createFont();
                getStyle().setFont(font);
            }

            return font;
        }

        /**
         * Gets the style for the workbook, if it doesn't exist, it will be created.
         * @return the style
         */
        private CellStyle getStyle() {
            return style == null ? workbook.createCellStyle() : style;
        }
    }

    public static void autoSizeColumns(Sheet sheet, int startColumn, int endColumn, double widthMultiplier) {
        for (int i = startColumn; i <= endColumn; i++) {
            sheet.autoSizeColumn(i);
            int width = (int) (sheet.getColumnWidth(i) * widthMultiplier);
            if (width <= 65280) { //MAX COLUMN WIDTH
                sheet.setColumnWidth(i, width);
            }
        }
    }
}
