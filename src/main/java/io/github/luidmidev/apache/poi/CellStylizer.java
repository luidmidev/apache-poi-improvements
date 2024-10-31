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

@Log4j2
public class CellStylizer {

    private final List<WorkbookHolderConsumer> consumers;

    private CellStylizer() {
        this.consumers = new ArrayList<>();
    }

    public static CellStylizer init() {
        return new CellStylizer();
    }

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

    public CellStylizer allBorders(BorderStyle borderStyle) {
        consumers.add(holder -> setBorders(holder.getStyle(), borderStyle, borderStyle, borderStyle, borderStyle));
        return this;
    }

    public CellStylizer onlyBorders(BorderStyle top, BorderStyle right, BorderStyle bottom, BorderStyle left) {
        consumers.add(holder -> setBorders(holder.getStyle(), top, right, bottom, left));
        return this;
    }

    public CellStylizer center() {
        consumers.add(holder -> setCenter(holder.getStyle()));
        return this;
    }

    public CellStylizer foregroundColor(IndexedColors indexedColors) {
        return foregroundColor(indexedColors.getIndex());
    }

    public CellStylizer foregroundColor(short index) {
        return foregroundColor(index, FillPatternType.SOLID_FOREGROUND);
    }

    public CellStylizer foregroundColor(IndexedColors indexedColors, FillPatternType fillPattern) {
        return foregroundColor(indexedColors.getIndex(), fillPattern);
    }

    public CellStylizer foregroundColor(int red, int green, int blue, FillPatternType fillPattern) {
        consumers.add(holder -> setForeground(holder.getStyle(), resolveColor(red, green, blue, holder.getWorkbook()), fillPattern));
        return this;
    }

    public CellStylizer foregroundColor(int red, int green, int blue) {
        return foregroundColor(red, green, blue, FillPatternType.SOLID_FOREGROUND);
    }

    public CellStylizer foregroundColor(short index, FillPatternType fillPattern) {
        consumers.add(holder -> setForeground(holder.getStyle(), index, fillPattern));
        return this;
    }

    public CellStylizer foregroundColor(Color color, FillPatternType fillPattern) {
        consumers.add(holder -> setForeground(holder.getStyle(), color, fillPattern));
        return this;
    }

    public CellStylizer fontColor(IndexedColors indexedColors) {
        return fontColor(indexedColors.getIndex());
    }

    public CellStylizer fontColor(short index) {
        consumers.add(holder -> holder.getFont().setColor(index));
        return this;
    }

    public CellStylizer fontBold() {
        consumers.add(holder -> holder.getFont().setBold(true));
        return this;
    }

    public CellStylizer fontSize(int fontSize) {
        consumers.add(holder -> holder.getFont().setFontHeightInPoints((short) fontSize));
        return this;
    }

    public CellStylizer fontName(String fontName) {
        consumers.add(holder -> holder.getFont().setFontName(fontName));
        return this;
    }

    public CellStylizer alignment(HorizontalAlignment horizontalAlignment) {
        consumers.add(holder -> holder.getStyle().setAlignment(horizontalAlignment));
        return this;
    }

    public CellStylizer alignment(VerticalAlignment verticalAlignment) {
        consumers.add(holder -> holder.getStyle().setVerticalAlignment(verticalAlignment));
        return this;
    }

    public CellStylizer wrapText() {
        consumers.add(holder -> holder.getStyle().setWrapText(true));
        return this;
    }

    private static void setBorders(CellStyle style, BorderStyle top, BorderStyle right, BorderStyle bottom, BorderStyle left) {
        style.setBorderTop(top == null ? BorderStyle.NONE : top);
        style.setBorderRight(right == null ? BorderStyle.NONE : right);
        style.setBorderBottom(bottom == null ? BorderStyle.NONE : bottom);
        style.setBorderLeft(left == null ? BorderStyle.NONE : left);
    }

    private static void setCenter(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    }


    private static void setForeground(CellStyle style, short color, FillPatternType fillPattern) {
        style.setFillForegroundColor(color);
        style.setFillPattern(fillPattern);
    }

    private static void setForeground(CellStyle style, Color color, FillPatternType fillPattern) {
        style.setFillForegroundColor(color);
        style.setFillPattern(fillPattern);
    }

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

    @Data
    @RequiredArgsConstructor
    public static class WorkbookHolder {

        private final Workbook workbook;

        private Font font;
        private CellStyle style;

        private Font getFont() {

            if (font == null) {
                font = workbook.createFont();
                getStyle().setFont(font);
            }

            return font;
        }

        private CellStyle getStyle() {
            return style == null ? workbook.createCellStyle() : style;
        }
    }
}
