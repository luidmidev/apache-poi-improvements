package io.github.luidmidev.apache.poi.model;


import lombok.RequiredArgsConstructor;

/**
 * Represents a type of workbook
 */
@RequiredArgsConstructor
public enum WorkbookType {

    /**
     * Excel workbook
     * @see <a href="https://en.wikipedia.org/wiki/Office_Open_XML">Office Open XML</a>
     */
    XLSX("xlsx"),
    /**
     * Excel 97-2003 workbook
     */
    XLS("xls"),
    /**
     * Excel with macros support
     */
    XLSM("xlsm");

    private final String extension;

    /**
     * Returns the extension of the workbook
     * @return the extension of the workbook
     */
    public final String getExtension() {
        return extension;
    }

    /**
     * Returns the workbook type from the extension
     * @param extension the extension
     * @return the workbook type
     * @throws IllegalArgumentException if the extension is unknown
     */
    public static WorkbookType fromExtension(String extension) {
        for (var type : values()) {
            if (type.extension.equals(extension)) return type;
        }
        throw new IllegalArgumentException("Unknown extension: " + extension);
    }

    /**
     * Returns the workbook type from the filename
     * @param filename the filename
     * @return the workbook type
     */
    public static WorkbookType fromFilename(String filename) {
        return fromExtension(filename.substring(filename.lastIndexOf('.') + 1));
    }
}
