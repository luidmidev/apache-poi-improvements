package io.github.luidmidev.apache.poi.model;


public enum WorkbookType {

    XLSX("xlsx"),
    XLS("xls"),
    XLSM("xlsm");

    private final String extension;

    WorkbookType(String extension) {
        this.extension = extension;
    }

    public final String getExtension() {
        return extension;
    }

    public static WorkbookType fromExtension(String extension) {
        for (var type : values()) {
            if (type.extension.equals(extension)) return type;
        }
        throw new IllegalArgumentException("Unknown extension: " + extension);
    }

    public String getMediaType() {
        return switch (this) {
            case XLSX -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            case XLS -> "application/vnd.ms-excel";
            case XLSM -> "application/vnd.ms-excel.sheet.macroEnabled.12";
        };
    }


    public static WorkbookType fromFilename(String filename) {
        return fromExtension(filename.substring(filename.lastIndexOf('.') + 1));
    }
}
