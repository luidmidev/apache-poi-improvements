package io.github.luidmidev.apache.poi.model;


import lombok.Getter;
import lombok.Setter;

/**
 * Represents a spreadsheet file
 */
@Setter
@Getter
public class SpreadSheetFile {
    private String filename;
    private WorkbookType type;
    private byte[] content;
}
