package io.github.luidmidev.apache.poi.model;


import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class ReportFile {
    private String filename;
    private String mediaType;
    private byte[] content;
}
