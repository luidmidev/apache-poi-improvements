package io.github.luidmidev.apache.poi;

import io.github.luidmidev.apache.poi.model.ReportFile;
import io.github.luidmidev.apache.poi.model.WorkbookType;
import io.github.luidmidev.apache.poi.utils.WorkbookUtils;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

class TestWorkbook {


    @Test
    void from() throws Exception {
        var timeMillis = System.currentTimeMillis();
        var persons = new ArrayList<Person>();
        for (var i = 0; i < 10000; i++) {
            persons.add(new Person("Juan", 12, "New York", "111@aa.com", "123456", i + " Doe"));
        }

        System.out.println("Time on generate list: " + (System.currentTimeMillis() - timeMillis));


        var style = CellStylizer.init()
                .withFontColor(IndexedColors.RED)
                .withFontBold()
                .withFontSize(8)
                .withFontName("Arial")
                .foregroundColor(IndexedColors.YELLOW, FillPatternType.SOLID_FOREGROUND)
                .withAlignment(HorizontalAlignment.CENTER)
                .withAlignment(VerticalAlignment.CENTER)
                .allBorders(BorderStyle.THIN)
                .withWrapText();

        var currentTimeMillis = System.currentTimeMillis();

        var workbookForReport = WorkbookManager.fromItems(persons)
                .withColumn("Name", Person::name, style)
                .withColumn("Complete Name", person -> person.name() + " \n" + person.lastName(), style)
                .withColumn("Age", Person::age, style)
                .withColumn("Address", Person::address, style)
                .withColumn("Email", Person::email, style)
                .withColumn("Phone", Person::phone, style)
                .withHeaderStyle(style)
                .configureSheet(sheet -> {
                    sheet.autoSizeColumn(0);
                    sheet.autoSizeColumn(1);
                    sheet.autoSizeColumn(2);
                    sheet.autoSizeColumn(3);
                    sheet.autoSizeColumn(4);
                    sheet.autoSizeColumn(5);
                })
                .forEachRow((row, person) -> WorkbookUtils.adjustRowHeightByLines(row))
                .build();

        try (workbookForReport) {

            var spreadsheet = workbookForReport.getSpreadsheet("Persons");

            System.out.println("Spreadsheet: " + spreadsheet.getFilename());

            save(spreadsheet);
        }

        System.out.println("Time on generate report: " + (System.currentTimeMillis() - currentTimeMillis));
        Assertions.assertTrue(true);
    }


    @Test
    void from2() throws Exception {

        var currentTimeMillis = System.currentTimeMillis();
        var persons = new ArrayList<Person>();
        for (var i = 0; i < 10000; i++) {
            persons.add(new Person("Juan", 12, "New York", "111@aa.com", "123456", i + " Doe"));
        }
        System.out.println("Time on generate list: " + (System.currentTimeMillis() - currentTimeMillis));

        var b = System.currentTimeMillis();
        var templateRource = getClass().getClassLoader().getResourceAsStream("sample_with_header_and_footer.xlsx");
        System.out.println("Time on load template: " + (System.currentTimeMillis() - b));

        var c = System.currentTimeMillis();
        var wookbook = new WorkbookManager(templateRource, WorkbookType.XLSX);
        System.out.println("Time on create workbook: " + (System.currentTimeMillis() - c));

        var d = System.currentTimeMillis();
        WorkbookManager workbookForReport = WorkbookManager.fromItems(persons, wookbook, 3)
                .configureSheet(sheet -> sheet.setColumnWidth(0, 5000))
                .withColumn("Name", Person::name)
                .withColumn("Complete Name", person -> person.name() + " " + person.lastName())
                .withColumn("Age", Person::age)
                .withColumn("Address", Person::address)
                .withColumn("Email", Person::email)
                .withColumn("Phone", Person::phone)
                .onProgress((current, total) -> System.out.println("Progress: " + current + " of " + total))
                .build();

        try (workbookForReport) {

            var spreadsheet = workbookForReport.getSpreadsheet("Persons");

            System.out.println("Spreadsheet: " + spreadsheet.getFilename());

            save(spreadsheet);
        }

        System.out.println("Time on generate report: " + (System.currentTimeMillis() - d));
        Assertions.assertTrue(true);
    }


    private void save(ReportFile report) throws FileNotFoundException {

        var file = new java.io.File(report.getFilename());
        try (var fos = new java.io.FileOutputStream(file)) {
            fos.write(report.getContent());
        } catch (IOException e) {
            System.err.println("Error saving file: " + e.getMessage());
        }


    }


    public record Person(String name, int age, String address, String email, String phone, String lastName) {
    }
}
