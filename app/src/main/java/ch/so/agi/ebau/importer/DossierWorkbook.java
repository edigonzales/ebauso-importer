package ch.so.agi.ebau.importer;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class DossierWorkbook {
    private final String sheetName;
    private final List<String> headers;
    private final List<DossierEntry> entries;
    private final int idColumnIndex;

    private DossierWorkbook(String sheetName, List<String> headers, List<DossierEntry> entries, int idColumnIndex) {
        this.sheetName = sheetName;
        this.headers = headers;
        this.entries = entries;
        this.idColumnIndex = idColumnIndex;
    }

    public static DossierWorkbook read(Path workbookPath) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(Files.newInputStream(workbookPath))) {
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter(Locale.GERMANY);
            Row headerRow = sheet.getRow(sheet.getFirstRowNum());
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(formatter.formatCellValue(cell));
            }

            int idIndex = headers.indexOf("ID");
            if (idIndex < 0) {
                throw new IllegalArgumentException("dossiers.xlsx must contain an 'ID' column");
            }

            List<DossierEntry> entries = new ArrayList<>();
            for (int rowIdx = headerRow.getRowNum() + 1; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null) {
                    continue;
                }
                List<String> values = new ArrayList<>();
                for (int col = 0; col < headers.size(); col++) {
                    Cell cell = row.getCell(col);
                    values.add(cell == null ? "" : formatter.formatCellValue(cell));
                }
                String id = values.get(idIndex).trim();
                if (!id.isEmpty()) {
                    entries.add(new DossierEntry(id, values, rowIdx));
                }
            }
            return new DossierWorkbook(sheet.getSheetName(), headers, entries, idIndex);
        }
    }

    public void writeFiltered(Path target, List<DossierEntry> filteredEntries) throws IOException {
        Files.createDirectories(target.getParent());
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                headerRow.createCell(i).setCellValue(headers.get(i));
            }
            int rowIndex = 1;
            for (DossierEntry entry : filteredEntries) {
                Row row = sheet.createRow(rowIndex++);
                List<String> values = entry.values();
                for (int col = 0; col < headers.size(); col++) {
                    row.createCell(col).setCellValue(col < values.size() ? values.get(col) : "");
                }
            }
            try (OutputStream out = Files.newOutputStream(target)) {
                workbook.write(out);
            }
        }
    }

    public List<DossierEntry> entries() {
        return entries;
    }

    public Map<String, DossierEntry> entriesById() {
        Map<String, DossierEntry> map = new HashMap<>();
        for (DossierEntry entry : entries) {
            map.put(entry.id(), entry);
        }
        return map;
    }

    public int idColumnIndex() {
        return idColumnIndex;
    }

    public record DossierEntry(String id, List<String> values, int rowIndex) {
    }
}
