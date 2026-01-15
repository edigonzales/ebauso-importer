package ch.so.agi.ebau.importer;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class DossierWorkbook {
    private static final Set<String> COORDINATE_HEADERS = Set.of("COORDINATE-N", "COORDINATE-E");
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
            Locale swissLocale = Locale.forLanguageTag("de-CH");
            DataFormatter formatter = new DataFormatter(swissLocale);
            DateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy", swissLocale);
            DecimalFormatSymbols coordinateSymbols = new DecimalFormatSymbols(swissLocale);
            coordinateSymbols.setDecimalSeparator('.');
            DecimalFormat coordinateFormat = new DecimalFormat("0.00", coordinateSymbols);
            coordinateFormat.setGroupingUsed(false);
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
                    String header = headers.get(col);
                    values.add(formatCellValue(cell, header, formatter, dateFormat, coordinateFormat));
                }
                String id = values.get(idIndex).trim();
                if (!id.isEmpty()) {
                    entries.add(new DossierEntry(id, values, rowIdx));
                }
            }
            return new DossierWorkbook(sheet.getSheetName(), headers, entries, idIndex);
        }
    }

    private static String formatCellValue(Cell cell, String header, DataFormatter formatter, DateFormat dateFormat,
            DecimalFormat coordinateFormat) {
        if (cell == null) {
            return "";
        }
        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA) {
            cellType = cell.getCachedFormulaResultType();
            return switch (cellType) {
                case STRING -> cell.getStringCellValue();
                case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
                case NUMERIC -> formatNumericValue(cell, cell.getNumericCellValue(), header, dateFormat, coordinateFormat,
                        formatter);
                case BLANK -> "";
                default -> formatter.formatCellValue(cell);
            };
        }
        if (cellType == CellType.NUMERIC) {
            return formatNumericValue(cell, cell.getNumericCellValue(), header, dateFormat, coordinateFormat, formatter);
        }
        return formatter.formatCellValue(cell);
    }

    private static String formatNumericValue(Cell cell, double numericValue, String header, DateFormat dateFormat,
            DecimalFormat coordinateFormat, DataFormatter formatter) {
        if (DateUtil.isCellDateFormatted(cell)) {
            return dateFormat.format(cell.getDateCellValue());
        }
        if (COORDINATE_HEADERS.contains(header)) {
            return coordinateFormat.format(numericValue);
        }
        return formatter.formatCellValue(cell);
    }

    public void writeFiltered(Path target, List<DossierEntry> filteredEntries) throws IOException {
        Files.createDirectories(target.getParent());
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);
            var coordinateStyle = workbook.createCellStyle();
            coordinateStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                headerRow.createCell(i).setCellValue(headers.get(i));
            }
            int rowIndex = 1;
            for (DossierEntry entry : filteredEntries) {
                Row row = sheet.createRow(rowIndex++);
                List<String> values = entry.values();
                for (int col = 0; col < headers.size(); col++) {
                    String value = col < values.size() ? values.get(col) : "";
                    Cell cell = row.createCell(col);
                    if (!value.isBlank() && COORDINATE_HEADERS.contains(headers.get(col))) {
                        try {
                            cell.setCellValue(Double.parseDouble(value));
                            cell.setCellStyle(coordinateStyle);
                        } catch (NumberFormatException ex) {
                            cell.setCellValue(value);
                        }
                    } else {
                        cell.setCellValue(value);
                    }
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
