package ch.so.agi.ebau.importer;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class PackagingStatistics {
    private static final String[] STATUS_HEADERS = { "SUBMITTED", "APPROVED", "REJECTED", "WRITTEN OFF", "DONE" };
    private final int originalRowCount;
    private final List<Assignment> assignments = new ArrayList<>();
    private final Map<String, Long> zipSizes = new HashMap<>();
    private final Map<String, Integer> dossierCounts = new HashMap<>();
    private final Map<String, Integer> folderCounts = new HashMap<>();
    private final Map<String, Integer> documentCounts = new HashMap<>();
    private final Map<String, Map<String, Integer>> statusCounts = new HashMap<>();
    private final Map<String, Long> uncompressedPackageSizes = new HashMap<>();

    public PackagingStatistics(int originalRowCount) {
        this.originalRowCount = originalRowCount;
    }

    public void addAssignment(String packageName, String folderId, long uncompressedBytes, long zipBytes) {
        assignments.add(new Assignment(packageName, folderId, uncompressedBytes, zipBytes));
    }

    public void registerZipSize(String packageName, long zipBytes) {
        zipSizes.put(packageName, zipBytes);
        for (int i = 0; i < assignments.size(); i++) {
            Assignment assignment = assignments.get(i);
            if (assignment.packageName().equals(packageName)) {
                assignments.set(i, new Assignment(assignment.packageName(), assignment.folderId(), assignment.uncompressedBytes(), zipBytes));
            }
        }
    }

    public void registerPackageTotals(String packageName, long uncompressedBytes, long zipBytes, int dossierCount,
            int folderCount, int documentCount, Map<String, Integer> statusTotals) {
        uncompressedPackageSizes.put(packageName, uncompressedBytes);
        zipSizes.put(packageName, zipBytes);
        dossierCounts.put(packageName, dossierCount);
        folderCounts.put(packageName, folderCount);
        documentCounts.put(packageName, documentCount);
        Map<String, Integer> totals = new HashMap<>();
        for (String status : STATUS_HEADERS) {
            totals.put(status, statusTotals.getOrDefault(status, 0));
        }
        statusCounts.put(packageName, totals);
    }

    public void write(Path target) throws IOException {
        Files.createDirectories(target.getParent());
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet assignmentSheet = workbook.createSheet("Packages");
            Row header = assignmentSheet.createRow(0);
            header.createCell(0).setCellValue("Package");
            header.createCell(1).setCellValue("FolderID");
            header.createCell(2).setCellValue("UncompressedBytes");
            header.createCell(3).setCellValue("ZipBytes");
            int rowIndex = 1;
            for (Assignment assignment : assignments) {
                Row row = assignmentSheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(assignment.packageName());
                row.createCell(1).setCellValue(assignment.folderId());
                row.createCell(2).setCellValue(assignment.uncompressedBytes());
                row.createCell(3).setCellValue(zipSizes.getOrDefault(assignment.packageName(), assignment.zipBytes()));
            }

            Sheet overview = workbook.createSheet("Dossiers");
            Row overviewHeader = overview.createRow(0);
            overviewHeader.createCell(0).setCellValue("OriginalRows");
            overviewHeader.createCell(1).setCellValue("PackagedRows");
            overviewHeader.createCell(2).setCellValue("Packages");
            Row dataRow = overview.createRow(1);
            dataRow.createCell(0).setCellValue(originalRowCount);
            dataRow.createCell(1).setCellValue(calculatePackagedRows());
            dataRow.createCell(2).setCellValue(zipSizes.size());

            Sheet details = workbook.createSheet("Details");
            createDetailsHeader(details);
            int detailRowIndex = 1;
            List<String> sortedPackages = zipSizes.keySet().stream().sorted().toList();
            for (String packageName : sortedPackages) {
                Row row = details.createRow(detailRowIndex);
                row.createCell(0).setCellValue(packageName);
                row.createCell(1).setCellValue(uncompressedPackageSizes.getOrDefault(packageName, 0L));
                row.createCell(2).setCellValue(zipSizes.getOrDefault(packageName, 0L));
                row.createCell(3).setCellValue(dossierCounts.getOrDefault(packageName, 0));
                row.createCell(4).setCellValue(folderCounts.getOrDefault(packageName, 0));
                row.createCell(5).setCellValue(documentCounts.getOrDefault(packageName, 0));
                Map<String, Integer> totals = statusCounts.getOrDefault(packageName, Map.of());
                int statusStartIndex = 6;
                for (int i = 0; i < STATUS_HEADERS.length; i++) {
                    row.createCell(statusStartIndex + i).setCellValue(totals.getOrDefault(STATUS_HEADERS[i], 0));
                }
                int totalCellIndex = statusStartIndex + STATUS_HEADERS.length;
                int excelRowNumber = detailRowIndex + 1;
                row.createCell(totalCellIndex)
                        .setCellFormula(String.format("SUM(%s%d:%s%d)", columnName(statusStartIndex),
                                excelRowNumber, columnName(totalCellIndex - 1), excelRowNumber));
                detailRowIndex++;
            }

            try (OutputStream out = Files.newOutputStream(target)) {
                workbook.write(out);
            }
        }
    }

    private int calculatePackagedRows() {
        return (int) assignments.stream().map(Assignment::folderId).distinct().count();
    }

    private void createDetailsHeader(Sheet sheet) {
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Package");
        header.createCell(1).setCellValue("Size Unzipped [Byte]");
        header.createCell(2).setCellValue("Size Zipped [Byte]");
        header.createCell(3).setCellValue("Anzahl Dossiers in dossiers.xlsx");
        header.createCell(4).setCellValue("Anzahl Dokumentenordner (1. Ebene)");
        header.createCell(5).setCellValue("Anzahl Dokumente");
        int statusStartIndex = 6;
        for (int i = 0; i < STATUS_HEADERS.length; i++) {
            header.createCell(statusStartIndex + i).setCellValue(STATUS_HEADERS[i]);
        }
        header.createCell(statusStartIndex + STATUS_HEADERS.length).setCellValue("Total");
    }

    private String columnName(int columnIndex) {
        int index = columnIndex;
        StringBuilder builder = new StringBuilder();
        while (index >= 0) {
            int remainder = index % 26;
            builder.insert(0, (char) ('A' + remainder));
            index = (index / 26) - 1;
        }
        return builder.toString();
    }

    private record Assignment(String packageName, String folderId, long uncompressedBytes, long zipBytes) {
    }
}
