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
    private final int originalRowCount;
    private final List<Assignment> assignments = new ArrayList<>();
    private final Map<String, Long> zipSizes = new HashMap<>();

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

            try (OutputStream out = Files.newOutputStream(target)) {
                workbook.write(out);
            }
        }
    }

    private int calculatePackagedRows() {
        return (int) assignments.stream().map(Assignment::folderId).distinct().count();
    }

    private record Assignment(String packageName, String folderId, long uncompressedBytes, long zipBytes) {
    }
}
