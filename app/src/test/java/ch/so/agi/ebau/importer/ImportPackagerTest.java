package ch.so.agi.ebau.importer;

import static org.assertj.core.api.Assertions.assertThat;

import ch.so.agi.ebau.importer.DossierWorkbook.DossierEntry;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class ImportPackagerTest {

    @Test
    void packagesDataAndCreatesLeftoverZip(@TempDir Path tempDir) throws Exception {
        Path municipality = tempDir.resolve("Biberist");
        Path dataFolder = municipality.resolve("Testdaten");
        Files.createDirectories(dataFolder);

        Path dossier = dataFolder.resolve("dossiers.xlsx");
        writeDossier(dossier, List.of("A", "B", "C"));

        createFolderWithFile(dataFolder.resolve("A"), "a.txt", "hello");
        createFolderWithFile(dataFolder.resolve("B"), "b.txt", "this folder should be placed alone");

        ImportPackager packager = new ImportPackager(tempDir, 20); // force small packages
        packager.execute("Biberist", DataType.TEST, 1);

        Path runFolder = municipality.resolve(Path.of("Import", "Testlauf_1"));
        assertThat(runFolder).isDirectoryContaining(path -> path.getFileName().toString().endsWith(".zip"));

        Path package1 = runFolder.resolve("Biberist_1");
        Path package2 = runFolder.resolve("Biberist_2");
        Path leftover = runFolder.resolve("Biberist_3");
        assertThat(package1).exists();
        assertThat(package2).exists();
        assertThat(leftover).exists();

        List<DossierEntry> entries1 = DossierWorkbook.read(package1.resolve("dossiers.xlsx")).entries();
        List<DossierEntry> entries2 = DossierWorkbook.read(package2.resolve("dossiers.xlsx")).entries();
        List<DossierEntry> leftoverEntries = DossierWorkbook.read(leftover.resolve("dossiers.xlsx")).entries();

        assertThat(entries1).extracting(DossierEntry::id).containsExactly("A");
        assertThat(entries2).extracting(DossierEntry::id).containsExactly("B");
        assertThat(leftoverEntries).extracting(DossierEntry::id).containsExactly("C");

        assertThat(package1.resolve("A")).isDirectory();
        assertThat(package2.resolve("B")).isDirectory();
        try (var stream = Files.list(leftover)) {
            assertThat(stream.filter(Files::isDirectory)).isEmpty();
        }

        Path stats = runFolder.resolve("statistics.xlsx");
        try (var workbook = WorkbookFactory.create(Files.newInputStream(stats))) {
            var sheet = workbook.getSheet("Dossiers");
            assertThat(sheet.getRow(1).getCell(0).getNumericCellValue()).isEqualTo(3);
            assertThat(sheet.getRow(1).getCell(1).getNumericCellValue()).isEqualTo(3);
            assertThat(sheet.getRow(1).getCell(2).getNumericCellValue()).isEqualTo(3);
        }
    }

    private void writeDossier(Path target, List<String> ids) throws IOException {
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            var sheet = workbook.createSheet("dossiers");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("ID");
            header.createCell(1).setCellValue("Name");
            int row = 1;
            for (String id : ids) {
                var r = sheet.createRow(row++);
                r.createCell(0).setCellValue(id);
                r.createCell(1).setCellValue("Dossier " + id);
            }
            Files.createDirectories(target.getParent());
            try (var out = Files.newOutputStream(target)) {
                workbook.write(out);
            }
        }
    }

    private void createFolderWithFile(Path folder, String filename, String content) throws IOException {
        Files.createDirectories(folder);
        Files.writeString(folder.resolve(filename), content);
    }
}
