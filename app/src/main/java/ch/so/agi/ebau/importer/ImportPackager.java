package ch.so.agi.ebau.importer;

import ch.so.agi.ebau.importer.DossierWorkbook.DossierEntry;
import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ImportPackager {
    private static final Logger LOGGER = LoggerFactory.getLogger(ImportPackager.class);
    private static final Set<String> KNOWN_STATUSES = Set.of("SUBMITTED", "APPROVED", "REJECTED", "WRITTEN OFF", "DONE");
    private static final String UNKNOWN_STATUS = "UNKNOWN";

    private final Path rootPath;
    private final long packageSizeBytes;

    public ImportPackager(Path rootPath, long packageSizeBytes) {
        this.rootPath = rootPath;
        this.packageSizeBytes = packageSizeBytes;
    }

    public void execute(String municipality, DataType dataType, int runNumber) throws IOException {
        Path municipalityFolder = rootPath.resolve(municipality);
        Path dataFolder = municipalityFolder.resolve(dataType.dataFolderName());
        if (!Files.isDirectory(dataFolder)) {
            throw new IllegalArgumentException("Data folder not found: " + dataFolder);
        }
        Path dossierPath = dataFolder.resolve("dossiers.xlsx");
        if (!Files.exists(dossierPath)) {
            throw new IllegalArgumentException("Missing dossiers.xlsx at " + dossierPath);
        }

        LOGGER.info("Starte Verpackung für {} ({}), Lauf {}", municipality, dataType, runNumber);
        DossierWorkbook workbook = DossierWorkbook.read(dossierPath);
        Map<String, DossierEntry> entriesById = workbook.entriesById();
        List<Path> availableFolders = Files.list(dataFolder)
                .filter(Files::isDirectory)
                .filter(path -> !path.getFileName().toString().equalsIgnoreCase("Import"))
                .sorted()
                .collect(Collectors.toList());

        validateFolderCoverage(availableFolders, entriesById.keySet());

        Map<Path, Long> folderSizes = new LinkedHashMap<>();
        for (Path folder : availableFolders) {
            long size = calculateSize(folder);
            folderSizes.put(folder, size);
            LOGGER.info("Ordner {} hat Größe {} Bytes", folder.getFileName(), size);
        }

        Set<String> usedIds = new HashSet<>();
        List<PackagePlan> plans = planPackages(folderSizes, entriesById, usedIds);
        List<DossierEntry> leftoverEntries = workbook.entries().stream()
                .filter(entry -> !usedIds.contains(entry.id()))
                .toList();

        Path runFolder = municipalityFolder.resolve("Import").resolve(dataType.runFolderName(runNumber));
        Files.createDirectories(runFolder);
        PackagingStatistics statistics = new PackagingStatistics(workbook.entries().size());

        int packageIndex = 1;
        for (PackagePlan plan : plans) {
            String packageName = municipality + "_" + packageIndex++;
            createPackage(runFolder, packageName, plan, workbook, statistics);
        }

        if (!leftoverEntries.isEmpty()) {
            String packageName = municipality + "_" + packageIndex;
            PackagePlan leftoverPlan = new PackagePlan(Collections.emptyList(), leftoverEntries);
            createPackage(runFolder, packageName, leftoverPlan, workbook, statistics);
        }

        Path statsPath = runFolder.resolve("statistics.xlsx");
        statistics.write(statsPath);
        LOGGER.info("Statistik geschrieben nach {}", statsPath);
    }

    private List<PackagePlan> planPackages(Map<Path, Long> folderSizes, Map<String, DossierEntry> entriesById, Set<String> usedIds) {
        List<Path> folders = new ArrayList<>(folderSizes.keySet());
        folders.sort(Comparator.comparing(path -> path.getFileName().toString()));
        List<PackagePlan> plans = new ArrayList<>();
        List<Path> current = new ArrayList<>();
        List<DossierEntry> currentEntries = new ArrayList<>();
        long currentSize = 0;

        for (Path folder : folders) {
            long folderSize = folderSizes.get(folder);
            String folderName = folder.getFileName().toString();
            DossierEntry entry = entriesById.get(folderName);
            if (entry == null) {
                continue;
            }
            if (!current.isEmpty() && currentSize + folderSize > packageSizeBytes) {
                plans.add(new PackagePlan(List.copyOf(current), List.copyOf(currentEntries)));
                current.clear();
                currentEntries.clear();
                currentSize = 0;
            }
            current.add(folder);
            currentEntries.add(entry);
            usedIds.add(entry.id());
            currentSize += folderSize;
        }
        if (!current.isEmpty()) {
            plans.add(new PackagePlan(List.copyOf(current), List.copyOf(currentEntries)));
        }
        return plans;
    }

    private void createPackage(Path runFolder, String packageName, PackagePlan plan, DossierWorkbook workbook, PackagingStatistics statistics)
            throws IOException {
        LOGGER.info("Erzeuge Paket {} mit {} Ordnern", packageName, plan.folders().size());
        Path packageFolder = runFolder.resolve(packageName);
        Files.createDirectories(packageFolder);

        Path dossierCopy = packageFolder.resolve("dossiers.xlsx");
        workbook.writeFiltered(dossierCopy, plan.entries());

        long uncompressedSum = 0L;
        for (Path folder : plan.folders()) {
            Path target = packageFolder.resolve(folder.getFileName());
            copyFolder(folder, target);
            long folderSize = calculateSize(folder);
            statistics.addAssignment(packageName, folder.getFileName().toString(), folderSize, 0);
        }

        if (plan.folders().isEmpty()) {
            for (DossierEntry entry : plan.entries()) {
                statistics.addAssignment(packageName, entry.id(), 0, 0);
            }
        }

        uncompressedSum = calculateSize(packageFolder);
        int documentCount = countDocuments(packageFolder);

        Path zipPath = runFolder.resolve(packageName + ".zip");
        zipDirectory(packageFolder, zipPath);
        long zipSize = Files.size(zipPath);
        statistics.registerZipSize(packageName, zipSize);
        statistics.registerPackageTotals(packageName, uncompressedSum, zipSize, plan.entries().size(), plan.folders().size(),
                documentCount, calculateStatusTotals(plan.entries(), workbook));
        LOGGER.info("Paket {} erstellt (ungepackt {} Bytes, gezippt {} Bytes)", packageName, uncompressedSum, zipSize);
    }

    private void validateFolderCoverage(List<Path> folders, Set<String> entryIds) {
        List<String> missingEntries = new ArrayList<>();
        for (Path folder : folders) {
            String name = folder.getFileName().toString();
            if (!entryIds.contains(name)) {
                missingEntries.add(name);
            }
        }
        if (!missingEntries.isEmpty()) {
            LOGGER.error("Ordner ohne passende Zeile in dossiers.xlsx: {}", String.join(", ", missingEntries));
        }
    }

    private long calculateSize(Path folder) throws IOException {
        try (var stream = Files.walk(folder)) {
            return stream.filter(Files::isRegularFile)
                    .mapToLong(path -> {
                        try {
                            return Files.size(path);
                        } catch (IOException e) {
                            throw new IllegalStateException("Kann Größe nicht bestimmen für " + path, e);
                        }
                    }).sum();
        }
    }

    private int countDocuments(Path folder) throws IOException {
        try (var stream = Files.walk(folder)) {
            return (int) stream.filter(Files::isRegularFile).count();
        }
    }

    private Map<String, Integer> calculateStatusTotals(List<DossierEntry> entries, DossierWorkbook workbook) {
        int statusIndex = workbook.headerIndex("STATUS");
        Map<String, Integer> totals = new LinkedHashMap<>();
        if (statusIndex < 0) {
            return totals;
        }
        for (DossierEntry entry : entries) {
            List<String> values = entry.values();
            if (statusIndex >= values.size()) {
                totals.put(UNKNOWN_STATUS, totals.getOrDefault(UNKNOWN_STATUS, 0) + 1);
                continue;
            }
            String status = values.get(statusIndex).trim().toUpperCase();
            if (status.isEmpty() || !KNOWN_STATUSES.contains(status)) {
                totals.put(UNKNOWN_STATUS, totals.getOrDefault(UNKNOWN_STATUS, 0) + 1);
                continue;
            }
            totals.put(status, totals.getOrDefault(status, 0) + 1);
        }
        return totals;
    }

    private void copyFolder(Path source, Path target) throws IOException {
        Files.walkFileTree(source, new SimpleFileVisitor<>() {
            @Override
            public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
                Path relative = source.relativize(dir);
                Files.createDirectories(target.resolve(relative));
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                Path relative = source.relativize(file);
                Files.copy(file, target.resolve(relative));
                return FileVisitResult.CONTINUE;
            }
        });
    }

    private void zipDirectory(Path sourceDir, Path zipFile) throws IOException {
        try (ZipOutputStream zipOutputStream = new ZipOutputStream(Files.newOutputStream(zipFile))) {
            Files.walk(sourceDir)
                    .filter(path -> !Files.isDirectory(path))
                    .forEach(path -> {
                        String entryName = sourceDir.relativize(path).toString().replace('\\', '/');
                        ZipEntry entry = new ZipEntry(entryName);
                        try {
                            zipOutputStream.putNextEntry(entry);
                            Files.copy(path, zipOutputStream);
                            zipOutputStream.closeEntry();
                        } catch (IOException e) {
                            throw new IllegalStateException("Fehler beim Zippen von " + path, e);
                        }
                    });
        }
    }

    private record PackagePlan(List<Path> folders, List<DossierEntry> entries) {
    }
}
