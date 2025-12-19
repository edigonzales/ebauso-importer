package ch.so.agi.ebau.importer;

import java.nio.file.Path;
import java.nio.file.Paths;

final class CommandLineArguments {
    private final String municipality;
    private final DataType dataType;
    private final int runNumber;
    private final Path rootPath;
    private final long packageSizeBytes;

    private CommandLineArguments(String municipality, DataType dataType, int runNumber, Path rootPath, long packageSizeBytes) {
        this.municipality = municipality;
        this.dataType = dataType;
        this.runNumber = runNumber;
        this.rootPath = rootPath;
        this.packageSizeBytes = packageSizeBytes;
    }

    public static CommandLineArguments parse(String[] args) {
        if (args.length < 3) {
            throw new IllegalArgumentException("Usage: <Gemeinde> <Datentyp> <Laufnummer> [--root=/pfad] [--packageSizeMb=900]");
        }

        String municipality = args[0];
        DataType dataType = DataType.fromValue(args[1]);
        int runNumber = Integer.parseInt(args[2]);
        Path root = Paths.get(".");
        long packageSizeMb = 900;

        for (int i = 3; i < args.length; i++) {
            String arg = args[i];
            if (arg.startsWith("--root=")) {
                root = Paths.get(arg.substring("--root=".length()));
            } else if (arg.startsWith("--packageSizeMb=")) {
                packageSizeMb = Long.parseLong(arg.substring("--packageSizeMb=".length()));
            }
        }

        return new CommandLineArguments(municipality, dataType, runNumber, root.toAbsolutePath().normalize(), packageSizeMb * 1024 * 1024);
    }

    public String municipality() {
        return municipality;
    }

    public DataType dataType() {
        return dataType;
    }

    public int runNumber() {
        return runNumber;
    }

    public Path rootPath() {
        return rootPath;
    }

    public long packageSizeBytes() {
        return packageSizeBytes;
    }
}
