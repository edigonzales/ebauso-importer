package ch.so.agi.ebau.importer;

import java.util.Locale;

public enum DataType {
    TEST("Testdaten", "Testlauf"),
    PRODUCTION("Produktivdaten", "Produktivlauf");

    private final String dataFolderName;
    private final String runFolderPrefix;

    DataType(String dataFolderName, String runFolderPrefix) {
        this.dataFolderName = dataFolderName;
        this.runFolderPrefix = runFolderPrefix;
    }

    public String dataFolderName() {
        return dataFolderName;
    }

    public String runFolderName(int runNumber) {
        return runFolderPrefix + "_" + runNumber;
    }

    public static DataType fromValue(String value) {
        String normalized = value.toLowerCase(Locale.ROOT);
        return switch (normalized) {
            case "test", "testdaten", "t" -> TEST;
            case "prod", "production", "produktiv", "produkt" , "produktivdaten", "p" -> PRODUCTION;
            default -> throw new IllegalArgumentException("Unsupported Datentyp: " + value);
        };
    }
}
