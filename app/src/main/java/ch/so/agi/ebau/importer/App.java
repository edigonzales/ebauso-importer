package ch.so.agi.ebau.importer;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public final class App {
    private static final Logger LOGGER = LoggerFactory.getLogger(App.class);

    private App() {
    }

    public static void main(String[] args) {
        try {
            CommandLineArguments arguments = CommandLineArguments.parse(args);
            ImportPackager packager = new ImportPackager(arguments.rootPath(), arguments.packageSizeBytes());
            packager.execute(arguments.municipality(), arguments.dataType(), arguments.runNumber());
        } catch (IllegalArgumentException ex) {
            LOGGER.error("Invalid arguments: {}", ex.getMessage());
            System.exit(1);
        } catch (Exception ex) {
            LOGGER.error("Failed to run importer", ex);
            System.exit(2);
        }
    }
}
