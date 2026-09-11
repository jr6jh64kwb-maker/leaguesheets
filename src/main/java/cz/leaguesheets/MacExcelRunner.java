package cz.leaguesheets;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.UUID;
import java.util.function.Consumer;

/** Native Excel rendering on macOS; all changes are confined to a temporary workbook. */
final class MacExcelRunner {
    private MacExcelRunner() {
    }

    static void run(Path workbook, LeagueKind league, boolean makePdf, Path output,
                    Consumer<ProgressUpdate> progress) throws IOException, InterruptedException {
        if (makePdf && output == null) {
            throw new IOException("Chybí cesta pro PDF.");
        }
        Path directory = Files.createTempDirectory("league-sheets-mac-");
        try {
            String name = workbook.getFileName().toString();
            String extension = name.substring(name.lastIndexOf('.'));
            // Excel identifies open workbooks by basename, even when folders differ.
            String suffix = UUID.randomUUID().toString().substring(0, 8);
            Path source = directory.resolve("source-" + suffix + extension);
            Files.copy(workbook, source);
            Path converted = directory.resolve("converted-" + suffix + ".xlsx");
            Path prepared = directory.resolve("prepared-" + suffix + ".xlsx");
            Path pdf = directory.resolve("result.pdf");
            progress.accept(new ProgressUpdate(18, "Spouštím Excel..."));
            execute(directory, "convert", CONVERT, List.of(source.toString(), converted.toString()));
            progress.accept(new ProgressUpdate(40, "Připravuji kola a nastavení stránek..."));
            XlsxPrintLayout.Selection selection = XlsxPrintLayout.prepare(converted, prepared, league);
            progress.accept(new ProgressUpdate(75, makePdf ? "Vytvářím PDF..." : "Odesílám na tiskárnu..."));
            execute(directory, "export", EXPORT, List.of(prepared.toString(), pdf.toString(),
                    makePdf ? "pdf" : "print", selection.scheduleName(), Integer.toString(selection.rounds().size())));
            if (makePdf) {
                if (!Files.isRegularFile(pdf) || Files.size(pdf) == 0) {
                    throw new IOException("Excel nevytvořil PDF.");
                }
                // Keep an existing output intact until Excel has successfully finished.
                Files.copy(pdf, output, StandardCopyOption.REPLACE_EXISTING);
            }
            progress.accept(new ProgressUpdate(100, "Hotovo"));
        } finally {
            try (var paths = Files.walk(directory)) {
                for (Path path : paths.sorted(Comparator.reverseOrder()).toList()) {
                    Files.deleteIfExists(path);
                }
            } catch (IOException ignored) {
                // A cleanup failure must not obscure the original Excel error.
            }
        }
    }

    private static void execute(Path directory, String name, String script, List<String> arguments)
            throws IOException, InterruptedException {
        Path scriptFile = directory.resolve(name + ".applescript");
        Path logFile = directory.resolve(name + ".log");
        Files.writeString(scriptFile, script, StandardCharsets.UTF_8);
        List<String> command = new ArrayList<>(List.of("/usr/bin/osascript", scriptFile.toString()));
        command.addAll(arguments);
        Process process = new ProcessBuilder(command).redirectErrorStream(true)
                .redirectOutput(logFile.toFile()).start();
        int code;
        try {
            code = process.waitFor();
        } catch (InterruptedException e) {
            process.destroy();
            throw e;
        }
        if (code != 0) {
            String message = Files.readString(logFile, StandardCharsets.UTF_8).trim();
            if (message.contains("-1743")) {
                throw new IOException("Povol ovládání Microsoft Excelu v Nastavení systému → Soukromí a zabezpečení → "
                        + "Automatizace pro aplikaci, ze které program spouštíš (např. IntelliJ IDEA).");
            }
            throw new IOException("Zpracování v Excelu na macOS selhalo: " + message);
        }
    }

    private static final String CONVERT = """
            on run argv
              with timeout of 600 seconds
                tell application "Microsoft Excel"
                  set previousAlerts to display alerts
                  set previousSecurity to automation security
                  set wb to missing value
                  set stage to "Otevírání sešitu"
                  try
                    set display alerts to false
                    set automation security to msoAutomationSecurityForceDisable
                    open workbook workbook file name (item 1 of argv) update links do not update links read only true
                    set wb to active workbook
                    set stage to "Ukládání výstupu"
                    save workbook as wb filename (item 2 of argv) file format Excel XML file format
                    set wb to active workbook
                    set stage to "Zavírání pracovní kopie"
                    close wb saving no
                    set wb to missing value
                    set display alerts to previousAlerts
                    set automation security to previousSecurity
                  on error msg number n
                    if wb is not missing value then
                      try
                        close wb saving no
                      end try
                    end if
                    set display alerts to previousAlerts
                    set automation security to previousSecurity
                    error (stage & ": " & msg) number n
                  end try
                end tell
              end timeout
            end run
            """;

    private static final String EXPORT = """
            on run argv
              with timeout of 600 seconds
                tell application "Microsoft Excel"
                  set previousAlerts to display alerts
                  set previousSecurity to automation security
                  set wb to missing value
                  set stage to "Otevírání sešitu"
                  try
                    set display alerts to false
                    set automation security to msoAutomationSecurityForceDisable
                    repeat with attempt from 1 to 4
                      try
                        open workbook workbook file name (item 1 of argv) update links do not update links read only true
                        set wb to active workbook
                        exit repeat
                      on error openMessage number openNumber
                        if attempt is 4 or (openNumber is not -50 and openNumber is not -1728) then error openMessage number openNumber
                        delay 1
                      end try
                    end repeat
                    set stage to "Nastavení stránek"
                    repeat with sheetIndex from 1 to ((item 5 of argv) as integer)
                      set ws to worksheet sheetIndex of wb
                      set stage to "Nastavení listu " & (name of ws)
                      set areaText to print area of page setup object of ws
                      set areaWidth to width of range areaText of ws
                      set areaHeight to height of range areaText of ws
                      set pageScale to 95
                      -- Font metrics differ on macOS. Keep 95% unless the form would spill onto another page.
                      -- Printable A4 size in points after two 0.64 cm margins, rounded down.
                      set widthScale to round (558.7 * 100 / areaWidth) rounding down
                      set heightScale to round (805.6 * 100 / areaHeight) rounding down
                      if widthScale < pageScale then set pageScale to widthScale
                      if heightScale < pageScale then set pageScale to heightScale
                      tell page setup object of ws
                        set page orientation to portrait
                        set zoom to pageScale
                      end tell
                    end repeat
                    tell page setup object of worksheet (item 4 of argv) of wb
                      set page orientation to landscape
                      set zoom to false
                      set fit to pages wide to 1
                      set fit to pages tall to 1
                    end tell
                    if item 3 of argv is "pdf" then
                      set stage to "Ukládání výstupu"
                      save workbook as wb filename (item 2 of argv) file format PDF file format
                      set wb to active workbook
                    else
                      print out wb copies 1 preview false
                    end if
                    set stage to "Zavírání pracovní kopie"
                    close wb saving no
                    set wb to missing value
                    set display alerts to previousAlerts
                    set automation security to previousSecurity
                  on error msg number n
                    if wb is not missing value then
                      try
                        close wb saving no
                      end try
                    end if
                    set display alerts to previousAlerts
                    set automation security to previousSecurity
                    error (stage & ": " & msg) number n
                  end try
                end tell
              end timeout
            end run
            """;
}
