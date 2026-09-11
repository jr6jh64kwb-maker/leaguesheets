package cz.leaguesheets;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/** Dependency-free regression tests: run this class with the main classes on the classpath. */
public final class XlsxPrintLayoutTest {
    private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static final String REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void main(String[] args) throws Exception {
        Path directory = Files.createTempDirectory("league-layout-test-");
        try {
            for (LeagueKind league : LeagueKind.values()) {
                Path source = directory.resolve("source.xlsx");
                Path output = directory.resolve("output.xlsx");
                fixture(source, league, false);
                byte[] original = Files.readAllBytes(source);
                var selected = XlsxPrintLayout.prepare(source, output, league);
                int count = league == LeagueKind.SIX_TEAMS ? 5 : 7;
                check(selected.rounds().size() == count, "Round count: " + league);
                check(selected.rounds().getFirst().startsWith("1"), "Numeric ordering");
                check(selected.rounds().getLast().startsWith(Integer.toString(count)), "Last round");
                check(java.util.Arrays.equals(original, Files.readAllBytes(source)), "Source unchanged");
                String workbook = read(output, "xl/workbook.xml");
                check(workbook.contains("localSheetId=\"0\" name=\"LocalTest\""), "Local name follows reordered sheet");
                check(workbook.matches("(?s).*<sheet(?=[^>]*name=\"data\")(?=[^>]*state=\"hidden\")[^>]*/>.*"), "Data retained but hidden");
                String area = league == LeagueKind.EIGHT_TEAMS ? "$A$1:$J$35" : "$A$1:$J$28";
                check(workbook.contains(area), "Round print area");
                check(workbook.contains("$A$1:$C$11"), "Four schedules and three separator rows");
                String round = read(output, "xl/worksheets/sheet" + (count + 1) + ".xml");
                check(round.contains("paperSize=\"9\"") && round.contains("scale=\"95\""), "A4 and scale");
                check(round.contains("orientation=\"portrait\""), "Round orientation");
                check(round.contains("<f>data!A3</f><v>1</v>"), "Formula and cache preserved");
                String schedule = read(output, "xl/worksheets/sheet" + (count + 2) + ".xml");
                check(schedule.contains("orientation=\"landscape\"") && schedule.contains("fitToPage=\"1\""), "Schedule layout");
                check(occurrences(schedule, "Schedule title") == 4, "Exactly four schedules");
                check(occurrences(schedule, "Team A - Team B") == 4, "Copies retain cached formula values");
                check(!schedule.contains("<f>"), "Print copies must not shift formulas");
                check(occurrences(schedule, "<mergeCell ") == 4, "Merged headings retained");
                Path macOutput = directory.resolve("mac-output.xlsx");
                var macSelection = XlsxPrintLayout.prepareForMac(source, macOutput, league);
                check(macSelection.equals(selected), "Mac selects the same rounds and schedule");
                check(java.util.Arrays.equals(original, Files.readAllBytes(source)), "Mac source unchanged");
                for (int sheetId = 2; sheetId <= count + 1; sheetId++) {
                    String macRound = read(macOutput, "xl/worksheets/sheet" + sheetId + ".xml");
                    check(macRound.contains("fitToPage=\"1\"") && macRound.contains("fitToWidth=\"1\"")
                            && macRound.contains("fitToHeight=\"1\""), "Mac round fits one page without Excel page setup calls");
                    check(macRound.contains("paperSize=\"9\"") && macRound.contains("orientation=\"portrait\""), "Mac portrait A4");
                }
                String macSchedule = read(macOutput, "xl/worksheets/sheet" + (count + 2) + ".xml");
                check(macSchedule.contains("orientation=\"landscape\"") && macSchedule.contains("fitToPage=\"1\""), "Mac schedule landscape");
                check(occurrences(macSchedule, "Schedule title") == 4, "Mac retains four schedule copies");
                check(read(macOutput, "xl/workbook.xml").contains(area), "Mac preserves round print area");
                check(read(macOutput, "xl/worksheets/sheet" + (count + 1) + ".xml").contains("<f>data!A3</f><v>1</v>"), "Mac preserves formula and cache");
                for (int copies = 1; copies <= 4; copies++) {
                    fixture(source, league, false, copies);
                    XlsxPrintLayout.prepare(source, output, league);
                    check(occurrences(read(output, "xl/worksheets/sheet" + (count + 2) + ".xml"), "Schedule title") == 4,
                            "Existing " + copies + " forms become four, not " + (copies * 4));
                }
                fixture(source, league, true);
                try {
                    XlsxPrintLayout.prepare(source, output, league);
                    throw new AssertionError("Missing required sheet must fail");
                } catch (IOException expected) {
                    check(expected.getMessage().contains("nenašel") || expected.getMessage().contains("Nenašel"), "Useful error");
                }
            }
            System.out.println("OK: 6/7/8 teams, ordering, hidden data, print areas, A4, local names, formulas, unchanged source, missing sheets.");
        } finally {
            try (var files = Files.walk(directory)) {
                for (Path path : files.sorted(Comparator.reverseOrder()).toList()) Files.deleteIfExists(path);
            }
        }
    }

    private static void fixture(Path path, LeagueKind league, boolean omitSchedule) throws IOException {
        fixture(path, league, omitSchedule, 3);
    }

    private static void fixture(Path path, LeagueKind league, boolean omitSchedule, int copies) throws IOException {
        int count = league == LeagueKind.SIX_TEAMS ? 5 : 7;
        List<String> names = new ArrayList<>(List.of("data"));
        for (int i = count; i >= 1; i--) names.add(league == LeagueKind.SIX_TEAMS ? i + " kolo" : i + ".");
        if (!omitSchedule) names.add(league == LeagueKind.SEVEN_TEAMS ? "los" : "Rozpis O'Brien");
        Map<String, String> entries = new LinkedHashMap<>();
        StringBuilder sheets = new StringBuilder();
        StringBuilder relationships = new StringBuilder();
        for (int i = 0; i < names.size(); i++) {
            int id = i + 1;
            sheets.append("<sheet name=\"").append(names.get(i)).append("\" sheetId=\"").append(id)
                    .append("\" r:id=\"rId").append(id).append("\"/>");
            relationships.append("<Relationship Id=\"rId").append(id).append("\" Target=\"worksheets/sheet")
                    .append(id).append(".xml\"/>");
            StringBuilder cells = new StringBuilder();
            if (i == 0) {
                for (int r = 3; r <= 9; r++) cells.append("<row r=\"").append(r).append("\"><c r=\"A").append(r)
                        .append("\" t=\"inlineStr\"><is><t>").append(r - 2).append(".</t></is></c></row>");
            } else if (i == count + 1) {
                for (int copy = 0; copy < copies; copy++) {
                    int row = 1 + copy * 4;
                    cells.append("<row r=\"").append(row).append("\" ht=\"25\" customHeight=\"1\"><c r=\"A").append(row)
                            .append("\" t=\"inlineStr\"><is><t>Schedule title</t></is></c></row>");
                    cells.append("<row r=\"").append(row + 1).append("\"><c r=\"C").append(row + 1)
                            .append("\" t=\"str\"><f>data!B3</f><v>Team A - Team B</v></c></row>");
                }
            } else cells.append("<row r=\"1\"><c r=\"A1\"><f>data!A3</f><v>1</v></c></row>");
            StringBuilder merges = new StringBuilder();
            if (i == count + 1) {
                merges.append("<mergeCells count=\"").append(copies).append("\">");
                for (int copy = 0; copy < copies; copy++) {
                    int row = 1 + copy * 4;
                    merges.append("<mergeCell ref=\"A").append(row).append(":C").append(row).append("\"/>");
                }
                merges.append("</mergeCells>");
            }
            entries.put("xl/worksheets/sheet" + id + ".xml", "<worksheet xmlns=\"" + NS
                    + "\"><dimension ref=\"A1:L30\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetData>"
                    + cells + "</sheetData>" + merges + "</worksheet>");
        }
        entries.put("xl/workbook.xml", "<workbook xmlns=\"" + NS + "\" xmlns:r=\"" + REL
                + "\"><bookViews><workbookView activeTab=\"1\"/></bookViews><sheets>" + sheets
                + "</sheets><definedNames><definedName name=\"LocalTest\" localSheetId=\"" + count
                + "\">$A$1</definedName></definedNames><calcPr/></workbook>");
        entries.put("xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + relationships + "</Relationships>");
        try (ZipOutputStream zip = new ZipOutputStream(Files.newOutputStream(path))) {
            for (var entry : entries.entrySet()) {
                zip.putNextEntry(new ZipEntry(entry.getKey()));
                zip.write(entry.getValue().getBytes(StandardCharsets.UTF_8));
                zip.closeEntry();
            }
        }
    }

    private static String read(Path path, String entry) throws IOException {
        try (ZipFile zip = new ZipFile(path.toFile()); var input = zip.getInputStream(zip.getEntry(entry))) {
            return new String(input.readAllBytes(), StandardCharsets.UTF_8);
        }
    }

    private static void check(boolean condition, String message) {
        if (!condition) throw new AssertionError(message);
    }

    private static int occurrences(String text, String part) { return text.split(java.util.regex.Pattern.quote(part), -1).length - 1; }
}
