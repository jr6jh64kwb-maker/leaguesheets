package cz.leaguesheets;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TreeSet;
import java.util.TreeMap;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import java.nio.file.Files;

/** Changes print metadata in an Excel-produced working copy, preserving cells and formulas. */
final class XlsxPrintLayout {
    private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static final String REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    record Selection(List<String> rounds, String scheduleName) { }
    private record Sheet(Element element, String name, String path, int originalIndex, Document xml) { }

    private XlsxPrintLayout() { }

    static Selection prepare(Path source, Path destination, LeagueKind league) throws IOException {
        try {
            Map<String, byte[]> entries = new LinkedHashMap<>();
            try (ZipFile zip = new ZipFile(source.toFile())) {
                var iterator = zip.entries();
                while (iterator.hasMoreElements()) {
                    ZipEntry entry = iterator.nextElement();
                    try (var input = zip.getInputStream(entry)) {
                        entries.put(entry.getName(), input.readAllBytes());
                    }
                }
            }
            Document workbook = parse(entries.get("xl/workbook.xml"));
            Document relationships = parse(entries.get("xl/_rels/workbook.xml.rels"));
            Map<String, String> paths = new HashMap<>();
            for (Element relationship : descendants(relationships, "Relationship")) {
                String target = relationship.getAttribute("Target");
                paths.put(relationship.getAttribute("Id"), target.startsWith("/")
                        ? target.substring(1) : Path.of("xl").resolve(target).normalize().toString().replace('\\', '/'));
            }
            List<Sheet> sheets = new ArrayList<>();
            for (Element element : descendants(workbook, "sheet")) {
                String path = paths.get(element.getAttributeNS(REL, "id"));
                sheets.add(new Sheet(element, element.getAttribute("name"), path, sheets.size(), parse(entries.get(path))));
            }
            List<String> strings = new ArrayList<>();
            if (entries.containsKey("xl/sharedStrings.xml")) {
                for (Element si : descendants(parse(entries.get("xl/sharedStrings.xml")), "si")) {
                    strings.add(textRuns(si));
                }
            }
            List<Sheet> rounds = new ArrayList<>();
            if (league == LeagueKind.SIX_TEAMS) {
                for (int number = 1; number <= 5; number++) {
                    String pattern = "^" + number + "\\s*\\.?\\s*kolo$|^" + number + "\\s*kolo\\b.*";
                    rounds.add(sheets.stream().filter(s -> normalize(s.name()).matches(pattern)).findFirst()
                            .orElseThrow(() -> new IOException("Pro 6 týmů jsem nenašel všechna kola 1 kolo až 5 kolo a rozpis.")));
                }
            } else {
                Sheet data = sheets.stream().filter(s -> normalize(s.name()).equals("data")).findFirst()
                        .orElseThrow(() -> new IOException("Nenašel jsem list data. Kola čtu z data!A3:A9."));
                TreeSet<Integer> numbers = new TreeSet<>();
                for (Element cell : descendants(data.xml(), "c")) {
                    if (cell.getAttribute("r").matches("A[3-9]")) {
                        Integer number = roundNumber(cellText(cell, strings));
                        if (number != null) numbers.add(number);
                    }
                }
                if (numbers.isEmpty()) throw new IOException("V data!A3:A9 nejsou žádná kola. Očekávám hodnoty jako 1., 2., 3.");
                for (Sheet sheet : sheets) {
                    Integer number = roundNumber(sheet.name());
                    if (number != null && numbers.contains(number)) rounds.add(sheet);
                }
                rounds.sort(Comparator.comparingInt(s -> roundNumber(s.name())));
            }
            Sheet schedule = sheets.stream().filter(s -> normalize(s.name()).contains("rozpis")).findFirst().orElse(null);
            if (schedule == null && league == LeagueKind.SEVEN_TEAMS) {
                schedule = sheets.stream().filter(s -> normalize(s.name()).equals("los")).findFirst().orElse(null);
            }
            if (schedule == null) throw new IOException(league == LeagueKind.SEVEN_TEAMS
                    ? "Nenašel jsem list rozpis ani los." : "Nenašel jsem list rozpis.");
            fourSchedules(schedule.xml(), strings);
            List<Sheet> selected = new ArrayList<>(rounds);
            selected.add(schedule);
            List<Sheet> ordered = new ArrayList<>(selected);
            sheets.stream().filter(s -> !selected.contains(s)).forEach(ordered::add);
            Element sheetList = (Element) sheets.getFirst().element().getParentNode();
            Map<Integer, Integer> newIndexes = new HashMap<>();
            for (int i = 0; i < ordered.size(); i++) {
                Sheet sheet = ordered.get(i);
                newIndexes.put(sheet.originalIndex(), i);
                sheetList.appendChild(sheet.element());
                sheet.element().setAttribute("state", selected.contains(sheet) ? "visible" : "hidden");
                for (Element view : descendants(sheet.xml(), "sheetView")) {
                    view.setAttribute("tabSelected", i == 0 ? "1" : "0");
                }
            }
            // Local defined names use sheet positions, not sheetId, so remap after ordering.
            for (Element defined : descendants(workbook, "definedName")) {
                if (defined.hasAttribute("localSheetId")) {
                    int index = Integer.parseInt(defined.getAttribute("localSheetId"));
                    defined.setAttribute("localSheetId", newIndexes.get(index).toString());
                }
            }
            for (Element view : descendants(workbook, "workbookView")) {
                view.setAttribute("activeTab", "0");
                view.setAttribute("firstSheet", "0");
            }
            Element definedNames = child(workbook.getDocumentElement(), "definedNames",
                    List.of("calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"));
            for (Sheet sheet : selected) {
                int index = ordered.indexOf(sheet);
                boolean isSchedule = sheet == schedule;
                String area = isSchedule ? usedRange(sheet.xml())
                        : (league == LeagueKind.EIGHT_TEAMS ? "$A$1:$J$35" : "$A$1:$J$28");
                for (Element defined : descendants(definedNames, "definedName")) {
                    if (defined.getAttribute("name").equals("_xlnm.Print_Area")
                            && defined.getAttribute("localSheetId").equals(Integer.toString(index))) {
                        definedNames.removeChild(defined);
                    }
                }
                Element printArea = workbook.createElementNS(NS, "definedName");
                printArea.setAttribute("name", "_xlnm.Print_Area");
                printArea.setAttribute("localSheetId", Integer.toString(index));
                printArea.setTextContent("'" + sheet.name().replace("'", "''") + "'!" + area);
                definedNames.appendChild(printArea);
                configure(sheet.xml(), isSchedule);
            }
            for (Sheet sheet : sheets) entries.put(sheet.path(), serialize(sheet.xml()));
            entries.put("xl/workbook.xml", serialize(workbook));
            try (ZipOutputStream zip = new ZipOutputStream(Files.newOutputStream(destination))) {
                for (var entry : entries.entrySet()) {
                    zip.putNextEntry(new ZipEntry(entry.getKey()));
                    zip.write(entry.getValue());
                    zip.closeEntry();
                }
            }
            return new Selection(rounds.stream().map(Sheet::name).toList(), schedule.name());
        } catch (IOException e) {
            throw e;
        } catch (Exception e) {
            throw new IOException("Nepodařilo se připravit tiskové nastavení: " + e.getMessage());
        }
    }

    private static void configure(Document document, boolean schedule) {
        Element root = document.getDocumentElement();
        Element properties = child(root, "sheetPr", List.of("dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData"));
        child(properties, "pageSetUpPr", List.of()).setAttribute("fitToPage", schedule ? "1" : "0");
        List<String> afterPrint = List.of("pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst");
        Element options = child(root, "printOptions", afterPrint);
        options.setAttribute("horizontalCentered", "1");
        options.setAttribute("verticalCentered", "1");
        options.setAttribute("headings", "0");
        options.setAttribute("gridLines", "0");
        Element margins = child(root, "pageMargins", afterPrint.subList(1, afterPrint.size()));
        for (String side : List.of("left", "right", "top", "bottom")) margins.setAttribute(side, Double.toString(0.64 / 2.54));
        margins.setAttribute("header", "0.1");
        margins.setAttribute("footer", "0.1");
        Element setup = child(root, "pageSetup", afterPrint.subList(2, afterPrint.size()));
        setup.setAttribute("paperSize", "9");
        setup.setAttribute("orientation", schedule ? "landscape" : "portrait");
        setup.setAttribute("scale", schedule ? "100" : "95");
        setup.setAttribute("fitToWidth", schedule ? "1" : "0");
        setup.setAttribute("fitToHeight", schedule ? "1" : "0");
        setup.setAttribute("usePrinterDefaults", "0");
        // Discard printer-specific settings that could override the portable A4 setup.
        setup.removeAttributeNS(REL, "id");
    }

    private static String usedRange(Document document) throws IOException {
        List<Element> dimensions = descendants(document, "dimension");
        if (dimensions.isEmpty()) throw new IOException("Rozpis nemá určený použitý rozsah.");
        return dimensions.getFirst().getAttribute("ref").replaceAll("([A-Z]+)([0-9]+)", "\\$$1\\$$2");
    }

    /** A schedule sheet may already contain repeated forms. Print four copies of its first form. */
    private static void fourSchedules(Document document, List<String> strings) throws IOException {
        TreeMap<Integer, Map<String, String>> populatedRows = new TreeMap<>();
        for (Element cell : descendants(document, "c")) {
            String text = normalize(cellText(cell, strings));
            if (!text.isEmpty()) {
                String ref = cell.getAttribute("r");
                populatedRows.computeIfAbsent(rowNumber(ref), ignored -> new TreeMap<>())
                        .put(ref.replaceAll("[0-9]", ""), text);
            }
        }
        if (populatedRows.isEmpty()) throw new IOException("List rozpis je prázdný.");
        int first = populatedRows.firstKey();
        int last = populatedRows.lastKey();
        Map<String, String> heading = populatedRows.get(first);
        for (var row : populatedRows.tailMap(first, false).entrySet()) {
            if (row.getValue().equals(heading)) {
                last = populatedRows.lowerKey(row.getKey());
                break;
            }
        }
        int height = last - first + 1;
        Element data = descendants(document, "sheetData").getFirst();
        List<Element> rows = descendants(data, "row").stream()
                .filter(row -> Integer.parseInt(row.getAttribute("r")) >= first)
                .filter(row -> Integer.parseInt(row.getAttribute("r")) < first + height).toList();
        List<Element> merges = descendants(document, "mergeCell").stream().filter(merge -> {
            String[] bounds = merge.getAttribute("ref").split(":");
            return rowNumber(bounds[0]) >= first && rowNumber(bounds[bounds.length - 1]) < first + height;
        }).toList();
        while (data.hasChildNodes()) data.removeChild(data.getFirstChild());
        List<Element> mergeContainers = descendants(document, "mergeCells");
        Element mergeContainer = mergeContainers.isEmpty() ? null : mergeContainers.getFirst();
        if (mergeContainer != null) while (mergeContainer.hasChildNodes()) mergeContainer.removeChild(mergeContainer.getFirstChild());
        int maxColumn = 1;
        for (int copy = 0; copy < 4; copy++) {
            int offset = 1 - first + copy * (height + 1);
            for (Element source : rows) {
                Element row = (Element) source.cloneNode(true);
                row.setAttribute("r", Integer.toString(Integer.parseInt(row.getAttribute("r")) + offset));
                for (Element cell : descendants(row, "c")) {
                    String ref = cell.getAttribute("r");
                    cell.setAttribute("r", shiftRows(ref, offset));
                    maxColumn = Math.max(maxColumn, columnNumber(ref));
                    // A print copy is a snapshot: moving relative formulas would change opponents/times.
                    for (Element formula : descendants(cell, "f")) cell.removeChild(formula);
                }
                data.appendChild(row);
            }
            if (mergeContainer != null) for (Element source : merges) {
                Element merge = (Element) source.cloneNode(true);
                String[] bounds = source.getAttribute("ref").split(":");
                merge.setAttribute("ref", shiftRows(bounds[0], offset) + ":" + shiftRows(bounds[1], offset));
                maxColumn = Math.max(maxColumn, columnNumber(bounds[1]));
                mergeContainer.appendChild(merge);
            }
        }
        if (mergeContainer != null) mergeContainer.setAttribute("count", Integer.toString(merges.size() * 4));
        for (String tag : List.of("rowBreaks", "colBreaks")) {
            for (Element breaks : descendants(document, tag)) breaks.getParentNode().removeChild(breaks);
        }
        descendants(document, "dimension").getFirst().setAttribute("ref", "A1:" + columnName(maxColumn) + (height * 4 + 3));
    }

    private static int rowNumber(String ref) { return Integer.parseInt(ref.replaceAll("[^0-9]", "")); }

    private static int columnNumber(String ref) {
        int column = 0;
        for (char c : ref.replaceAll("[^A-Z]", "").toCharArray()) column = column * 26 + c - 'A' + 1;
        return column;
    }

    private static String columnName(int column) {
        StringBuilder result = new StringBuilder();
        while (column > 0) { column--; result.insert(0, (char) ('A' + column % 26)); column /= 26; }
        return result.toString();
    }

    private static String shiftRows(String ref, int offset) {
        return ref.replaceAll("[0-9]", "") + (rowNumber(ref) + offset);
    }

    private static Integer roundNumber(String text) {
        String normalized = normalize(text);
        if (!normalized.matches("\\d+\\.?")) return null;
        try { return Integer.valueOf(normalized.replace(".", "")); }
        catch (NumberFormatException e) { return null; }
    }

    private static String normalize(String text) { return text.strip().toLowerCase(Locale.ROOT); }

    private static String cellText(Element cell, List<String> strings) {
        if (cell.getAttribute("t").equals("inlineStr")) return textRuns(cell);
        List<Element> values = descendants(cell, "v");
        if (values.isEmpty()) return "";
        String value = values.getFirst().getTextContent();
        return cell.getAttribute("t").equals("s") ? strings.get(Integer.parseInt(value)) : value;
    }

    private static String textRuns(Element element) {
        StringBuilder result = new StringBuilder();
        for (Element text : descendants(element, "t")) result.append(text.getTextContent());
        return result.toString();
    }

    private static List<Element> descendants(Node node, String name) {
        var nodes = node instanceof Document document ? document.getElementsByTagNameNS("*", name)
                : ((Element) node).getElementsByTagNameNS("*", name);
        List<Element> result = new ArrayList<>();
        for (int i = 0; i < nodes.getLength(); i++) result.add((Element) nodes.item(i));
        return result;
    }

    private static Element child(Element parent, String name, List<String> following) {
        Node insertion = null;
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (name.equals(node.getLocalName())) return (Element) node;
            if (insertion == null && following.contains(node.getLocalName())) insertion = node;
        }
        Element element = parent.getOwnerDocument().createElementNS(NS, name);
        parent.insertBefore(element, insertion);
        return element;
    }

    private static Document parse(byte[] bytes) throws Exception {
        if (bytes == null) throw new IOException("V sešitu chybí požadovaná XML část.");
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
        return factory.newDocumentBuilder().parse(new ByteArrayInputStream(bytes));
    }

    private static byte[] serialize(Document document) throws Exception {
        TransformerFactory factory = TransformerFactory.newInstance();
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
        factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_STYLESHEET, "");
        var transformer = factory.newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        transformer.transform(new DOMSource(document), new StreamResult(output));
        return output.toByteArray();
    }
}
