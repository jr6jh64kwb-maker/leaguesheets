package cz.leaguesheets;

import java.io.ByteArrayOutputStream;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.function.Consumer;

public final class ExcelRunner {
    private ExcelRunner() {
    }

    public static void run(Path workbook, LeagueKind league, boolean makePdf, Path output,
                           Consumer<ProgressUpdate> progress) throws IOException, InterruptedException {
        String os = System.getProperty("os.name", "");
        if (os.startsWith("Mac")) {
            MacExcelRunner.run(workbook, league, makePdf, output, progress);
            return;
        }
        if (!os.startsWith("Windows")) {
            throw new IOException("Aplikace vyžaduje Windows nebo macOS s nainstalovaným Microsoft Excelem.");
        }
        Path script = Files.createTempFile("league-sheets-", ".ps1");
        Files.write(script, utf8WithBom(SCRIPT));

        progress.accept(new ProgressUpdate(8, "Kontroluji Excel..."));
        ProcessBuilder builder = new ProcessBuilder(
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", script.toString(),
                "-WorkbookPath", workbook.toString(),
                "-League", league.scriptName(),
                "-Mode", makePdf ? "pdf" : "print"
        );
        if (makePdf) {
            builder.command().add("-OutputPath");
            builder.command().add(output.toString());
        }
        builder.redirectErrorStream(true);

        progress.accept(new ProgressUpdate(18, "Spouštím Excel..."));
        Process process = builder.start();
        String out = readPowerShellOutput(process, progress);
        int code = process.waitFor();
        try {
            Files.deleteIfExists(script);
        } catch (IOException ignored) {
        }
        if (code != 0) {
            throw new IOException(cleanPowerShellError(out));
        }
        progress.accept(new ProgressUpdate(100, "Hotovo"));
    }

    private static String readPowerShellOutput(Process process, Consumer<ProgressUpdate> progress) throws IOException {
        StringBuilder output = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream(), StandardCharsets.UTF_8))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.startsWith("LS_PROGRESS|")) {
                    ProgressUpdate update = parseProgressLine(line);
                    if (update != null) {
                        progress.accept(update);
                    }
                } else {
                    output.append(line).append(System.lineSeparator());
                }
            }
        }
        String text = output.toString();
        if (!text.contains("\uFFFD")) {
            return text;
        }
        return decodePowerShellOutput(text.getBytes(StandardCharsets.UTF_8));
    }

    private static ProgressUpdate parseProgressLine(String line) {
        String[] parts = line.split("\\|", 3);
        if (parts.length != 3) {
            return null;
        }
        try {
            return new ProgressUpdate(Integer.parseInt(parts[1]), parts[2]);
        } catch (NumberFormatException ignored) {
            return null;
        }
    }

    private static String decodePowerShellOutput(byte[] bytes) {
        String utf8 = new String(bytes, StandardCharsets.UTF_8);
        if (!utf8.contains("\uFFFD")) {
            return utf8;
        }
        return new String(bytes, Charset.forName("windows-1250"));
    }

    private static byte[] utf8WithBom(String text) throws IOException {
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        bytes.write(0xEF);
        bytes.write(0xBB);
        bytes.write(0xBF);
        bytes.write(text.getBytes(StandardCharsets.UTF_8));
        return bytes.toByteArray();
    }

    private static String cleanPowerShellError(String output) {
        if (output == null || output.isBlank()) {
            return "Zpracování selhalo.";
        }

        for (String line : output.split("\\R")) {
            String trimmed = line.trim();
            if (trimmed.isEmpty()
                    || trimmed.startsWith("At ")
                    || trimmed.startsWith("+")
                    || trimmed.startsWith("~")
                    || trimmed.startsWith("LS_PROGRESS|")
                    || trimmed.startsWith("CategoryInfo")
                    || trimmed.startsWith("FullyQualifiedErrorId")) {
                continue;
            }
            return trimmed;
        }

        return output.trim();
    }

    private static final String SCRIPT = """
            param(
              [Parameter(Mandatory=$true)][string]$WorkbookPath,
              [Parameter(Mandatory=$true)][ValidateSet('six','seven','eight')][string]$League,
              [Parameter(Mandatory=$true)][ValidateSet('pdf','print')][string]$Mode,
              [string]$OutputPath
            )

            $ErrorActionPreference = 'Stop'
            [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
            $OutputEncoding = [System.Text.UTF8Encoding]::new($false)
            $xlTypePdf = 0
            $xlQualityStandard = 0
            $xlPaperA4 = 9
            $xlPortrait = 1
            $xlLandscape = 2

            function Normalize([string]$text) {
              if ($null -eq $text) { return '' }
              return $text.Trim().ToLowerInvariant()
            }

            function Report-Progress([int]$percent, [string]$message) {
              [Console]::Out.WriteLine(('LS_PROGRESS|{0}|{1}' -f $percent, $message))
              [Console]::Out.Flush()
            }

            function Parse-RoundNumber([string]$text) {
              $n = Normalize $text
              if ($n -match '^(\\d+)\\.?$') { return [int]$Matches[1] }
              return $null
            }

            function Get-Sheets($workbook) {
              $items = @()
              for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
                $ws = $workbook.Worksheets.Item($i)
                $items += [pscustomobject]@{ Name = [string]$ws.Name; Worksheet = $ws }
              }
              return $items
            }

            function Find-Rozpis($sheets) {
              foreach ($s in $sheets) {
                $name = Normalize $s.Name
                if ($name -eq 'rozpis' -or $name.Contains('rozpis')) { return $s }
              }
              return $null
            }

            function Is-SixRound([string]$name, [int]$round) {
              $n = Normalize $name
              return $n -match ('^' + $round + '\\s*\\.?\\s*kolo$') -or $n -match ('^' + $round + '\\s*kolo\\b')
            }

            function Configure-Page($worksheet, [bool]$isRozpis, [string]$roundArea) {
              $used = $worksheet.UsedRange
              $ps = $worksheet.PageSetup
              if ($isRozpis) { $ps.PrintArea = $used.Address() } else { $ps.PrintArea = $roundArea }
              $ps.PaperSize = $xlPaperA4
              $ps.Orientation = $(if ($isRozpis) { $xlLandscape } else { $xlPortrait })
              if ($isRozpis) {
                $ps.Zoom = $false
                $ps.FitToPagesWide = 1
                $ps.FitToPagesTall = 1
              } else {
                $ps.Zoom = 95
                $ps.FitToPagesWide = $false
                $ps.FitToPagesTall = $false
              }
              $ps.CenterHorizontally = $true
              $ps.CenterVertically = $true
              $ps.LeftMargin = $worksheet.Application.CentimetersToPoints(0.64)
              $ps.RightMargin = $worksheet.Application.CentimetersToPoints(0.64)
              $ps.TopMargin = $worksheet.Application.CentimetersToPoints(0.64)
              $ps.BottomMargin = $worksheet.Application.CentimetersToPoints(0.64)
              $ps.HeaderMargin = $worksheet.Application.InchesToPoints(0.1)
              $ps.FooterMargin = $worksheet.Application.InchesToPoints(0.1)
              $ps.PrintGridlines = $false
              $ps.PrintHeadings = $false
            }

            function Resolve-SixTargets($sheets) {
              Report-Progress 35 'Hledám listy pro 6 týmů...'
              $targets = @()
              for ($round = 1; $round -le 5; $round++) {
                Report-Progress (35 + ($round * 4)) ('Kontroluji ' + $round + '. kolo...')
                $match = $sheets | Where-Object { Is-SixRound $_.Name $round } | Select-Object -First 1
                if ($null -ne $match) {
                  Report-Progress (40 + ($round * 4)) ('Nastavuji tisk: ' + $match.Name)
                  Configure-Page $match.Worksheet $false '$A$1:$J$28'
                  $targets += $match
                }
              }
              Report-Progress 62 'Hledám rozpis...'
              $rozpis = Find-Rozpis $sheets
              if ($null -ne $rozpis) {
                Report-Progress 66 'Nastavuji tisk: rozpis'
                $targets += New-FourSchedules $rozpis.Worksheet
              }
              if ($targets.Count -lt 6) { throw 'Pro 6 týmů jsem nenašel všechna kola 1 kolo až 5 kolo a rozpis.' }
              return $targets
            }

            function Get-RowSignature($worksheet, [int]$row, [int]$firstColumn, [int]$lastColumn) {
              $values = @()
              $populated = $false
              for ($column = $firstColumn; $column -le $lastColumn; $column++) {
                $value = Normalize ([string]$worksheet.Cells.Item($row, $column).Text)
                if ($value -ne '') { $populated = $true }
                $values += $value
              }
              if (-not $populated) { return '' }
              return [string]::Join([string][char]31, [string[]]$values)
            }

            function New-FourSchedules($worksheet) {
              Report-Progress 70 'Skládám čtyři rozpisy na jednu A4...'
              $used = $worksheet.UsedRange
              $firstColumn = $used.Column
              $lastColumn = $firstColumn + $used.Columns.Count - 1
              $firstRow = 0
              $lastRow = 0
              $heading = ''
              for ($row = $used.Row; $row -lt ($used.Row + $used.Rows.Count); $row++) {
                $signature = Get-RowSignature $worksheet $row $firstColumn $lastColumn
                if ($signature -eq '') { continue }
                if ($firstRow -eq 0) {
                  $firstRow = $row
                  $heading = $signature
                } elseif ($signature -eq $heading) {
                  break
                }
                $lastRow = $row
              }
              if ($firstRow -eq 0) { throw 'List rozpis je prázdný.' }
              $height = $lastRow - $firstRow + 1
              $width = $lastColumn - $firstColumn + 1
              $source = $worksheet.Range($worksheet.Cells.Item($firstRow, $firstColumn), $worksheet.Cells.Item($lastRow, $lastColumn))
              $values = $source.Value2
              $book = $worksheet.Parent
              $page = $book.Worksheets.Add([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
              $page.Name = 'rozpis 4x ' + (Get-Random -Minimum 100000 -Maximum 999999)
              for ($column = 1; $column -le $width; $column++) {
                $page.Columns.Item($column).ColumnWidth = $worksheet.Columns.Item($firstColumn + $column - 1).ColumnWidth
              }
              for ($copy = 0; $copy -lt 4; $copy++) {
                $start = 1 + $copy * ($height + 1)
                $destination = $page.Range($page.Cells.Item($start, 1), $page.Cells.Item($start + $height - 1, $width))
                [void]$source.Copy($destination)
                # Keep the same opponents and times instead of shifting relative formulas.
                $destination.Value2 = $values
                for ($row = 0; $row -lt $height; $row++) {
                  $page.Rows.Item($start + $row).RowHeight = $worksheet.Rows.Item($firstRow + $row).RowHeight
                }
                if ($copy -lt 3) { $page.Rows.Item($start + $height).RowHeight = 15 }
              }
              Configure-Page $page $true ''
              $page.PageSetup.PrintArea = $page.Range($page.Cells.Item(1, 1), $page.Cells.Item($height * 4 + 3, $width)).Address()
              return [pscustomobject]@{ Name = [string]$page.Name; Worksheet = $page }
            }

            function Resolve-DataTargets($sheets, [int]$teamCount, [string]$roundArea) {
              Report-Progress 35 'Čtu kola z listu data...'
              $data = $sheets | Where-Object { (Normalize $_.Name) -eq 'data' } | Select-Object -First 1
              if ($null -eq $data) { throw ('Nenašel jsem list data. Pro ' + $teamCount + ' týmů čtu kola z data!A3:A9.') }
              $rounds = @{}
              $cells = $data.Worksheet.Range('A3:A9')
              for ($i = 1; $i -le $cells.Cells.Count; $i++) {
                $num = Parse-RoundNumber ([string]$cells.Cells.Item($i).Text)
                if ($null -ne $num) { $rounds[$num] = $true }
              }
              if ($rounds.Count -eq 0) { throw 'V data!A3:A9 nejsou žádná kola. Očekávám hodnoty jako 1., 2., 3.' }
              $targets = @()
              foreach ($s in $sheets) {
                $num = Parse-RoundNumber $s.Name
                if ($null -ne $num -and $rounds.ContainsKey($num)) {
                  Report-Progress 52 ('Nastavuji tisk: kolo ' + $num)
                  Configure-Page $s.Worksheet $false $roundArea
                  $targets += $s
                }
              }
              $targets = $targets | Sort-Object { Parse-RoundNumber $_.Name }
              Report-Progress 64 'Hledám rozpis...'
              $rozpis = Find-Rozpis $sheets
              if ($teamCount -eq 7 -and $null -eq $rozpis) {
                $rozpis = $sheets | Where-Object { (Normalize $_.Name) -eq 'los' } | Select-Object -First 1
              }
              if ($null -eq $rozpis) {
                if ($teamCount -eq 7) { throw 'Nenašel jsem list rozpis ani los.' }
                throw 'Nenašel jsem list rozpis.'
              }
              Report-Progress 68 'Nastavuji tisk: rozpis'
              $targets = @($targets) + @(New-FourSchedules $rozpis.Worksheet)
              return $targets
            }

            $excel = $null
            $workbook = $null
            try {
              Report-Progress 20 'Spouštím Excel...'
              $excel = New-Object -ComObject Excel.Application
              $excel.Visible = $false
              $excel.DisplayAlerts = $false
              Report-Progress 28 'Otevírám Excel soubor...'
              $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
              Report-Progress 32 'Načítám listy...'
              $sheets = Get-Sheets $workbook
              $targets = switch ($League) {
                'six' { Resolve-SixTargets $sheets }
                'seven' { Resolve-DataTargets $sheets 7 '$A$1:$J$28' }
                'eight' { Resolve-DataTargets $sheets 8 '$A$1:$J$35' }
              }
              if ($Mode -eq 'pdf') {
                if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'Chybí cesta pro PDF.' }
                if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
                Report-Progress 82 'Vybírám listy pro PDF...'
                [string[]]$names = @($targets | ForEach-Object { $_.Name })
                $workbook.Worksheets.Item($names).Select()
                Report-Progress 90 'Ukládám PDF...'
                $excel.ActiveSheet.ExportAsFixedFormat($xlTypePdf, $OutputPath, $xlQualityStandard, $true, $false, [Type]::Missing, [Type]::Missing, $false, [Type]::Missing)
              } else {
                $count = @($targets).Count
                $index = 0
                foreach ($target in $targets) {
                  $index++
                  Report-Progress (80 + [int](15 * $index / $count)) ('Tisknu: ' + $target.Name)
                  $target.Worksheet.PrintOut([Type]::Missing, [Type]::Missing, 1, $false)
                }
              }
            } finally {
              if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
              if ($null -ne $excel) { try { $excel.Quit() } catch {} }
            }
            """;
}
