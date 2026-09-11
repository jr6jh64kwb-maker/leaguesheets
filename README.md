# LeagueSheets Java

JavaFX verze aplikace pro přípravu ligových zápisů. UI je uložené ve FXML, takže ho můžeš otevřít ve Scene Builderu.

## Otevření v IntelliJ

1. Otevři IntelliJ IDEA.
2. Zvol `File > Open`.
3. Vyber složku `work/LeagueSheetsJava`.
4. Spusť třídu `cz.leaguesheets.Main`.

## Otevření ve Scene Builderu

Otevři tento soubor:

`src/main/resources/cz/leaguesheets/main-view.fxml`

V IntelliJ můžeš na FXML kliknout pravým tlačítkem a zvolit `Open in Scene Builder`, pokud máš Scene Builder nastavený v IntelliJ.

## Důležité

Tisk a export PDF vyžadují nainstalovaný Microsoft Excel a Javu 21 nebo novější.

- Windows používá Windows PowerShell a COM automatizaci Excelu.
- macOS používá systémový AppleScript a Excel pro Mac. Při prvním spuštění povol ovládání Excelu aplikaci, ze které program spouštíš (např. IntelliJ IDEA). Oprávnění najdeš v Nastavení systému → Soukromí a zabezpečení → Automatizace.
- Na macOS se zpracovává dočasná kopie sešitu. Původní soubor se neukládá a ostatní otevřené sešity zůstávají otevřené. Není potřeba instalovat PowerShell.

Maven si při prvním otevření může stáhnout JavaFX knihovny. Pro spuštění na Macu použij IntelliJ nebo `mvn javafx:run`; knihovny v `target/fx-libs` jsou určené pro Windows.

## Co umí

- liga pro 6 týmů
- liga pro 7 týmů (kola podle data!A3:A9, zápisy A1:J28, list rozpis nebo los dvakrát)
- liga pro 8 týmů
- výstup do PDF
- přímý tisk
- zápisy A4 na výšku, 95 %, centrované na stránce
- rozpis A4 na šířku, přizpůsobený na jednu stránku

Na macOS se měřítko zápisu v případě potřeby sníží pod 95 %, aby se kvůli odlišným metrikám písem vešel na jednu A4. Výběr kol, tiskové oblasti a počty kopií rozpisu jsou na obou systémech stejné.

## Kontrola tiskové logiky

Regresní kontroly pro 6, 7 a 8 týmů jsou ve třídě `src/test/java/cz/leaguesheets/XlsxPrintLayoutTest.java`. Lze ji spustit přímo v IntelliJ jako Java program. Nevyžaduje Excel ani tiskárnu. Kontroluje pořadí kol, tiskové oblasti, A4, zachování vzorců a zdrojového souboru i chybějící listy.

Samotný export a přímý tisk používají nativní Excel a ověřují se na příslušném operačním systému.
