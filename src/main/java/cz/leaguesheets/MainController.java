package cz.leaguesheets;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.RadioButton;
import javafx.scene.control.ToggleGroup;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Modality;

import java.awt.Desktop;
import java.io.File;
import java.nio.file.Path;

public final class MainController {
    @FXML private ComboBox<LeagueKind> leagueBox;
    @FXML private RadioButton pdfButton;
    @FXML private RadioButton printButton;
    @FXML private CheckBox openPdfBox;
    @FXML private Button runButton;
    @FXML private ProgressBar progressBar;
    @FXML private Label statusLabel;

    @FXML
    private void initialize() {
        leagueBox.getItems().setAll(LeagueKind.values());
        leagueBox.getSelectionModel().select(LeagueKind.SIX_TEAMS);

        ToggleGroup outputGroup = new ToggleGroup();
        pdfButton.setToggleGroup(outputGroup);
        printButton.setToggleGroup(outputGroup);
        pdfButton.setSelected(true);

        openPdfBox.disableProperty().bind(pdfButton.selectedProperty().not());
        progressBar.setVisible(false);
        statusLabel.setVisible(false);
    }

    @FXML
    private void chooseFileAndRun() {
        FileChooser chooser = new FileChooser();
        chooser.setTitle("Vyber Excel soubor");
        chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel soubory", "*.xls", "*.xlsx", "*.xlsm"));
        File selected = chooser.showOpenDialog(runButton.getScene().getWindow());
        if (selected == null) {
            return;
        }

        LeagueKind league = leagueBox.getSelectionModel().getSelectedItem();
        boolean makePdf = pdfButton.isSelected();
        boolean openPdf = openPdfBox.isSelected();
        Path workbook = selected.toPath();
        Path output = makePdf ? defaultPdfPath(workbook) : null;

        setBusy(true);
        Task<String> task = new Task<>() {
            @Override
            protected String call() throws Exception {
                ExcelRunner.run(workbook, league, makePdf, output, update -> {
                    updateProgress(update.percent(), 100);
                    updateMessage(update.message());
                });
                if (makePdf && openPdf && Desktop.isDesktopSupported()) {
                    Desktop.getDesktop().open(output.toFile());
                }
                return makePdf ? "PDF vytvořeno:\n" + output : "Odesláno na výchozí tiskárnu.";
            }
        };

        progressBar.progressProperty().bind(task.progressProperty());
        statusLabel.textProperty().bind(task.messageProperty());

        task.setOnSucceeded(event -> {
            setBusy(false);
            showInfo("Hotovo", task.getValue());
        });
        task.setOnFailed(event -> {
            setBusy(false);
            Throwable error = rootCause(task.getException());
            showError(error.getMessage() == null ? error.toString() : error.getMessage());
        });

        Thread worker = new Thread(task, "league-sheets-worker");
        worker.setDaemon(true);
        worker.start();
    }

    @FXML
    private void showHelp() {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle("O aplikaci");
        alert.setHeaderText("Příprava ligových zápisů");
        alert.setContentText("""
                Co aplikace dělá
                Připraví ligové zápisy z Excelu buď jako PDF, nebo je rovnou pošle na tiskárnu.

                Postup
                1. Vyber ligu pro 6, 7 nebo 8 týmů.
                2. Vyber PDF nebo Tisk.
                3. Klikni na Vybrat Excel a vyber soubor.

                6 týmů
                - používá listy 1 kolo až 5 kolo
                - zápisy bere z rozsahu A1:J28
                - rozpis přidá jednou

                7 týmů
                - listy kol se jmenují 1, 2, 3...
                - aplikace bere jen kola uvedená v data!A3:A9
                - fungují hodnoty jako 1., 2., 3.
                - zápisy bere z rozsahu A1:J28
                - rozpis bere z listu rozpis nebo los a přidá ho dvakrát

                8 týmů
                - listy kol se jmenují 1, 2, 3...
                - aplikace bere jen kola uvedená v data!A3:A9
                - fungují hodnoty jako 1., 2., 3.
                - zápisy bere z rozsahu A1:J35
                - rozpis přidá dvakrát

                Nastavení stránky
                - zápisy jsou A4 na výšku, úzké okraje, 95 %
                - zápisy jsou centrované na stránce
                - rozpis je A4 na šířku a přizpůsobený na jednu stránku

                Windows a macOS
                - na obou systémech je potřeba nainstalovaný Microsoft Excel
                - na Macu povol aplikaci ovládání Excelu v nastavení Automatizace
                - na Macu se zápis podle potřeby zmenší pod 95 %, aby se vešel na A4
                """);
        alert.showAndWait();
    }

    private void setBusy(boolean busy) {
        leagueBox.setDisable(busy);
        pdfButton.setDisable(busy);
        printButton.setDisable(busy);
        runButton.setDisable(busy);
        progressBar.setVisible(busy);
        statusLabel.setVisible(busy);
        if (!busy) {
            progressBar.progressProperty().unbind();
            statusLabel.textProperty().unbind();
            progressBar.setProgress(0);
            statusLabel.setText("");
        }
    }

    private static Path defaultPdfPath(Path workbook) {
        String fileName = workbook.getFileName().toString();
        int dot = fileName.lastIndexOf('.');
        String base = dot > 0 ? fileName.substring(0, dot) : fileName;
        return workbook.resolveSibling(base + ".pdf");
    }

    private static Throwable rootCause(Throwable error) {
        Throwable cause = error;
        while (cause.getCause() != null) {
            cause = cause.getCause();
        }
        return cause;
    }

    private static void showInfo(String title, String message) {
        Platform.runLater(() -> new Alert(Alert.AlertType.INFORMATION, message, ButtonType.OK).showAndWait());
    }

    private static void showError(String message) {
        Platform.runLater(() -> {
            Dialog<Void> dialog = new Dialog<>();
            dialog.setTitle("Chyba");
            dialog.initModality(Modality.APPLICATION_MODAL);

            Label icon = new Label("!");
            icon.getStyleClass().add("error-icon");

            Label title = new Label("Zpracování se nepovedlo");
            title.getStyleClass().add("dialog-title");

            Label body = new Label(message);
            body.setWrapText(true);
            body.getStyleClass().add("dialog-body");
            body.setMaxWidth(460);

            VBox text = new VBox(8, title, body);
            HBox content = new HBox(16, icon, text);
            content.getStyleClass().add("dialog-content");

            ButtonType ok = new ButtonType("OK", ButtonBar.ButtonData.OK_DONE);
            dialog.getDialogPane().getButtonTypes().add(ok);
            dialog.getDialogPane().setContent(content);
            dialog.getDialogPane().getStylesheets().add(MainController.class.getResource("styles.css").toExternalForm());
            dialog.showAndWait();
        });
    }
}
