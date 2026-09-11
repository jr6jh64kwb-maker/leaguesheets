package cz.leaguesheets;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;

import java.io.IOException;

public final class LeagueSheetsApp extends Application {
    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader loader = new FXMLLoader(LeagueSheetsApp.class.getResource("main-view.fxml"));
        Scene scene = new Scene(loader.load());
        scene.getStylesheets().add(LeagueSheetsApp.class.getResource("styles.css").toExternalForm());
        stage.setTitle("Příprava ligových zápisů");
        stage.setResizable(false);
        stage.getIcons().add(new Image(LeagueSheetsApp.class.getResourceAsStream("bowling-icon.png")));
        stage.setScene(scene);
        stage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
