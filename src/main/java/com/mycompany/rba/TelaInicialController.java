package com.mycompany.rba;

import java.io.IOException;
import javafx.fxml.FXML;

public class TelaInicialController {

    @FXML
    private void switchToSecondary() throws IOException {
        App.setRoot("cadastro");
    }
}
