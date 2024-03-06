package com.mycompany.rba;

import java.io.IOException;
import javafx.fxml.FXML;

public class PreencherDadosEscolaresController {

    @FXML
    private void proximo() throws IOException {
        App.setRoot("add_lista_chamada");
    }
}