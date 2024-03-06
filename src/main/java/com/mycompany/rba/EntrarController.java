package com.mycompany.rba;

import java.io.IOException;
import javafx.fxml.FXML;

public class EntrarController {
    
    @FXML
    private void cadastrar() throws IOException {
        App.setRoot("cadastro");
    }
    
    @FXML
    private void entrar() throws IOException {
        App.setRoot("preencher_dados_escolares");
    }
}