package com.mycompany.rba;

import java.io.IOException;
import javafx.fxml.FXML;

public class CentralAutomacoes {

    @FXML
    private void minhaEscola() throws IOException {
//        App.setRoot("cadastro");
    }
    
    @FXML
    private void executar() throws IOException {
        App.setRoot("gerar_arquivos");
    }
}