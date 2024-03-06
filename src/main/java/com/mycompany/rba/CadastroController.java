package com.mycompany.rba;

import java.io.IOException;
import javafx.fxml.FXML;

public class CadastroController {

    @FXML
    private void voltarParaTelaInicial() throws IOException {
        App.setRoot("tela_inicial");
    }
    
    @FXML
    private void cadastrar() throws IOException {
        App.setRoot("preencher_dados_escolares");
    }
    
    @FXML
    private void entrar() throws IOException {
        App.setRoot("entrar");
    }
}