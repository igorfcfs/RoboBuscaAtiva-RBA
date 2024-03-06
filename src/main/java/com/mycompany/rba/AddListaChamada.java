package com.mycompany.rba;

import java.io.IOException;
import javafx.fxml.FXML;

public class AddListaChamada {

    @FXML
    private void adicionar() throws IOException {
//        App.setRoot("cadastro");
    }
    
    @FXML
    private void proximo() throws IOException {
        App.setRoot("central_automacoes");
    }
}