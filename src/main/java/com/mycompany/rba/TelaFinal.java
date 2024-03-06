package com.mycompany.rba;

import java.io.IOException;
import javafx.fxml.FXML;

public class TelaFinal {
    
    @FXML
    private void voltarParaCentralDeAutomacoes() throws IOException {
        App.setRoot("central_automacoes");
    }
}