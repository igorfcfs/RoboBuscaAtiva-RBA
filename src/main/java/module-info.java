module com.mycompany.rba {
    requires javafx.controls;
    requires javafx.fxml;

    opens com.mycompany.rba to javafx.fxml;
    exports com.mycompany.rba;
}
