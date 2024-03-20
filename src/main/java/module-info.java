module com.example.creatpivotsheets {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;
    requires java.desktop;


    opens com.lowman.creatpivotsheets to javafx.fxml;
    exports com.lowman.creatpivotsheets;
}