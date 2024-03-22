package com.lowman.creatpivotsheets;

import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.stage.FileChooser;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.concurrent.TimeUnit;

public class MainController {
    @FXML
    public Button pivotButton;
    @FXML
    public Button exitButton;
    @FXML
    public TextArea textArea;
    @FXML
    public ImageView imageView;
    public Button excelButton;

    private Excel excel;
    private File excelFile;
    private boolean ifExcelLoaded = false;
    private Loop loop;


    @FXML
    public void initialize(){
        excel = new Excel();
    }

    @FXML
    protected void onPivotButtonClick() throws IOException {
        /*if(ifExcelLoaded){
            imageView.setVisible(true);

            pivotButton.setDisable(true);
            pivotButton.setVisible(false);

            int noOfSheets = excel.howManySheets();
            textArea.appendText("Workbook has " + noOfSheets + " sheets \n");

            for(int i = 0; i < noOfSheets; i++){
                boolean over = excel.areThereOver25Lines(i);
                boolean normal = excel.checkIfNormalSheet(i);
                if(over && normal){
                    textArea.appendText("Sheet number " + i + " has over 25 lines and we are creating a pivot sheet\n");
                    excel.createPivotSheet(i);
                }else{
                    textArea.appendText("Sheet number " + i + "\n" + "Over 25 lines? " + over + "\n" + "Normal sheet? " + normal + "\n");
                }
            }
        }else{
            Dialog<String> dialog = new Dialog<>();
            ButtonType buttonType = new ButtonType("OK", ButtonBar.ButtonData.OK_DONE);
            dialog.setContentText("You need to select an excel file to modify");
            dialog.getDialogPane().getButtonTypes().add(buttonType);
            dialog.showAndWait();
        }*/

        if(ifExcelLoaded){
            imageView.setVisible(true);

            pivotButton.setDisable(true);
            pivotButton.setVisible(false);

            int noOfSheets = excel.howManySheets();
            textArea.appendText("Workbook has " + noOfSheets + " sheets \n");

            loop = new Loop(excel);

            ChangeListener<String> listener = new ChangeListener<String>() {
                @Override
                public void changed(ObservableValue<? extends String> observableValue, String s, String t1) {
                    if(t1.equals("end")){
                        return;
                    }
                    textArea.appendText(t1);
                }
            };

            loop.valueProperty().addListener(listener);
            new Thread(loop).start();
        }else{
            Dialog<String> dialog = new Dialog<>();
            ButtonType buttonType = new ButtonType("OK", ButtonBar.ButtonData.OK_DONE);
            dialog.setContentText("You need to select an excel file to modify");
            dialog.getDialogPane().getButtonTypes().add(buttonType);
            dialog.showAndWait();
        }

    }

    @FXML
    protected void exit(){
        System.exit(0);
    }

    @FXML
    protected void loadExcel(){
        //open file chooser
        FileChooser fileChooser = new FileChooser();
        FileChooser.ExtensionFilter extensionFilter = new FileChooser.ExtensionFilter("Excel files", "*.xlsx", "*.xlsm");
        fileChooser.getExtensionFilters().add(extensionFilter);
        fileChooser.setSelectedExtensionFilter(extensionFilter);
        fileChooser.setInitialDirectory(new File("C:\\Users\\low85\\Desktop"));
        excelFile = fileChooser.showOpenDialog(exitButton.getScene().getWindow());

        if (excelFile != null) {
            excel.setExcelFile(excelFile);
            ifExcelLoaded = true;
        }
    }

}
