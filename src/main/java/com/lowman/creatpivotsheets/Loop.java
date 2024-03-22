package com.lowman.creatpivotsheets;

import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.TextArea;
import javafx.scene.image.ImageView;

public class Loop extends Task {

    private Excel excel;


    public Loop(Excel excel1){
        this.excel = excel1;
    }

    @Override
    protected String call() throws Exception {
        int noOfSheets = excel.howManySheets();

        for(int i = 0; i < noOfSheets; i++){
            boolean over = excel.areThereOver25Lines(i);
            boolean normal = excel.checkIfNormalSheet(i);
            if(over && normal){
                updateValue("Sheet number " + i + " has over 25 lines and we are creating a pivot sheet\n");
                excel.createPivotSheet(i);
            }else{
                updateValue("Sheet number " + i + "\n" + "Over 25 lines? " + over + "\n" + "Normal sheet? " + normal + "\n");
            }
        }
        return "end";
    }
}
