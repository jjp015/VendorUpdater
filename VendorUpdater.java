/**
 * @author: Jeremy Park
 * about: This program will create an application that will take the
 * master excel file and the vendor excel file and output the updated
 * excel file
 */

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.layout.GridPane;
import javafx.scene.text.Font;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * The VendorUpdater class extends the Application library that will contain
 * methods to create and run the application
 */
public class VendorUpdater extends Application {
    FileInputStream masterStream;
    FileInputStream vendorStream;
    FileOutputStream updateMaster;
    FileOutputStream updateMasterDetail;

    String masterString = new String();
    String vendorString = new String();
    String saveString = new String();

    int masterSize; //string length of the path to the master file
    int vendorIndex;  //string length of the path to the vendor file
    int saveIndex;  //string length of the path to the save folder
    int masterRows;  //number of rows in master Excel file
    int masterColumns; //number of columns in master Excel file
    int vendorColumns;  //number of volumns in vendor Excel file

    int masterSkuPosition;  //column position of SKU in master file
    int masterMapPosition;  //column position of MAP in master file
    int masterDescriptionPosition;  //column position of Description
    //in master file
    int masterStockPosition;  //column position of Stock in master file
    int masterPricePosition;  //column position of Price in master file

    int vendorSkuPosition;  //column position of SKU in vendor file
    int vendorMapPosition;  //column position of MAP in vendor file

    boolean checkMaster;  //check if master file is selected
    boolean checkVendor;  //check if vendor file is selected
    boolean checkSave;  //check if save location is selected
    boolean masterCheck; //check if master file is valid
    boolean vendorCheck;  //check if vendor file is valid

    XSSFWorkbook masterBook = new XSSFWorkbook();
    XSSFWorkbook vendorBook = new XSSFWorkbook();
    XSSFWorkbook updateBook = new XSSFWorkbook();
    XSSFWorkbook updateBookDetail = new XSSFWorkbook();

    XSSFSheet masterSheet;
    XSSFSheet vendorSheet;
    XSSFSheet updateSheet;
    XSSFSheet updateSheetDetail;

    String[] masterSku;
    int[] masterMap;
    String[] masterDescription;
    int[] masterStock;
    int[] masterPrice;
    String[] mapDetail;

    String[] vendorSku;
    int[] vendorMap;

    String masterFileString; //master file directory to string
    String vendorFileString; //vendor file directory to string
    String saveFileString;  //save location directory to string

    final int PANE_HEIGHT = 300;
    final int PANE_WIDTH = 600;
    final int COLUMN_3 = 2;
    final int COLUMN_4 = 3;
    final int COLUMN_5 = 4;
    final int COLUMN_6 = 5;
    final int FONT_SIZE = 30;
    final int PANE_GAP = 10;
    final int INSET_PAD = 12;
    final int VENDOR_ROW = 3;
    final int SAVE_ROW = 6;
    final int VENDOR_FIELD_ROW = 3;
    final int UPDATE_ROW = 8;

    /**
     * Program control automatically jumps to this method after
     * executing the program. This is the method to start and
     * run the application
     * @param stage the stage of the application
     */
    @Override
    public void start(Stage stage) {
        GridPane gridPane = new GridPane();
        Scene scene = new Scene(gridPane, PANE_WIDTH, PANE_HEIGHT);

        Button openMaster = new Button("Open Master File");
        Button openVendor = new Button("Open Vendor File");
        Button saveButton = new Button("   Save Location   ");
        Button runUpdate = new Button("     Run     ");

        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select Excel File");

        //to only display and select xlsx files
        fileChooser.getExtensionFilters().add(0,
                new ExtensionFilter("Excel Files", "*.xlsx"));
        DirectoryChooser directoryChooser = new DirectoryChooser();

        TextField masterField = new TextField();
        TextField vendorField = new TextField();
        TextField saveField = new TextField();
        masterField.setEditable(false);
        vendorField.setEditable(false);
        saveField.setEditable(false);

        //select a file after clicking on Open Master File
        openMaster.setOnAction((a) -> {
            File masterFile = fileChooser.showOpenDialog(stage);
            if(masterFile != null) {
                try {
                    masterFileString = masterFile.toString();
                    masterStream = new FileInputStream(masterFileString);
                } catch(Exception e) {
                    e.printStackTrace();
                }
                masterField.setText(masterFile.toString());
                masterString = masterFile.toString();
            }
        });

        //select a file after clicking on Open Vendor File
        openVendor.setOnAction((a) -> {
            File vendorFile = fileChooser.showOpenDialog(stage);
            if(vendorFile != null) {
                try {
                    vendorFileString = vendorFile.toString();
                    vendorStream = new FileInputStream(vendorFileString);
                } catch(Exception e) {
                    e.printStackTrace();
                }
                vendorField.setText(vendorFile.toString());
                vendorString = vendorFile.toString();
            }
        });

        //select a file after clicking on Save Location
        saveButton.setOnAction((a) -> {
            File saveFile = directoryChooser.showDialog(stage);
            if(saveFile != null) {
                try {
                    saveFileString = saveFile.toString();
                    saveField.setText(saveFile.getAbsolutePath());
                } catch(Exception e) {
                    e.printStackTrace();
                }
                saveField.setText(saveFile.toString());
                saveString = saveField.toString();
            }
        });

        //run checks and create new files after clicking on Run
        runUpdate.setOnAction((a) -> {
            masterSize = masterString.length();
            vendorIndex = vendorString.length();
            saveIndex = saveString.length();
            Alert emptyAlert;
            Alert fileAlert;

            //check if the master file exists in the path
            if(masterSize == 0) {
                emptyAlert = new Alert(AlertType.ERROR);
                emptyAlert.setTitle("Master File Error");
                emptyAlert.setHeaderText("File Selection is Empty");
                emptyAlert.setContentText("Please select an excel file for" +
                        " the Master Vendor.");
                emptyAlert.showAndWait();
                checkMaster = false;
            }
            //check if the file is Excel
            else if(!(masterString.endsWith("xlsx"))) {
                fileAlert = new Alert(AlertType.ERROR);
                fileAlert.setTitle("Master File Error");
                fileAlert.setHeaderText("Incorrect File Type");
                fileAlert.setContentText("Please select an excel file for" +
                        " the Master Vendor.");
                fileAlert.showAndWait();
                checkMaster = false;
            }
            else {
                checkMaster = true;
            }

            //check if the vendor file exists in the path
            if(vendorIndex == 0) {
                emptyAlert = new Alert(AlertType.ERROR);
                emptyAlert.setTitle("New Vendor File Error");
                emptyAlert.setHeaderText("File Selection is Empty");
                emptyAlert.setContentText("Please select an excel file for" +
                        " the New Vendor.");
                emptyAlert.showAndWait();
                checkVendor = false;
            }
            //check if the file is Excel
            else if(!(vendorString.endsWith("xlsx"))) {
                fileAlert = new Alert(AlertType.ERROR);
                fileAlert.setTitle("Master File Error");
                fileAlert.setHeaderText("Incorrect File Type");
                fileAlert.setContentText("Please select an excel file for" +
                        " the Master Vendor.");
                fileAlert.showAndWait();
                checkVendor = false;
            }
            else {
                checkVendor = true;
            }

            //check if save location exists in the path
            if(saveIndex == 0) {
                emptyAlert = new Alert(AlertType.ERROR);
                emptyAlert.setTitle("Save File Location Error");
                emptyAlert.setHeaderText("Save File Directory is Empty");
                emptyAlert.setContentText("Please select a directory to" +
                        " save updated vendor sheet.");
                emptyAlert.showAndWait();
                checkSave = false;
            }
            else {
                checkSave = true;
            }

            //continue if master, vendor file and save location exists
            if(checkMaster && checkVendor && checkSave) {
                int skuCount1 = 0;
                int descriptionCount1 = 0;
                int mapCount1 = 0;
                int stockCount1 = 0;
                int priceCount1 = 0;

                int skuCount2 = 0;
                int mapCount2 = 0;

                //call the Master Excel sheet and determine the
                //column positions of the title and number of columns
                try {
                    masterStream = new FileInputStream(masterFileString);
                    masterBook = new XSSFWorkbook(masterStream);
                    masterSheet = masterBook.getSheetAt(0);

                    Row rowMaster = masterSheet.getRow(0);

                    Iterator<Cell> cellIterator = rowMaster.cellIterator();
                    while(cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String cellString = cell.toString();

                        if(cellString.equals("SKU")) {
                            skuCount1 += 1;
                            masterSkuPosition = masterColumns;
                        }
                        else if(cell.toString().equals("Description")) {
                            descriptionCount1 += 1;
                            masterDescriptionPosition = masterColumns;
                        }
                        else if(cell.toString().equals("MAP")) {
                            mapCount1 += 1;
                            masterMapPosition = masterColumns;
                        }
                        else if(cell.toString().equals("Stock")) {
                            stockCount1 += 1;
                            masterStockPosition = masterColumns;
                        }
                        else if(cell.toString().equals("ListPrice")) {
                            priceCount1 += 1;
                            masterPricePosition = masterColumns;
                        }
                        masterColumns += 1;
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }

                //check if necessary titles exist in the master Excel sheet
                if((skuCount1 == 1) && (descriptionCount1 == 1) &&
                        (mapCount1 == 1) && (stockCount1 == 1) && (priceCount1 == 1)) {
                    masterCheck = true;
                }
                else {
                    fileAlert = new Alert(AlertType.ERROR);
                    fileAlert.setTitle("Incorrect Master File Error");
                    fileAlert.setHeaderText("Incorrect Content Titles in File");
                    fileAlert.setContentText("Please select a correct master" +
                            " file to read.");
                    fileAlert.showAndWait();
                }

                //call the Vendor Excel sheet and determine the
                //column positions of the title and number of columns
                try {
                    vendorStream = new FileInputStream(vendorFileString);
                    vendorBook = new XSSFWorkbook(vendorStream);
                    Sheet vendorSheet = vendorBook.getSheetAt(0);

                    Row rowVendor = vendorSheet.getRow(0);

                    Iterator<Cell> cellIterator = rowVendor.cellIterator();
                    while(cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String cellString = cell.toString();

                        if(cellString.equals("SKU")) {
                            skuCount2 += 1;
                            vendorSkuPosition = vendorColumns;
                        }
                        else if(cell.toString().equals("MAP")) {
                            mapCount2 += 1;
                            vendorMapPosition = vendorColumns;
                        }
                        vendorColumns += 1;
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }

                //check if necessary titles exist in the master Excel sheet
                if((skuCount2 == 1) && (mapCount2 == 1)) {
                    vendorCheck = true;
                }
                else {
                    fileAlert = new Alert(AlertType.ERROR);
                    fileAlert.setTitle("Incorrect Vendor File Error");
                    fileAlert.setHeaderText("Incorrect Content Titles in File");
                    fileAlert.setContentText("Please select a correct vendor" +
                            " file to read.");
                    fileAlert.showAndWait();
                }
            }

            //continue if master and vendor files are valid
            if(masterCheck && vendorCheck){
                try {
                    masterStream = new FileInputStream(masterFileString);
                    masterBook = new XSSFWorkbook(masterStream);
                    masterSheet = masterBook.getSheetAt(0);
                    masterRows = masterSheet.getPhysicalNumberOfRows();

                    int vendorRows = 0;
                    vendorStream = new FileInputStream(vendorFileString);
                    vendorBook = new XSSFWorkbook(vendorStream);
                    vendorSheet = vendorBook.getSheetAt(0);
                    vendorRows = vendorSheet.getPhysicalNumberOfRows();

                    masterSku = new String[masterRows];
                    masterMap = new int[masterRows];
                    masterDescription = new String[masterRows];
                    masterStock = new int[masterRows];
                    masterPrice = new int[masterRows];
                    mapDetail = new String[masterRows];

                    vendorSku = new String[vendorRows];
                    vendorMap = new int[vendorRows];

                    //read the Master excel sheet and store them in the proper array
                    for(int i = 1; i < masterRows; i++) {
                        masterSku[i] = (masterSheet.getRow(i).getCell(
                                (short)masterSkuPosition).toString());
                        masterMap[i] = (int)(masterSheet.getRow(i).getCell(
                                (short)masterMapPosition).getNumericCellValue());
                        masterDescription[i] = (masterSheet.getRow(i).getCell(
                                (short)masterDescriptionPosition).toString());
                        masterStock[i] = (int)(masterSheet.getRow(i).getCell(
                                (short)masterStockPosition).getNumericCellValue());
                        masterPrice[i] = (int)(masterSheet.getRow(i).getCell(
                                (short)masterPricePosition).getNumericCellValue());
                    }

                    //read the Vendor excel sheet and store them in the proper array
                    for(int i = 1; i < vendorRows; i++) {
                        vendorSku[i] = (vendorSheet.getRow(i).getCell(
                                (short)vendorSkuPosition).toString());
                        vendorMap[i] = (int)(vendorSheet.getRow(i).getCell(
                                (short)vendorMapPosition).getNumericCellValue());
                    }

                    //determine the MAP change status and store them in an array
                    for(int i = 1; i < masterRows; i++) {
                        for(int j = 1; j < vendorRows; j++) {
                            if(masterSku[i].equals(vendorSku[j])) {
                                if(masterMap[i] > vendorMap[j]) {
                                    mapDetail[i] = "MAP decreased";
                                    masterMap[i] = (vendorMap[j]);
                                    break;
                                }
                                else if(masterMap[i] < vendorMap[j]) {
                                    mapDetail[i] = "MAP increased";
                                    masterMap[i] = (vendorMap[j]);
                                    break;
                                }
                                else {
                                    mapDetail[i] = "MAP unchanged";
                                    masterMap[i] = (vendorMap[j]);
                                    break;
                                }
                            }
                            else {
                                mapDetail[i] = "Not Found";
                            }
                        }
                    }

                    //create and store the updated sheet in a new Excel sheet and write
                    //it in a file
                    updateSheet = updateBook.createSheet();
                    for(int i = 0; i < masterRows; i++) {
                        Row updateRow = updateSheet.createRow(i);
                        Cell cellSku = updateRow.createCell(0);
                        Cell cellDescription = updateRow.createCell(1);
                        Cell cellMap = updateRow.createCell(COLUMN_3);
                        Cell cellStock = updateRow.createCell(COLUMN_4);
                        Cell cellPrice = updateRow.createCell(COLUMN_5);
                        cellSku.setCellValue(masterSku[i]);
                        cellDescription.setCellValue(masterDescription[i]);
                        cellMap.setCellValue(masterMap[i]);
                        cellStock.setCellValue(masterStock[i]);
                        cellPrice.setCellValue(masterPrice[i]);
                    }
                    Row updateRow = updateSheet.createRow(0);
                    Cell cellSku = updateRow.createCell(0);
                    Cell cellDescription = updateRow.createCell(1);
                    Cell cellMap = updateRow.createCell(COLUMN_3);
                    Cell cellStock = updateRow.createCell(COLUMN_4);
                    Cell cellPrice = updateRow.createCell(COLUMN_5);
                    cellSku.setCellValue("SKU");
                    cellDescription.setCellValue("Description");
                    cellMap.setCellValue("MAP");
                    cellStock.setCellValue("Stock");
                    cellPrice.setCellValue("ListPrice");

                    updateMaster = new FileOutputStream(saveFileString +
                            "/master_vendor_output.xlsx");
                    updateBook.write(updateMaster);
                    updateMaster.close();

                    //create and store the updated detail sheet in a new Excel sheet
                    //and write in in a file
                    updateSheetDetail = updateBookDetail.createSheet();
                    for(int i = 0; i < masterRows; i++) {
                        Row updateRowDetail = updateSheetDetail.createRow(i);
                        Cell cellSkuDetail = updateRowDetail.createCell(0);
                        Cell cellDescriptionDetail = updateRowDetail.createCell(1);
                        Cell cellMapDetail = updateRowDetail.createCell(COLUMN_3);
                        Cell cellStockDetail = updateRowDetail.createCell(COLUMN_4);
                        Cell cellPriceDetail = updateRowDetail.createCell(COLUMN_5);
                        Cell cellDetail = updateRowDetail.createCell(COLUMN_6);
                        cellSkuDetail.setCellValue(masterSku[i]);
                        cellDescriptionDetail.setCellValue(masterDescription[i]);
                        cellMapDetail.setCellValue(masterMap[i]);
                        cellStockDetail.setCellValue(masterStock[i]);
                        cellPriceDetail.setCellValue(masterPrice[i]);
                        cellDetail.setCellValue(mapDetail[i]);
                    }
                    Row updateRowDetail = updateSheetDetail.createRow(0);
                    Cell cellSkuDetail = updateRowDetail.createCell(0);
                    Cell cellDescriptionDetail = updateRowDetail.createCell(1);
                    Cell cellMapDetail = updateRowDetail.createCell(COLUMN_3);
                    Cell cellStockDetail = updateRowDetail.createCell(COLUMN_4);
                    Cell cellPriceDetail = updateRowDetail.createCell(COLUMN_5);
                    Cell cellDetail = updateRowDetail.createCell(COLUMN_6);
                    cellSkuDetail.setCellValue("SKU");
                    cellDescriptionDetail.setCellValue("Description");
                    cellMapDetail.setCellValue("MAP");
                    cellStockDetail.setCellValue("Stock");
                    cellPriceDetail.setCellValue("ListPrice");
                    cellDetail.setCellValue("Detail");

                    updateMasterDetail = new FileOutputStream(saveFileString +
                            "/master_vendor_detail_output.xlsx");
                    updateBookDetail.write(updateMasterDetail);
                    updateMasterDetail.close();

                } catch (Exception e) {
                    e.printStackTrace();
                }
                masterColumns = 0; //reset the counter columns of Master file
                vendorColumns = 0; //reset the counter columns of Vendor file
            }
        });



        stage.setTitle("Vendor Updater");
        Label heading = new Label("     Update Vendor                      ");
        heading.setFont(new Font("Arial", FONT_SIZE));
        gridPane.setHgap(PANE_GAP);
        gridPane.setVgap(PANE_GAP);
        gridPane.setPadding(new Insets(INSET_PAD, INSET_PAD, INSET_PAD, INSET_PAD));
        gridPane.add(heading, 1, 0);
        gridPane.add(openMaster, 0, 1);
        gridPane.add(openVendor, 0, VENDOR_ROW);
        gridPane.add(saveButton, 0, SAVE_ROW);
        gridPane.add(masterField, 1, 1);
        gridPane.add(vendorField, 1, VENDOR_ROW);
        gridPane.add(saveField, 1, SAVE_ROW);
        gridPane.add(runUpdate, 1, UPDATE_ROW);
        stage.setScene(scene);
        stage.show();

        return;
    }

    /**
     * The main method starts at program execution and launches the program
     * @param args array of strings
     */
    public static void main(String[] args) {
        launch(args);

        return;
    }
}
