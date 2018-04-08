package com.bryan.example;

import java.awt.Desktop;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import javax.swing.JOptionPane;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Hyperlink;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.Slider;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.AnchorPane;
import javafx.stage.FileChooser;
import javafx.util.StringConverter;

public class EventHandlingController {
	
	final static String DEFAULT_OUTPUT_FILE_NAME = "AutoMappingExcel_Output.xls";
	
	@FXML
	private Button runButton;
	
	@FXML
	private TextField outputPathTextField;
	
	@FXML
	private Button outputPathSelectButton;
	
	@FXML
	private AnchorPane accessRecordLayout;
	
	@FXML
	private TextField accessPathTextField;
	
	@FXML
	private Button accessRecordSelectButton;
	
	@FXML
	private AnchorPane leaveRecordLayout;
	
	@FXML
	private TextField leavePathTextField;
	
	@FXML
	private Button leaveRecordSelectButton;
	
	@FXML
	private AnchorPane fillAccessRecordLayout;
	
	@FXML
	private TextField fillAccessRecordTextField;
	
	@FXML
	private Button fillAccessRecordSelectButton;
	
	/**
	 * Initializes the controller class. This method is automatically called
	 * after the fxml file has been loaded.
	 */
	@FXML
	private void initialize() {
		
		accessRecordLayout.setStyle("-fx-background-color: #cccccc;");
		leaveRecordLayout.setStyle("-fx-background-color: #cccccc;");
		fillAccessRecordLayout.setStyle("-fx-background-color: #cccccc;");
		
		final FileChooser fileChooser = new FileChooser();
		 
		initDefaultOutputFilePath();
		
		// Handle Button event.
		runButton.setOnAction((event) -> {
			System.out.println("button event");
			
			if (accessPathTextField.getText() == null) {
				
			}
			if (leavePathTextField.getText() == null) {
				
			}
			if (outputPathTextField.getText() == null) {
				
			}
			
			AutoMappingModel autoMapping = new AutoMappingModel();
			autoMapping.setAttendancePath(accessPathTextField.getText());
			autoMapping.setLeavePath(leavePathTextField.getText());
			autoMapping.setLogPath(outputPathTextField.getText());
			
			autoMapping.run();
				
		});
		
		outputPathSelectButton.setOnAction((event) ->{
			fileChooser.setTitle("儲存產出紀錄");
			fileChooser.setInitialFileName(DEFAULT_OUTPUT_FILE_NAME);
			File savedFile = fileChooser.showSaveDialog(null);

			if (savedFile != null) {

			    try {
			        saveFileRoutine(savedFile);
			    } catch(IOException e) {
			        e.printStackTrace();
			        System.out.println("An ERROR occurred while saving the file!");
			        return;
			    }

			    outputPathTextField.setText(savedFile.getAbsolutePath());
			}
		});
		
		accessRecordSelectButton.setOnAction((event) ->{
			File selectedFile = fileChooser.showOpenDialog(null);
			if (selectedFile != null) {
				accessPathTextField.setText(selectedFile.getAbsolutePath());
			}
		});
		
		leaveRecordSelectButton.setOnAction((event) ->{
			File selectedFile = fileChooser.showOpenDialog(null);
			if (selectedFile != null) {
				leavePathTextField.setText(selectedFile.getAbsolutePath());
			}
		});
		
		fillAccessRecordSelectButton.setOnAction((event) ->{
			File selectedFile = fileChooser.showOpenDialog(null);
			if (selectedFile != null) {
				fillAccessRecordTextField.setText(selectedFile.getAbsolutePath());
			}
		});
		
		initAccessRecordLayoutDragEvent();
		initLeaveRecordLayoutDragEvent();
		initFillAccessRecordLayoutDragEvent();
	}
	
	private void saveFileRoutine(File file)
			throws IOException{
		// Creates a new file and writes the txtArea contents into it
		file.createNewFile();
	}
	
	private void initDefaultOutputFilePath() {
		File directory = new File(".");
		String path="";
		try {
			directory.getCanonicalPath();
			path = directory.getCanonicalPath() + "/src/" + DEFAULT_OUTPUT_FILE_NAME;
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		outputPathTextField.setText(path);
	}
	
	private void initAccessRecordLayoutDragEvent() {
		//build drag 
		accessRecordLayout.setOnDragOver(new EventHandler<DragEvent>() { //node添加拖入文件事件  
	        public void handle(DragEvent event) {  
	            Dragboard dragboard = event.getDragboard();   
	            if (dragboard.hasFiles()) {  
	                File file = dragboard.getFiles().get(0);  
	                if (file.getAbsolutePath().endsWith(".xls")) { //用來過濾拖入文件副檔名  
	                    event.acceptTransferModes(TransferMode.COPY);//接受拖入文件  
	                }  
	            }  
	  
	        }  
	    });
		accessRecordLayout.setOnDragDropped(new EventHandler<DragEvent>() { //拖入後鬆開滑鼠觸發的事件  
	        public void handle(DragEvent event) {  
	            // get drag enter file  
	            Dragboard dragboard = event.getDragboard();  
	            if (event.isAccepted()) {  
	                File file = dragboard.getFiles().get(0); //取得拖入的文件  
	                accessPathTextField.setText(file.getAbsolutePath());
	            }  
	        }  
	    }); 
	}
	
	private void initLeaveRecordLayoutDragEvent() {
		//build drag 
		leaveRecordLayout.setOnDragOver(new EventHandler<DragEvent>() { //node添加拖入文件事件  
	        public void handle(DragEvent event) {  
	            Dragboard dragboard = event.getDragboard();   
	            if (dragboard.hasFiles()) {  
	                File file = dragboard.getFiles().get(0);  
	                if (file.getAbsolutePath().endsWith(".xls")) { //用來過濾拖入文件副檔名  
	                    event.acceptTransferModes(TransferMode.COPY);//接受拖入文件  
	                }  
	            }  
	  
	        }  
	    });
		leaveRecordLayout.setOnDragDropped(new EventHandler<DragEvent>() { //拖入後鬆開滑鼠觸發的事件  
	        public void handle(DragEvent event) {  
	            // get drag enter file  
	            Dragboard dragboard = event.getDragboard();  
	            if (event.isAccepted()) {  
	                File file = dragboard.getFiles().get(0); //取得拖入的文件  
	                leavePathTextField.setText(file.getAbsolutePath());
	            }  
	        }  
	    }); 
	}
	
	private void initFillAccessRecordLayoutDragEvent() {
		//build drag 
		fillAccessRecordLayout.setOnDragOver(new EventHandler<DragEvent>() { //node添加拖入文件事件  
	        public void handle(DragEvent event) {  
	            Dragboard dragboard = event.getDragboard();   
	            if (dragboard.hasFiles()) {  
	                File file = dragboard.getFiles().get(0);  
	                if (file.getAbsolutePath().endsWith(".xls")) { //用來過濾拖入文件副檔名  
	                    event.acceptTransferModes(TransferMode.COPY);//接受拖入文件  
	                }  
	            }  
	  
	        }  
	    });
		fillAccessRecordLayout.setOnDragDropped(new EventHandler<DragEvent>() { //拖入後鬆開滑鼠觸發的事件  
	        public void handle(DragEvent event) {  
	            // get drag enter file  
	            Dragboard dragboard = event.getDragboard();  
	            if (event.isAccepted()) {  
	                File file = dragboard.getFiles().get(0); //取得拖入的文件  
	                fillAccessRecordTextField.setText(file.getAbsolutePath());
	            }  
	        }  
	    }); 
	}
}
