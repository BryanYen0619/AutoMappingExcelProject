package com.bryan.example;

import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.swing.JOptionPane;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;

import java.awt.Dialog;
import java.awt.Frame;
import java.io.File;

import jxl.Cell;
import jxl.CellType;
import jxl.CellView;
import jxl.DateCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class AutoMappingModel {
	final static long MAX_WORKING_MINUTE = 480; // 60 * 8
	final static String OUTPUT_LOG_EXCEL_SHEET_NAME = "Sheet";		// Excel 分頁名稱
	// Dialog text
	final static String DIALOG_TITLE_INFO = "提示";
	final static String DIALOG_HEADER_SUCCESS = "完成";
	final static String DIALOG_HEADER_ERROR = "失敗";
	
	// Label Tag
	final static String EXCEL_CELL_TITLE = "Title";
	final static String EXCEL_CELL_HOLIDAY = "Holiday";		// 假日
	final static String EXCEL_CELL_DEFAULT = "Default";		// 預設
	final static String EXCEL_CELL_NOT_ENOUGH_WORK_TIME = "NotEnoughWorkTime";		// 上班時數不足
	final static String EXCEL_CELL_IS_ABSENTEEISM = "IsAbsenteeism";		// 曠職
	final static String EXCEL_CELL_NO_OF_DUTY_RECORD = "NoOffDutyRecord";		// 沒有下班卡紀錄
	final static String EXCEL_CELL_ERROR_USER = "ErrorUser";		//	異常員工
	
	private String attendancePath;
	private String leavePath;
	private String logPath;
	private String fillAttendanceRecordPath;
	
//	private String attendanceGroupId;
//	private String attendanceGroupName;
//	private String attendanceId;
//	private String attendanceName;
//	private String attendanceDate;
//	private String startWorkingTime;		// YYY-MM-dd HH:mm:ss
//	private String endWorkingTime;		// YYY-MM-dd HH:mm:ss
//	private String lateTotalTime;
//	private String leaveEarlyTotalTime;
//	private String complexWorkingTime;
//	private String overtimeCategory;
//	private String allowanceCategory;
//	private String bookingRecord;
//	private String workingItem;
//	private String leaveId;
//	private String leaveDate;
	private String currentId = null;
	private String currentName = null;
	private String path = "";	// 目前檔案路徑
//	private String holidayPath;		// Excel 休假日期 路徑
	
//	private long workingMinute = 0;
	private int logSheetPosition = 1;
	private int oldCountPosition = 2;
	
//	private boolean isWeekDay = false;
//	private boolean isErrorWorkingTime = false;
//	private boolean isIncludeNoonTime = false;
	
	private Sheet fillAttendanceRecordSheet = null;
	
//	private List<WorkDataModel> workDataModelArrayList = new ArrayList<>();
	
	public AutoMappingModel() {
		
	}
	
	public void run() {
		List<WorkDataModel> workDataModelArrayList = new ArrayList<>();
		
		try {
			Workbook workbook = Workbook.getWorkbook(new File(attendancePath));
			Sheet attendanceSheet = workbook.getSheet(0);
			int attendanceSheetSize = attendanceSheet.getRows();	
//			System.out.println("Attendance Excel Size : " + attendanceSheetSize);		// 共幾筆
			
			Workbook leaveBook = Workbook.getWorkbook(new File(leavePath));
			Sheet leaveSheet = leaveBook.getSheet(0);
			int leaveSheetSize = leaveSheet.getRows();	
//			System.out.println("Leave Excel Size : " + leaveSheetSize);		// 共幾筆

			Workbook fillAttendanceRecordBook = null;
			if (fillAttendanceRecordPath != null) {
				fillAttendanceRecordBook = Workbook.getWorkbook(new File(fillAttendanceRecordPath));
				fillAttendanceRecordSheet = fillAttendanceRecordBook.getSheet(0);
//				System.out.println("Fill Attendance Record Excel Size : " + fillAttendanceRecordSheetSize);		// 共幾筆
			}
			
			System.out.println("logPath : " + logPath);
			
			WritableWorkbook logBook = Workbook.createWorkbook(new File(logPath));
			WritableSheet logSheet = logBook.createSheet(OUTPUT_LOG_EXCEL_SHEET_NAME, 0);

			for (int position = 1; position < attendanceSheetSize; position++) {		// 0是title
				WorkDataModel workDataModel = new WorkDataModel();

				loadAttendenceDataFromExcel(attendanceSheet, position, workDataModel);		// 讀取門禁資料
				List<LeaveDataModel> leaveDataModels = loadLeaveDataFromExcel(leaveSheet, leaveSheetSize, workDataModel);		// 讀取請假資料
	
				workDataModel.setLeaveData(leaveDataModels);

				workDataModelArrayList.add(workDataModel);
			}
			
			setLogExcelData(logSheet, workDataModelArrayList);

			// 補上最後一個計數
			checkSearchEnd(logSheet, workDataModelArrayList.get(workDataModelArrayList.size() - 1), oldCountPosition, logSheetPosition);
			
			logBook.write();
			logBook.close();

			leaveBook.close();
			
			if (fillAttendanceRecordBook != null) {
				fillAttendanceRecordBook.close();
			}
			workbook.close();
			
	        showDialog(DIALOG_TITLE_INFO, DIALOG_HEADER_SUCCESS, "已產出結果!");
		} catch (BiffException e) {
			e.printStackTrace();
	        showDialog(DIALOG_TITLE_INFO,DIALOG_HEADER_ERROR, e.getMessage());
		} catch (IOException e) {
			e.printStackTrace();
			showDialog(DIALOG_TITLE_INFO,DIALOG_HEADER_ERROR, e.getMessage());
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			showDialog(DIALOG_TITLE_INFO,DIALOG_HEADER_ERROR, e.getMessage());
		}
	}
	
	private String loadHolidayExcelData() {
		File directory = new File(".");//设定为当前文件夹
		//System.out.println(directory.getCanonicalPath());//获取标准的路径
		//System.out.println(directory.getAbsolutePath());//获取绝对路径
		try {
			path = directory.getCanonicalPath()+"/src/assets/";
			System.out.println("Project Path : " + path);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		String holidayPath = path+"107ygov_yfyshop_weekday.xls";
		
		return holidayPath;
	}
	
	/**
	 * 組織代碼
	 * 組織
	 * 員編
	 * 姓名
	 * 日期
	 * 第一次刷卡
	 * 最後一次刷卡
	 * 遲到總時長
	 * 早退總時長
	 * 請假狀況
	 * 綜合工時
	 * 加班類別/時數
	 * 津貼類別/時數
	 * 刷卡紀錄
	 * 班別名稱
	 */
	private static void loadAttendenceDataFromExcel(Sheet attendanceSheet, int position, WorkDataModel workDataModel) {
		workDataModel.setAttendanceGroupId(attendanceSheet.getCell(0, position ).getContents());
		workDataModel.setAttendanceGroupName(attendanceSheet.getCell(1, position ).getContents());
		workDataModel.setAttendanceId(attendanceSheet.getCell(2, position ).getContents());
		workDataModel.setAttendanceName(attendanceSheet.getCell(3, position ).getContents());
		workDataModel.setAttendanceDate(attendanceSheet.getCell(4,position).getContents());
		workDataModel.setStartWorkingTime(attendanceSheet.getCell(5,position).getContents());
		workDataModel.setEndWorkingTime(attendanceSheet.getCell(6, position).getContents());
		workDataModel.setLateTotalTime(attendanceSheet.getCell(7, position).getContents());	// 遲到
		workDataModel.setLeaveEarlyTotalTime(attendanceSheet.getCell(8, position).getContents());		// 早退
		workDataModel.setComplexWorkingTime(attendanceSheet.getCell(10, position).getContents());
		workDataModel.setOvertimeCategory(attendanceSheet.getCell(11, position).getContents());
		workDataModel.setAllowanceCategory(attendanceSheet.getCell(12, position).getContents());
		workDataModel.setBookingRecord(attendanceSheet.getCell(13, position).getContents());
		workDataModel.setWorkingItem(attendanceSheet.getCell(14, position).getContents());
	}
	
	private static List<LeaveDataModel> loadLeaveDataFromExcel(Sheet leaveSheet, int leaveSheetSize, WorkDataModel workDataModel) {
		float leaveSum = getWorkingMinute(workDataModel.getStartWorkingTime(), workDataModel.getEndWorkingTime());
		
		List<LeaveDataModel> leaveDataModels = new ArrayList();
		for (int search = 1; search < leaveSheetSize; search++) {		// 0是title
			String leaveId = leaveSheet.getCell(2, search ).getContents();
			String leaveDate = leaveSheet.getCell(5,search).getContents();

			if (workDataModel.getAttendanceId().equals(leaveId)) {
				if (workDataModel.getAttendanceDate().equals(leaveDate)) {
					leaveSum += Float.parseFloat(leaveSheet.getCell(10,search).getContents()) * 60;
					
					LeaveDataModel leaveDataModel = new LeaveDataModel();

					leaveDataModel.setCategory(leaveSheet.getCell(4,search).getContents());
					leaveDataModel.setCount(leaveSheet.getCell(10,search).getContents());			
					leaveDataModel.setStartTime(leaveSheet.getCell(7,search).getContents());
					leaveDataModel.setEndTime(leaveSheet.getCell(9, search).getContents());
					leaveDataModel.setLeaveSum(leaveSum);

					leaveDataModels.add(leaveDataModel);
				}
			}
		}
		
		return leaveDataModels;
	}
	
	/**
	 * 
	 * 工號
	 * 姓名
	 * 組織名稱
	 * 表單號
	 * 表單狀態
	 * 簽核
	 * 簽核人
	 * 表單種類
	 * 簽核順序
	 * 補刷卡日期
	 * 補刷卡時間
	 * 補刷卡原因
	 * 備註
	 * 不使用
	 * 不使用1
	 * 
	 * @param fillAttendanceRecordSheet
	 * @param position
	 */
	private void loadFillAttendenceRecordDataFromExcel(Sheet fillAttendanceRecordSheet, WritableSheet logSheet, int logPosition, boolean isNoEndWorkingTime) {
		int fillAttendanceRecordSheetSize = fillAttendanceRecordSheet.getRows();
		
		String fillAttendanceRecordId;
		String fillAttendanceRecordDate;
		String fillAttendanceRecordTime;
		for (int search = 1; search < fillAttendanceRecordSheetSize; search++) {
			fillAttendanceRecordId = fillAttendanceRecordSheet.getCell(0, search ).getContents();
			fillAttendanceRecordDate = fillAttendanceRecordSheet.getCell(9, search ).getContents();
			fillAttendanceRecordTime = fillAttendanceRecordSheet.getCell(10, search ).getContents();
		
		
		}
	}
	
	private static void checkSearchEnd(WritableSheet logSheet, WorkDataModel data, int oldPosition, int endPosition) {
//		System.out.println("end current name : " + data.getAttendanceName());
//		System.out.println("end current count begin : " + oldPosition);
//		System.out.println("end current count end : " + endPosition);

		Label labelEndName = new Label(3, endPosition, data.getAttendanceName() + " 計數", getEndInfoExcelCellSetting());
		String excel = "COUNTA(D"+String.valueOf(oldPosition) + ":" + "D"+String.valueOf(endPosition) +")";
		Formula labelEndCount = new Formula(20, endPosition, excel, getDateExcelCellSetting());
	
		try {
			logSheet.addCell(labelEndName);
			logSheet.addCell(labelEndCount); 
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
	private void checkSearchEnd(WritableSheet logSheet, WorkDataModel data, int oldPosition) {
		if (currentId == null) {
			currentId = data.getAttendanceId();
			currentName = data.getAttendanceName();
		}
		
		if (!currentId.equals(data.getAttendanceId())) {
//			System.out.println("current count begin : " + oldPosition);
//			System.out.println("current count end : " + logSheetPosition);

			Label labelEndName = new Label(3, logSheetPosition, currentName + " 計數", getEndInfoExcelCellSetting());
			String excel = "COUNTA(D"+String.valueOf(oldPosition) + ":" + "D"+String.valueOf(logSheetPosition) +")";
			Formula labelEndCount = new Formula(20, logSheetPosition, excel, getDateExcelCellSetting());
			
			try {
				logSheet.addCell(labelEndName);
				logSheet.addCell(labelEndCount); 
			} catch (RowsExceededException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WriteException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} 
			
			currentId = data.getAttendanceId();
			currentName = data.getAttendanceName();
			logSheetPosition++;
			oldCountPosition = logSheetPosition + 1;
		}
	}
	
	/**
	 * 0 組織代碼
	 * 1 組織
	 * 2 員編
	 * 3 姓名
	 * 4 日期
	 * 5 第一次刷卡
	 * 6 最後一次刷卡
	 * 7 總工時 (計算)
	 * 8 工時	(時)
	 * 9 工時	(分)
	 * 10 遲到總時長
	 * 11 早退總時長
	 * 12 請假狀況 (請假紀錄)
	 * 13 假別 (請假紀錄)
	 * 14 時數 (請假紀錄)
	 * 15 加總 (請假紀錄)
	 * 16 綜合工時
	 * 17 加班類別/時數
	 * 18 津貼類別/時數
	 * 19 刷卡紀錄
	 * 20 班別名稱
	 */
	private void setLogExcelData(WritableSheet logSheet, List<WorkDataModel> dataArrayList) {
		// 讀取假日資料Excel
		String holidayPath = loadHolidayExcelData();
		Workbook holidaydayBook = null;
		Sheet holidaySheet = null;
		int holidaySheetSize = 0;
		try {
			holidaydayBook = Workbook.getWorkbook(new File(holidayPath));
			holidaySheet = holidaydayBook.getSheet(0);
			holidaySheetSize = holidaySheet.getRows();	
		} catch (BiffException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	
		// 設定title
		setTitle(logSheet);
		
		for(int i = 0 ; i < dataArrayList.size(); i++) {
			// 取得請假資料
			List<LeaveDataModel> leaveDataArrayList = dataArrayList.get(i).getLeaveData();
			
			// 檢查是否目標使用者資料已結束
			// 暫時關閉，預計更改成以每個月工時為基準，扣除異常天數後的總和工作天
//			checkSearchEnd(logSheet, dataArrayList.get(i), oldCountPosition);
		
			List<String> labelAttribute = new ArrayList<>();
			
			// 檢查是否在假日
			boolean isHoliday = false;
			for (int w = 1; w < holidaySheetSize; w++) {
				String holiday = holidaySheet.getCell(0, w).getContents();
//				System.out.println("week day : " + weekday);
				if (dataArrayList.get(i).getAttendanceDate().equals(holiday)) {
					labelAttribute.add(EXCEL_CELL_HOLIDAY);
					dataArrayList.get(i).setLabelAttribute(labelAttribute);
					isHoliday = true;
					break;
				}
			}
			
			if (!isHoliday) {
				// 清除狀態
				labelAttribute.clear();
				
				// 檢查上班狀態 - 有上班 沒下班
				if (!dataArrayList.get(i).getStartWorkingTime().equals("") && dataArrayList.get(i).getEndWorkingTime().length() < 3) {
					labelAttribute.add(EXCEL_CELL_NO_OF_DUTY_RECORD);
				}
				
				// 計算上班時數
				float leaveSum = 0;
				if (leaveDataArrayList.size() > 0) {
					leaveSum = leaveDataArrayList.get(leaveDataArrayList.size() - 1).getLeaveSum();
				} else {
					if (!dataArrayList.get(i).getComplexWorkingTime().equals("")) {
						leaveSum = Float.valueOf(dataArrayList.get(i).getComplexWorkingTime()) * 60;
					}
				}
				
				// 檢查上班狀態 - 上班總時數不足
				if (leaveSum > 0 && leaveSum < MAX_WORKING_MINUTE) {
					labelAttribute.add(EXCEL_CELL_NOT_ENOUGH_WORK_TIME);
//					System.out.println("name: " +dataArrayList.get(i).getAttendanceName() + ", date: " + dataArrayList.get(i).getAttendanceDate());
//					System.out.println("leaveSum: " +leaveSum);
				}
				
				// 檢查上班狀態 - 曠職
				if (dataArrayList.get(i).getStartWorkingTime().length() < 3 && dataArrayList.get(i).getEndWorkingTime().length() < 3) {
					if (dataArrayList.get(i).getComplexWorkingTime().equals("")) {		// 曠職
						if (leaveSum == 0) {
							labelAttribute.add(EXCEL_CELL_IS_ABSENTEEISM);
	//							System.out.println("name: " +dataArrayList.get(i).getAttendanceName() + ", date: " + dataArrayList.get(i).getAttendanceDate());
						}
					}
				}
				
				dataArrayList.get(i).setLabelAttribute(labelAttribute);
			}
		
			// 儲存Log 主資料
			if (leaveDataArrayList.size() == 0) {
				setMainData(logSheet, dataArrayList.get(i), logSheetPosition);
				setLeaveData(logSheet, null, dataArrayList.get(i).getLabelAttribute(), logSheetPosition);
				logSheetPosition++;
			} else {
				// 儲存Log 請假資料
				for(int x = 0; x <leaveDataArrayList.size(); x++ ) {
					setMainData(logSheet, dataArrayList.get(i), logSheetPosition);
					setLeaveData(logSheet, leaveDataArrayList.get(x), dataArrayList.get(i).getLabelAttribute(), logSheetPosition);
					logSheetPosition++;
				}
			}
		}
			
		if (holidaydayBook != null) {
			holidaydayBook.close();
		}
	}
	
	private void setTitle(WritableSheet logSheet) {
		Label label101 = new Label(0, 0, "組織代碼",getLogExcelTitleCellSetting());
		Label label102 = new Label(1, 0, "組織",getLogExcelTitleCellSetting());
		Label label103 = new Label(2, 0, "員編",getLogExcelTitleCellSetting());
		Label label104 = new Label(3, 0, "姓名",getLogExcelTitleCellSetting());
		Label label105 = new Label(4, 0, "日期",getLogExcelTitleCellSetting());
		Label label106 = new Label(5, 0, "第一次刷卡",getLogExcelTitleCellSetting());
		Label label107 = new Label(6, 0, "最後一次刷卡",getLogExcelTitleCellSetting());
		Label label108 = new Label(7, 0, "總工時(分鐘)",getLogExcelTitleCellSetting());
		Label label1091 = new Label(8, 0, "工時(時)",getLogExcelTitleCellSetting());
		Label label1092 = new Label(9, 0, "工時(分)",getLogExcelTitleCellSetting());
		Label label110 = new Label(10, 0, "遲到總時長",getLogExcelTitleCellSetting());
		Label label111 = new Label(11, 0, "早退總時長",getLogExcelTitleCellSetting());
		Label label112 = new Label(12, 0, "請假狀況",getLogExcelTitleCellSetting());
		Label label113 = new Label(13, 0, "假別",getLogExcelTitleCellSetting());
		Label label114 = new Label(14, 0, "時數",getLogExcelTitleCellSetting());
		Label label115 = new Label(15, 0, "加總",getLogExcelTitleCellSetting());
		Label label116 = new Label(16, 0, "綜合工時",getLogExcelTitleCellSetting());
		Label label117 = new Label(17, 0, "加班類別/時數",getLogExcelTitleCellSetting());
		Label label118 = new Label(18, 0, "津貼類別/時數",getLogExcelTitleCellSetting());
		Label label119 = new Label(19, 0, "刷卡紀錄",getLogExcelTitleCellSetting());
		Label label120 = new Label(20, 0, "班別名稱",getLogExcelTitleCellSetting());
		
		try {
			logSheet.addCell(label101);
			logSheet.addCell(label102); 
			logSheet.addCell(label103); 
			logSheet.addCell(label104); 
			logSheet.addCell(label105); 
			logSheet.addCell(label106); 
			logSheet.addCell(label107); 
			logSheet.addCell(label108); 
			logSheet.addCell(label1091); 
			logSheet.addCell(label1092); 
			logSheet.addCell(label110); 
			logSheet.addCell(label111); 
			logSheet.addCell(label112); 
			logSheet.addCell(label113); 
			logSheet.addCell(label114); 
			logSheet.addCell(label115); 
			logSheet.addCell(label116); 
			logSheet.addCell(label117); 
			logSheet.addCell(label118); 
			logSheet.addCell(label119); 
			logSheet.addCell(label120);	
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
	private void setMainData(WritableSheet logSheet, WorkDataModel data, int position) {
		Label labelName = null;
		Label labelDate = null;
		Label labelWorkingMinute = null;
		Label labelWorkingHours = null;
		Label labelWorkingMin = null;
		Label labelComplexWorkingTime = null;
		Label labelStartWorkingTime = null;
		Label labelendWorkingTime= null;
		
		String startWorkingTime = data.getStartWorkingTime();
		String endWorkingTime = data.getEndWorkingTime();
		
		long workingMinute = getWorkingMinute(startWorkingTime, endWorkingTime);
		List<String> workTime = getHourTime(workingMinute * 60);
		
		List<String> labelAttribute = data.getLabelAttribute();
		if (labelAttribute != null) {
			for(int i = 0; i < labelAttribute.size(); i++ ) {
				if (labelAttribute.get(i).equals(EXCEL_CELL_HOLIDAY)) {		// 假日
					labelDate = new Label(4, position, data.getAttendanceDate(), getHolidayExcelCellSetting());
					labelWorkingMinute = new Label(7, position, String.valueOf(workingMinute), getLeaveExcelCellSetting(true));
					labelWorkingHours = new Label(8, position, workTime.get(0), getLeaveExcelCellSetting(true));
					labelWorkingMin = new Label(9, position, workTime.get(1), getLeaveExcelCellSetting(true));
				}
		
				if (labelAttribute.get(i).equals(EXCEL_CELL_NOT_ENOUGH_WORK_TIME)) {		// 上班時數不足
					labelName = new Label(3, position, data.getAttendanceName(), getWorkingTimeNoEnoughExcelCellSetting(false));
					labelComplexWorkingTime = new Label(16, position, data.getComplexWorkingTime(),getTotalExcelCellSetting());
				}	
		
				if (labelAttribute.get(i).equals(EXCEL_CELL_IS_ABSENTEEISM)) {		// 曠職
					labelName = new Label(3, position, data.getAttendanceName(), getWorkingTimeNoEnoughExcelCellSetting(false));
					labelStartWorkingTime = new Label(5, position, fromatDate(startWorkingTime, "HH:mm:ss"),getNoWorkExcelCellSetting());
					labelendWorkingTime = new Label(6, position, fromatDate(endWorkingTime, "HH:mm:ss"),getNoWorkExcelCellSetting());
					labelComplexWorkingTime = new Label(16, position, data.getComplexWorkingTime(),getNoWorkExcelCellSetting());
				}
				
				if (labelAttribute.get(i).equals(EXCEL_CELL_NO_OF_DUTY_RECORD)) {	// 沒有下班卡紀錄
					labelName = new Label(3, position, data.getAttendanceName(), getWorkingTimeNoEnoughExcelCellSetting(false));
					labelendWorkingTime = new Label(6, position, fromatDate(endWorkingTime, "HH:mm:ss"),getNoWorkRecordErrorExcelCellSetting());
				}
			}
		}
		
		Label labelGroupId = new Label(0, position, data.getAttendanceGroupId());
		Label labelGroupName = new Label(1, position, data.getAttendanceGroupName());
		Label labelId = new Label(2, position, data.getAttendanceId());
		if(labelName == null) {
			labelName = new Label(3, position, data.getAttendanceName());
		}
		if(labelDate == null) {
			labelDate = new Label(4, position, data.getAttendanceDate());
		}
		if(labelStartWorkingTime == null) {
			labelStartWorkingTime = new Label(5, position, fromatDate(startWorkingTime, "HH:mm:ss"), getDateExcelCellSetting());
		}
		if(labelendWorkingTime == null) {
			labelendWorkingTime = new Label(6, position, fromatDate(endWorkingTime, "HH:mm:ss"), getDateExcelCellSetting());
		}
		
		if(labelWorkingMinute == null) {
			labelWorkingMinute = new Label(7, position, String.valueOf(workingMinute), getLeaveExcelCellSetting(false));
		}
		if(labelWorkingHours == null) {
			labelWorkingHours = new Label(8, position, workTime.get(0), getLeaveExcelCellSetting(false));
		}
		if(labelWorkingMin == null) {
			labelWorkingMin = new Label(9, position, workTime.get(1), getLeaveExcelCellSetting(false));
		}
		
		Label labelLateTotalTime = new Label(10, position, data.getLateTotalTime());
		Label labelLeaveEarlyTotalTime = new Label(11, position, data.getLeaveEarlyTotalTime());
		
		if (labelComplexWorkingTime == null) {
			labelComplexWorkingTime = new Label(16, position, data.getComplexWorkingTime());
		}
		Label labelOvertimeCategory = new Label(17, position, data.getOvertimeCategory());

		Label labelAllowanceCategory = new Label(18, position, data.getAllowanceCategory());
		Label labelBookingRecord = new Label(19, position, data.getBookingRecord());
		Label labelWorkingItem = new Label(20, position, data.getWorkingItem());
		
		try {
			logSheet.addCell(labelGroupId);
			logSheet.addCell(labelGroupName); 
			logSheet.addCell(labelId); 
			logSheet.addCell(labelName); 
			logSheet.addCell(labelDate);
			logSheet.addCell(labelStartWorkingTime); 
			logSheet.addCell(labelendWorkingTime); 
			logSheet.addCell(labelWorkingMinute); 
			logSheet.addCell(labelWorkingHours); 
			logSheet.addCell(labelWorkingMin);
			logSheet.addCell(labelLateTotalTime); 
			logSheet.addCell(labelLeaveEarlyTotalTime); 
			logSheet.addCell(labelComplexWorkingTime); 
			logSheet.addCell(labelOvertimeCategory); 
			logSheet.addCell(labelAllowanceCategory); 
			logSheet.addCell(labelBookingRecord); 
			logSheet.addCell(labelWorkingItem); 
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
	private void setLeaveData(WritableSheet logSheet, LeaveDataModel data, List<String> excelCellTag, int position) {
		Label labelLeaveStatus = null;
		Label labelLeaveSum = null;
		Label labelLeaveCategory = null;
		Label labelLeaveCount = null;
		
		float leaveSum = 0;
		
		if (data != null) {
			String startTime = data.getStartTime();
			String endTime = data.getEndTime();
			leaveSum = data.getLeaveSum();

			if (startTime != null && endTime != null) {
				labelLeaveStatus = new Label(12, position, startTime + "-" + endTime);
			}
			labelLeaveCategory = new Label(13, position, data.getCategory());
			labelLeaveCount = new Label(14, position, data.getCount());
			
			if (leaveSum >= MAX_WORKING_MINUTE) {
				labelLeaveSum = new Label(15, position, String.valueOf(leaveSum / 60));
			}
		}
		
		if (excelCellTag != null) {
			for(int i = 0 ; i < excelCellTag.size(); i++) {
				if (excelCellTag.get(i).equals(EXCEL_CELL_NOT_ENOUGH_WORK_TIME)) {		// 上班時數不足
					if (leaveSum != 0) {
						labelLeaveSum = new Label(15, position, String.valueOf(leaveSum / 60), getTotalExcelCellSetting());
					} else {
						labelLeaveSum = new Label(15, position, "", getTotalExcelCellSetting());
					}
				}
				
				if (excelCellTag.get(i).equals(EXCEL_CELL_IS_ABSENTEEISM)) {		// 曠職
					labelLeaveSum = new Label(15, position, "", getNoWorkExcelCellSetting());
				}
			}
		}
		
		try {
			if (labelLeaveStatus != null) {
				logSheet.addCell(labelLeaveStatus);
			}
			if (labelLeaveCategory != null) {
				logSheet.addCell(labelLeaveCategory);
			}
			if (labelLeaveCount != null) {
				logSheet.addCell(labelLeaveCount);
			}
			if (labelLeaveSum != null) {
				logSheet.addCell(labelLeaveSum);
			}
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
	private static void setLogExcelDataFromLeaveData(WritableSheet logSheet, int logPosition, String startTime, String endTime, String leaveCategory, String leaveCount, float leaveSum, boolean isNullData, boolean isNoEndWorkingTime, boolean isWorkingTimeNotEnough, boolean isWeekDay) {
		Label labelLeaveStatus = null;
		Label labelLeaveCategory = null;
		Label labelLeaveCount = null;
		Label labelLeaveSum = null;
		
		if (!isWeekDay && isNoEndWorkingTime) {
			if (isNullData) {
				labelLeaveStatus = new Label(12, logPosition, "" ,getWorkingTimeNoEnoughExcelCellSetting(false));
				labelLeaveCategory = new Label(13, logPosition, "",getWorkingTimeNoEnoughExcelCellSetting(false));
				labelLeaveCount = new Label(14, logPosition, "",getWorkingTimeNoEnoughExcelCellSetting(false));
				labelLeaveSum = new Label(15, logPosition, "",getWorkingTimeNoEnoughExcelCellSetting(false));
			} else {
				if (startTime != null && endTime != null) {
					labelLeaveStatus = new Label(12, logPosition, startTime + "-" + endTime,getWorkingTimeNoEnoughExcelCellSetting(false));
				} else {
					labelLeaveStatus = new Label(12, logPosition, "" ,getWorkingTimeNoEnoughExcelCellSetting(false));
				}
				
				labelLeaveCategory = new Label(13, logPosition, leaveCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
				labelLeaveCount = new Label(14, logPosition, leaveCount,getWorkingTimeNoEnoughExcelCellSetting(false));
				if (leaveSum > 0) {
					labelLeaveSum = new Label(15, logPosition, String.valueOf(leaveSum / 60),getWorkingTimeNoEnoughExcelCellSetting(false));
				}
			}
		} else {
			if (startTime != null && endTime != null) {
				labelLeaveStatus = new Label(12, logPosition, startTime + "-" + endTime);	
			}
			
			labelLeaveCategory = new Label(13, logPosition, leaveCategory);
			labelLeaveCount = new Label(14, logPosition, leaveCount);
			if (leaveSum > 0) {
				if (isWorkingTimeNotEnough) {
					if (leaveSum < (MAX_WORKING_MINUTE - 60)) {		// 扣掉中午
						labelLeaveSum = new Label(15, logPosition, String.valueOf(leaveSum / 60),getTotalExcelCellSetting());
					} else {
						labelLeaveSum = new Label(15, logPosition, String.valueOf(leaveSum / 60));
					}
				} else {
					labelLeaveSum = new Label(15, logPosition, String.valueOf(leaveSum / 60));
				}
			}
		}
		
		try {
			if (labelLeaveStatus != null) {
				logSheet.addCell(labelLeaveStatus);
			}
			logSheet.addCell(labelLeaveCategory);
			logSheet.addCell(labelLeaveCount); 
			if (labelLeaveSum != null) {
				logSheet.addCell(labelLeaveSum);
			}
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
	private static long getWorkingMinute(String startDateString, String endDateString) {
//		取得 12 - 13的時間
//		判斷 end是否小於12:00
//		判斷 start是否大於13:00
		
		
		//要判斷有沒有跨中午
		if (!startDateString.equals("") && !endDateString.equals("")) {
			SimpleDateFormat format = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");  
			Date beginDate = null;
			Date endDate = null;
			try {
				beginDate = format.parse(startDateString);
				endDate= format.parse(endDateString); 
			} catch (ParseException e) {
				// TODO Auto-generated catch block
				//				e.printStackTrace();
				return 0;
			}  
			 
			
			long difference=endDate.getTime()-beginDate.getTime();
			long minute=difference/(60*1000);
			
			return minute;
		} else {
			return 0;
		}
	}
	
	private static long getLeaveMinute(String startTimeString, String endTimeString) {
		if (!startTimeString.equals("") && !endTimeString.equals("")) {
			SimpleDateFormat format = new java.text.SimpleDateFormat("HH:mm");  
			Date beginDate = null;
			Date endDate = null;
			try {
				beginDate = format.parse(startTimeString);
				endDate= format.parse(endTimeString); 
			} catch (ParseException e) {
				// TODO Auto-generated catch block
				//				e.printStackTrace();
				return 0;
			}  
			 
			
			long difference=endDate.getTime()-beginDate.getTime();
			long minute=difference/(60*1000);
			
			return minute;
		} else {
			return 0;
		}
	}
	
	private static String fromatDate(String date, String dateFormat) {
		//  準備輸出的格式
	    SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
	   
	    // 取得現在時間
	    Calendar calendar = Calendar.getInstance();
	   
		if (!date.equals("")) {
			SimpleDateFormat format = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");  
			Date endDate = null;
			try {
				endDate= format.parse(date); 
			} catch (ParseException e) {
				// TODO Auto-generated catch block
				//				e.printStackTrace();
				return null;
			} 
			
//			System.out.println("format date : " + sdf.format(calendar.getTime()));

			calendar.setTime(endDate);
			return sdf.format(calendar.getTime());
		} else {
			return null;
		}
	}
	
	private static List<String> getHourTime(long seconds) {
		List<String> output = new ArrayList<String>(); 
		
		int day = (int)TimeUnit.SECONDS.toDays(seconds); 
		long hours = TimeUnit.SECONDS.toHours(seconds) - (day *24);
		long minute = TimeUnit.SECONDS.toMinutes(seconds) - (TimeUnit.SECONDS.toHours(seconds)* 60);
		
		output.add(String.valueOf(hours) + "時");
		output.add(String.valueOf(minute) + "分");
		return output;
	}
	
	private static WritableCellFormat getLogExcelTitleCellSetting() {
		WritableFont myFont = new WritableFont(WritableFont.createFont("Arial"), 10);
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();
		try {
			myFont.setColour(Colour.WHITE);
			// 色碼參閱 http://www.cnblogs.com/smilsy/articles/2126377.html
			cellFormat.setBackground(Colour.BLUE);
			cellFormat.setFont(myFont); // 指定字型
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} // 背景顏色
		
		return cellFormat;
	}
	
	private static WritableCellFormat getHolidayExcelCellSetting() {
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			// 參閱 http://www.cnblogs.com/smilsy/articles/2126377.html
			cellFormat.setBackground(Colour.ROSE);
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} // 背景顏色
		
		return cellFormat;
	}
	
	private static WritableCellFormat getTotalExcelCellSetting() {
		WritableFont cellFont = new WritableFont(WritableFont.createFont("Arial"), 10);
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			// 參閱 http://www.cnblogs.com/smilsy/articles/2126377.html
			cellFont.setColour(Colour.RED);
			cellFormat.setFont(cellFont);
			
			cellFormat.setBackground(Colour.LAVENDER);
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} // 背景顏色
		
		return cellFormat;
	}
	
	private static WritableCellFormat getLeaveExcelCellSetting(boolean isWeekdy) {
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			// 參閱 http://www.cnblogs.com/smilsy/articles/2126377.html
			if (!isWeekdy) {
				cellFormat.setBackground(Colour.YELLOW);
			}
			cellFormat.setAlignment(Alignment.RIGHT); // 對齊方式
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN); // 加框線
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} // 背景顏色
		
		return cellFormat;
	}
	
	private static WritableCellFormat getDateExcelCellSetting() {
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			cellFormat.setAlignment(Alignment.RIGHT); // 對齊方式
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return cellFormat;
	}
	
	private static WritableCellFormat getEndInfoExcelCellSetting() {
		WritableFont cellFont = new WritableFont(WritableFont.createFont("Arial"), 10);
		
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			cellFont.setBoldStyle(WritableFont.BOLD);
			cellFormat.setFont(cellFont);
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return cellFormat;
	}
	
	private static WritableCellFormat getWorkingTimeNoEnoughExcelCellSetting(boolean isAlignmentRight) {
		return getWorkingTimeNoEnoughExcelCellSetting(isAlignmentRight, false);
	}
	
	private static WritableCellFormat getWorkingTimeNoEnoughExcelCellSetting(boolean isAlignmentRight, boolean isLeaveGroup) {
		WritableFont cellFont = new WritableFont(WritableFont.createFont("Arial"), 10);
		
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			if (isLeaveGroup) {
				cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN); // 加框線
			}
			cellFont.setColour(Colour.RED);
			cellFormat.setFont(cellFont);
			
			cellFormat.setBackground(Colour.CORAL);
			if (isAlignmentRight) {
				cellFormat.setAlignment(Alignment.RIGHT); // 對齊方式
			}
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return cellFormat;
	}
	
	private static WritableCellFormat getNoWorkRecordErrorExcelCellSetting() {
		WritableFont cellFont = new WritableFont(WritableFont.createFont("Arial"), 10);
		
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			cellFont.setColour(Colour.RED);
			cellFormat.setFont(cellFont);
			
			cellFormat.setBackground(Colour.PALE_BLUE);
			cellFormat.setAlignment(Alignment.RIGHT); // 對齊方式
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return cellFormat;
	}
	
	private static WritableCellFormat getNoWorkExcelCellSetting() {
		
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			cellFormat.setBackground(Colour.LIGHT_GREEN);
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return cellFormat;
	}
	
	public void setAttendancePath(String path) {
		this.attendancePath = path;
	}
	
	public void setLeavePath(String path) {
		this.leavePath = path;
	}
	
	public void setLogPath(String path) {
		this.logPath = path;
	}
	
	private static void showDialog(String title, String header, String message) {
		Alert alert = new Alert(AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(header);
        alert.setContentText(message);
        alert.showAndWait();
	}
}
