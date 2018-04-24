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
	final static long MAX_WORKING_MINUTE = 540; // 60 * 9
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
//	private String currentId = null;
//	private String currentName = null;
	private String path = "";	// 目前檔案路徑
//	private String holidayPath;		// Excel 休假日期 路徑
	
//	private long workingMinute = 0;
	private int logSheetPosition = 0;
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
			
//			Workbook weekdayBook = Workbook.getWorkbook(new File(weekdayPath));
//			Sheet weekdaySheet = weekdayBook.getSheet(0);
//			int weekdaySheetSize = weekdaySheet.getRows();	
//			System.out.println("Weekday Excel Size : " + weekdaySheetSize);		// 共幾筆
			
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
				
//				// 檢查是否在假日
//				String weekday;
//				isWeekDay = false;
//				for (int weekdayPosition = 1; weekdayPosition < weekdaySheetSize; weekdayPosition++) {
//					weekday = weekdaySheet.getCell(0, weekdayPosition).getContents();
////					System.out.println("week day : " + weekday);
//					if (attendanceDate.equals(weekday)) {
//						isWeekDay = true;
//						break;
//					}
//				}
				
				loadAttendenceDataFromExcel(attendanceSheet, position, workDataModel);		// 讀取門禁資料
				
//				if (startWorkingTime.length() < 3 && endWorkingTime.length() < 3  && complexWorkingTime.equals("")) {
//					isErrorWorkingTime = true;
//				} else {
//					if (!startWorkingTime.equals("") && endWorkingTime.length() < 3) {
//						isErrorWorkingTime = true;
//					} else {
//						isErrorWorkingTime = false;
//					}
//				}
				
//				if (currentId == null) {
//					currentId = attendanceId;
//					currentName = attendanceName;
//				}
//				
//				checkSearchEnd(logSheet, oldCountPosition);
				
//				setLogExcelData(logSheet, logPosition, isWeekDay, isErrorWorkingTime);
				
				List<LeaveDataModel> leaveDataModels = loadLeaveDataFromExcel(leaveSheet, leaveSheetSize, workDataModel);		// 讀取請假資料
				
//				for (int search = 1; search < leaveSheetSize; search++) {
//					String leaveId = leaveSheet.getCell(2, search ).getContents();
//					String leaveDate = leaveSheet.getCell(5,search).getContents();
//					
////						System.out.println("search "+search+", id : " + leaveId);
////						System.out.println("search "+search+", date : " + leaveDate);
//					
//					if (workDataModel.getAttendanceId().equals(leaveId)) {
//						if (workDataModel.getAttendanceDate().equals(leaveDate)) {
//							LeaveDataModel leaveDataModel = new LeaveDataModel();
//							
////							leaveDataModel.setPosition(logPosition);
//							
//							
//							
////							leaveDataModel.isNoEndWorkingTime(isErrorWorkingTime);
//							
//							String leaveCategory = leaveSheet.getCell(4,search).getContents();
//							String leaveCount = leaveSheet.getCell(10,search).getContents();
//						
//							leaveDataModel.setCategory(leaveCategory);
//							leaveDataModel.setCount(leaveCount);
//							
////								System.out.println("leaveCategory : " + leaveCategory);
////								System.out.println("leaveCount : " + leaveCount);
//							
//							 String startTime = leaveSheet.getCell(7,search).getContents();
//							 String endTime = leaveSheet.getCell(9, search).getContents();
//							 
//							// 用不到 暫時關閉 Begin.
//							// long leaveMinute = getLeaveMinute(startTime, endTime);
//							// End.
//							 
//							 leaveDataModel.setStartTime(startTime);
//							 leaveDataModel.setEndTime(endTime);
//							
////								System.out.println("search "+search+", start time : " + startTime);
////								System.out.println("search "+search+", end time : " + endTime);
////								System.out.println("search "+search+", leave minute : " + leaveMinute);
//						
//							leaveDataModels.add(leaveDataModel);
//							
////							setLogExcelData(logSheet, logPosition, isWeekDay, isErrorWorkingTime);
////							logPosition++;
//						}
//					}
//				}
				
				workDataModel.setLeaveData(leaveDataModels);

//				if (leaveDataModels.size() > 0) {
//					float leaveSum = workingMinute;
//					for(int i = 0; i < leaveDataModels.size(); i++) {
//						// 鎖定最後一筆才顯示加總不足提示
//						boolean isEndPosition = false;
//						if (i == leaveDataModels.size() - 1) {
//							isEndPosition = true;	
////							System.out.println("leave Sum : " + leaveSum);
//						}
//						leaveSum += Float.parseFloat(leaveDataModels.get(i).getCount()) * 60;
//						
//						if (leaveDataModels.get(i).isNoEndWorkingTime()) {
//							if (leaveSum == 0) {
//								leaveDataModels.get(i).isNoEndWorkingTime(true);
//							} else {
//								leaveDataModels.get(i).isNoEndWorkingTime(false);
//							}
//						}
//						
//						setLogExcelDataFromLeaveData(logSheet, leaveDataModels.get(i).getPosition(), leaveDataModels.get(i).getStartTime(),
//								leaveDataModels.get(i).getEndTime(),leaveDataModels.get(i).getCategory(),leaveDataModels.get(i).getCount(), 
//								leaveSum, false, leaveDataModels.get(i).isNoEndWorkingTime(), isEndPosition, isWeekDay);
//					}
//				} else {
//					setLogExcelDataFromLeaveData(logSheet, logPosition, null, null, null, null, 0, true, isErrorWorkingTime, true, isWeekDay);
//					logPosition++;
//				}
				
				workDataModelArrayList.add(workDataModel);
			}
			
			setLogExcelData(logSheet, workDataModelArrayList);

			// 補上最後一個計數
//			checkSearchEnd(logSheet, currentName, oldCountPosition, logPosition);
			
			logBook.write();
			logBook.close();
			
//			weekdayBook.close();
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
		System.out.println("end current name : " + data.getAttendanceName());
		System.out.println("end current count begin : " + oldPosition);
		System.out.println("end current count end : " + endPosition);

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
		String currentId = null;
		String currentName = null;
		if (currentId == null) {
			currentId = data.getAttendanceId();
			currentName = data.getAttendanceName();
		}
		
		if (!currentId.equals(data.getAttendanceId())) {
			System.out.println("current count begin : " + oldPosition);
			System.out.println("current count end : " + logSheetPosition);

			Label labelEndName = new Label(3, logSheetPosition, currentName + " 計數", getEndInfoExcelCellSetting());
			String excel = "COUNTA(D"+String.valueOf(oldPosition) + ":" + "D"+String.valueOf(logSheetPosition) +")";
			Formula labelEndCount = new Formula(20, logSheetPosition, excel, getDateExcelCellSetting());
			
			logSheetPosition++;
			
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
//	private WritableSheet initLogExcel(WritableWorkbook logBook, String logSheetName) {
//		WritableSheet logSheet = null;
//		logSheet = logBook.createSheet(logSheetName, 0);
//		
//		// 初始Title設定
//			Label label101 = new Label(0, 0, "組織代碼",getLogExcelTitleCellSetting());
//			Label label102 = new Label(1, 0, "組織",getLogExcelTitleCellSetting());
//			Label label103 = new Label(2, 0, "員編",getLogExcelTitleCellSetting());
//			Label label104 = new Label(3, 0, "姓名",getLogExcelTitleCellSetting());
//			Label label105 = new Label(4, 0, "日期",getLogExcelTitleCellSetting());
//			Label label106 = new Label(5, 0, "第一次刷卡",getLogExcelTitleCellSetting());
//			Label label107 = new Label(6, 0, "最後一次刷卡",getLogExcelTitleCellSetting());
//			Label label108 = new Label(7, 0, "總工時(分鐘)",getLogExcelTitleCellSetting());
//			Label label1091 = new Label(8, 0, "工時(時)",getLogExcelTitleCellSetting());
//			Label label1092 = new Label(9, 0, "工時(分)",getLogExcelTitleCellSetting());
//			Label label110 = new Label(10, 0, "遲到總時長",getLogExcelTitleCellSetting());
//			Label label111 = new Label(11, 0, "早退總時長",getLogExcelTitleCellSetting());
//			Label label112 = new Label(12, 0, "請假狀況",getLogExcelTitleCellSetting());
//			Label label113 = new Label(13, 0, "假別",getLogExcelTitleCellSetting());
//			Label label114 = new Label(14, 0, "時數",getLogExcelTitleCellSetting());
//			Label label115 = new Label(15, 0, "加總",getLogExcelTitleCellSetting());
//			Label label116 = new Label(16, 0, "綜合工時",getLogExcelTitleCellSetting());
//			Label label117 = new Label(17, 0, "加班類別/時數",getLogExcelTitleCellSetting());
//			Label label118 = new Label(18, 0, "津貼類別/時數",getLogExcelTitleCellSetting());
//			Label label119 = new Label(19, 0, "刷卡紀錄",getLogExcelTitleCellSetting());
//			Label label120 = new Label(20, 0, "班別名稱",getLogExcelTitleCellSetting());
//			
//			logSheet.addCell(label101);
//			logSheet.addCell(label102); 
//			logSheet.addCell(label103); 
//			logSheet.addCell(label104); 
//			logSheet.addCell(label105); 
//			logSheet.addCell(label106); 
//			logSheet.addCell(label107); 
//			logSheet.addCell(label108); 
//			logSheet.addCell(label1091); 
//			logSheet.addCell(label1092); 
//			logSheet.addCell(label110); 
//			logSheet.addCell(label111); 
//			logSheet.addCell(label112); 
//			logSheet.addCell(label113); 
//			logSheet.addCell(label114); 
//			logSheet.addCell(label115); 
//			logSheet.addCell(label116); 
//			logSheet.addCell(label117); 
//			logSheet.addCell(label118); 
//			logSheet.addCell(label119); 
//			logSheet.addCell(label120);	
//		logPosition++;
//		
//		return logSheet;
//	}
	
//	private WorkDataModel getLogSheetTitle() {
//		WorkDataModel workDataModel = new WorkDataModel();
//		
//		workDataModel.setAttendanceGroupId("組織代碼");
//		workDataModel.setAttendanceGroupName("組織");
//		workDataModel.setAttendanceId("員編");
//		workDataModel.setAttendanceName("姓名");
//		workDataModel.setAttendanceDate("日期");
//		workDataModel.setStartWorkingTime("第一次刷卡");
//		workDataModel.setEndWorkingTime("最後一次刷卡");
//		//
//		workDataModel.setWorkingMinute("總工時(分鐘)");		
//		workDataModel.setWorkingHours("工時(時)");		
//		workDataModel.setWorkingMin("工時(分)");		
//		//
//		workDataModel.setLateTotalTime("遲到總時長");	
//		workDataModel.setLeaveEarlyTotalTime("早退總時長");
//		//
//		List<LeaveDataModel> leaveDataModelArrayList = new ArrayList<>();
//		LeaveDataModel leaveDataModel = new LeaveDataModel();
//		leaveDataModel.setCategory("請假狀況");
//		leaveDataModel.setCount("假別");
//		leaveDataModel.setStartTime("時數");
//		leaveDataModel.setEndTime("加總");
//		
//		workDataModel.setLeaveData(leaveDataModelArrayList);	
//		//
//		workDataModel.setComplexWorkingTime("綜合工時");
//		workDataModel.setOvertimeCategory("加班類別/時數");
//		workDataModel.setAllowanceCategory("津貼類別/時數");
//		workDataModel.setBookingRecord("刷卡紀錄");
//		workDataModel.setWorkingItem("班別名稱");
//
//		workDataModel.setLabelAttribute(EXCEL_CELL_TITLE);
//		
//		return workDataModel;
//	}
	
	private void setLogExcelData(WritableSheet logSheet, List<WorkDataModel> dataArrayList) {
//		Label labelGroupId;
//		Label labelGroupName;
//		Label labelId;
//		Label labelName;
//		Label labelDate;
//		Label labelStartWorkingTime;
//		Label labelendWorkingTime;
//		Label labelWorkingMinute;
//		Label labelWorkingHours;
//		Label labelWorkingMin;
//		Label labelLateTotalTime;
//		Label labelLeaveEarlyTotalTime;
//		Label labelLeaveStatus;
//		Label labelLeaveCategory;
//		Label labelLeaveCount;
//		Label labelLeaveSum;
//		Label labelComplexWorkingTime;
//		Label labelOvertimeCategory;
//		Label labelAllowanceCategory;
//		Label labelBookingRecord;
//		Label labelWorkingItem;
		
		int logSheetPosition = 0;
//		String currentId;
//		String currentName;
		
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
		
		for(int i = 1 ; i < dataArrayList.size(); i++) {
			logSheetPosition++;
			
			checkSearchEnd(logSheet, dataArrayList.get(i), logSheetPosition);
		
			// 檢查是否在假日
			for (int weekdayPosition = 1; weekdayPosition < holidaySheetSize; weekdayPosition++) {
				String holiday = holidaySheet.getCell(0, weekdayPosition).getContents();
//				System.out.println("week day : " + weekday);
				if (dataArrayList.get(i).getAttendanceDate().equals(holiday)) {
					dataArrayList.get(i).setLabelAttribute(EXCEL_CELL_HOLIDAY);
					break;
				}
			}
			
			// 檢查上班狀態 - 有上班 沒下班
			if (!dataArrayList.get(i).getStartWorkingTime().equals("") && dataArrayList.get(i).getEndWorkingTime().length() < 3) {
				dataArrayList.get(i).setLabelAttribute(EXCEL_CELL_NO_OF_DUTY_RECORD);
			}
			
			// 計算上班時數
			List<LeaveDataModel> leaveDataArrayList = dataArrayList.get(i).getLeaveData();
			float leaveSum = leaveDataArrayList.get(leaveDataArrayList.size()).getLeaveSum();
			
			// 檢查上班狀態 - 上班總時數不足
			if (leaveSum < MAX_WORKING_MINUTE) {
				if (dataArrayList.get(i).getStartWorkingTime().length() < 3 && dataArrayList.get(i).getEndWorkingTime().length() < 3) {
					if (dataArrayList.get(i).getComplexWorkingTime().equals("")) {		// 曠職
						dataArrayList.get(i).setLabelAttribute(EXCEL_CELL_IS_ABSENTEEISM);
					} else {	// 工時不足
						dataArrayList.get(i).setLabelAttribute(EXCEL_CELL_NOT_ENOUGH_WORK_TIME);
					}
				} 
			}

			for(int x = 0; x <leaveDataArrayList.size(); x++ ) {
				setMainData(logSheet, dataArrayList.get(i), logSheetPosition);
				setLeaveData(logSheet, leaveDataArrayList.get(x), logSheetPosition);
				
				logSheetPosition++;
			}
//			} else {
				//setMainData(logSheet, dataArrayList.get(i), EXCEL_CELL_HOLIDAY, logSheetPosition);
//			}
			
//			WritableCellFormat excelCellSetting = getExcelCellSetting(dataArrayList.get(i).getLabelAttribute());
			
//			labelGroupId = new Label(0, i, dataArrayList.get(i).getAttendanceGroupId(), excelCellSetting);
//			labelGroupName = new Label(1, i, dataArrayList.get(i).getAttendanceGroupName(), excelCellSetting);
//			labelId = new Label(2, i, dataArrayList.get(i).getAttendanceId(), excelCellSetting);
//			labelName = new Label(3, i, dataArrayList.get(i).getAttendanceName(), excelCellSetting);
//			labelDate = new Label(4, i, dataArrayList.get(i).getAttendanceDate(), excelCellSetting);	
//			
//			labelLateTotalTime = new Label(10, i, dataArrayList.get(i).getLateTotalTime(),excelCellSetting);
//			labelLeaveEarlyTotalTime = new Label(11, i, dataArrayList.get(i).getLeaveEarlyTotalTime(),excelCellSetting);
//			
//			labelComplexWorkingTime = new Label(16, i, dataArrayList.get(i).getComplexWorkingTime(),excelCellSetting);
//			labelOvertimeCategory = new Label(17, i, dataArrayList.get(i).getOvertimeCategory(),excelCellSetting);
//			labelAllowanceCategory = new Label(18, i, dataArrayList.get(i).getAllowanceCategory(),excelCellSetting);
//			labelBookingRecord = new Label(19, i, dataArrayList.get(i).getBookingRecord(),excelCellSetting);
//			labelWorkingItem = new Label(20, i, dataArrayList.get(i).getWorkingItem(),excelCellSetting);
//			
//				String startWorkingTime = dataArrayList.get(i).getStartWorkingTime();
//				String endWorkingTime = dataArrayList.get(i).getEndWorkingTime();
//				
//				long workingMinute = getWorkingMinute(startWorkingTime, endWorkingTime);
//				List<String> workTime = getHourTime(workingMinute * 60);
//				
//				labelStartWorkingTime = new Label(5, i, fromatDate(startWorkingTime, "HH:mm:ss"),excelCellSetting);
//				labelendWorkingTime = new Label(6, i, fromatDate(endWorkingTime, "HH:mm:ss"),excelCellSetting);
//				labelWorkingMinute = new Label(7, i, String.valueOf(workingMinute),excelCellSetting);
//				labelWorkingHours = new Label(8, i, workTime.get(0),excelCellSetting);
//				labelWorkingMin = new Label(9, i, workTime.get(1),excelCellSetting);
				
//				List<LeaveDataModel> leaveDataArrayList = dataArrayList.get(i).getLeaveData();
//				float leaveSum = getWorkingMinute(dataArrayList.get(i).getStartWorkingTime(), dataArrayList.get(i).getEndWorkingTime());
//				for(int x = 0; x <leaveDataArrayList.size(); x++ ) {
//					String excelCellTag = null;
//					
//					// 鎖定最後一筆才顯示加總不足提示
//					if (i == leaveDataArrayList.size() - 1) {
//						
//					}
//					leaveSum += Float.parseFloat(leaveDataArrayList.get(i).getCount()) * 60;
//						
////					String startTime = leaveDataArrayList.get(x).getStartTime();
////					String endTime = leaveDataArrayList.get(x).getEndTime();
////					
////					if (startTime != null && endTime != null) {
////						labelLeaveStatus = new Label(12, logSheetPosition, startTime + "-" + endTime,excelCellSetting);
////					}
////					labelLeaveCategory = new Label(13, logSheetPosition, leaveDataArrayList.get(x).getCategory(),excelCellSetting);
////					labelLeaveCount = new Label(14, logSheetPosition, leaveDataArrayList.get(x).getCount(),excelCellSetting);
////					if (leaveSum > 0) {
////						labelLeaveSum = new Label(15, logSheetPosition, String.valueOf(leaveSum / 60),excelCellSetting);
////					}
//					
//					setLeaveData(logSheet, leaveDataArrayList.get(x), leaveSum, excelCellTag, logSheetPosition);
//					
//					logSheetPosition++;
//				}
			}
			
		if (holidaydayBook != null) {
			holidaydayBook.close();
		}
		
		
		
//		if (!isWeekday && isNoEndWorkingTime) {
//			labelGroupId = new Label(0, logPosition, attendanceGroupId,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelGroupName = new Label(1, logPosition, attendanceGroupName,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelId = new Label(2, logPosition, attendanceId,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelName = new Label(3, logPosition, attendanceName,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelDate = new Label(4, logPosition, attendanceDate,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelStartWorkingTime = new Label(5, logPosition, fromatDate(startWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
//			labelendWorkingTime = new Label(6, logPosition, fromatDate(endWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
//			 
//			// 檢查補刷卡紀錄
////			if (fillAttendanceRecordSheet != null) {
////				loadFillAttendenceRecordDataFromExcel(fillAttendanceRecordSheet, attendanceId, attendanceDate, logSheet, logPosition, isNoEndWorkingTime);
////			}
//			
//			labelWorkingMinute = new Label(7, logPosition, String.valueOf(workingMinute),getWorkingTimeNoEnoughExcelCellSetting(true, true));
//			labelWorkingHours = new Label(8, logPosition, workTime.get(0),getWorkingTimeNoEnoughExcelCellSetting(true, true));
//			labelWorkingMin = new Label(9, logPosition, workTime.get(1),getWorkingTimeNoEnoughExcelCellSetting(true, true));
//			
//			labelLateTotalTime = new Label(10, logPosition, lateTotalTime,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelLeaveEarlyTotalTime = new Label(11, logPosition, leaveEarlyTotalTime,getWorkingTimeNoEnoughExcelCellSetting(false));
//
//			labelComplexWorkingTime = new Label(16, logPosition, complexWorkingTime,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelOvertimeCategory = new Label(17, logPosition, overtimeCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelAllowanceCategory = new Label(18, logPosition, allowanceCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelBookingRecord = new Label(19, logPosition, bookingRecord,getWorkingTimeNoEnoughExcelCellSetting(false));
//			labelWorkingItem = new Label(20, logPosition, workingItem,getWorkingTimeNoEnoughExcelCellSetting(false));
//		} else {
//			labelGroupId = new Label(0, logPosition, attendanceGroupId);
//			labelGroupName = new Label(1, logPosition, attendanceGroupName);
//			labelId = new Label(2, logPosition, attendanceId);
//			labelName = new Label(3, logPosition, attendanceName);
//			
//			if (isWeekday) {
//				labelDate = new Label(4, logPosition, attendanceDate, getWeekdayExcelCellSetting());
//			} else {
//				labelDate = new Label(4, logPosition, attendanceDate);
//			}
//			
//			labelStartWorkingTime = new Label(5, logPosition, fromatDate(startWorkingTime, "HH:mm:ss"),getDateExcelCellSetting());
//			labelendWorkingTime = new Label(6, logPosition, fromatDate(endWorkingTime, "HH:mm:ss"),getDateExcelCellSetting());
//			
//			labelWorkingMinute = new Label(7, logPosition, String.valueOf(workingMinute),getLeaveExcelCellSetting(isWeekday));
//			labelWorkingHours = new Label(8, logPosition, workTime.get(0),getLeaveExcelCellSetting(isWeekday));
//			labelWorkingMin = new Label(9, logPosition, workTime.get(1),getLeaveExcelCellSetting(isWeekday));
//			
//			labelLateTotalTime = new Label(10, logPosition, lateTotalTime);
//			labelLeaveEarlyTotalTime = new Label(11, logPosition, leaveEarlyTotalTime);
//
//			labelComplexWorkingTime = new Label(16, logPosition, complexWorkingTime);
//			labelOvertimeCategory = new Label(17, logPosition, overtimeCategory);
//			labelAllowanceCategory = new Label(18, logPosition, allowanceCategory);
//			labelBookingRecord = new Label(19, logPosition, bookingRecord);
//			labelWorkingItem = new Label(20, logPosition, workingItem);
//		}
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
		Label labelOvertimeCategory = null;
		Label labelStartWorkingTime = null;
		Label labelendWorkingTime= null;
		
		String startWorkingTime = data.getStartWorkingTime();
		String endWorkingTime = data.getEndWorkingTime();
		
		long workingMinute = getWorkingMinute(startWorkingTime, endWorkingTime);
		List<String> workTime = getHourTime(workingMinute * 60);
		
		switch (data.getLabelAttribute()) {
			case EXCEL_CELL_HOLIDAY:		// 假日
				labelDate = new Label(4, position, data.getAttendanceDate(), getHolidayExcelCellSetting());
				labelWorkingMinute = new Label(7, position, String.valueOf(workingMinute), getLeaveExcelCellSetting(true));
				labelWorkingHours = new Label(8, position, workTime.get(0), getLeaveExcelCellSetting(true));
				labelWorkingMin = new Label(9, position, workTime.get(1), getLeaveExcelCellSetting(true));
				
				break;
			case EXCEL_CELL_NOT_ENOUGH_WORK_TIME:		// 上班時數不足
				labelName = new Label(3, position, data.getAttendanceName(), getWorkingTimeNoEnoughExcelCellSetting(false));
				
				break;
			case EXCEL_CELL_IS_ABSENTEEISM:	// 曠職
				labelName = new Label(3, position, data.getAttendanceName(), getWorkingTimeNoEnoughExcelCellSetting(false));
				labelStartWorkingTime = new Label(5, position, fromatDate(startWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
				labelendWorkingTime = new Label(6, position, fromatDate(endWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
				labelOvertimeCategory = new Label(17, position, data.getOvertimeCategory(),getWorkingTimeNoEnoughExcelCellSetting(false));
	
				break;
			case EXCEL_CELL_NO_OF_DUTY_RECORD:		// 沒有下班卡紀錄
				labelName = new Label(3, position, data.getAttendanceName(), getWorkingTimeNoEnoughExcelCellSetting(false));
				labelendWorkingTime = new Label(6, position, fromatDate(endWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
				
				break;
			default:
				break;
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
			labelStartWorkingTime = new Label(5, position, fromatDate(startWorkingTime, "HH:mm:ss"));
		}
		if(labelendWorkingTime == null) {
			labelendWorkingTime = new Label(6, position, fromatDate(endWorkingTime, "HH:mm:ss"));
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
		
		Label labelComplexWorkingTime = new Label(16, position, data.getComplexWorkingTime());
		if(labelOvertimeCategory == null) {
			labelOvertimeCategory = new Label(17, position, data.getOvertimeCategory());
		}
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
	
	private void setLeaveData(WritableSheet logSheet, LeaveDataModel data, float leaveSum, String excelCellTag, int position) {
		WritableCellFormat excelCellFormat = null;
		
//		switch (excelCellTag) {
//		case 
//			
//		}
		
		String startTime = data.getStartTime();
		String endTime = data.getEndTime();
		
		Label labelLeaveStatus = null;
		if (startTime != null && endTime != null) {
			labelLeaveStatus = new Label(12, position, startTime + "-" + endTime,excelCellFormat);
		}
		Label labelLeaveCategory = new Label(13, position, data.getCategory(),excelCellFormat);
		Label labelLeaveCount = new Label(14, position, data.getCount(),excelCellFormat);
		Label labelLeaveSum = null;
		if (leaveSum > 0) {
			labelLeaveSum = new Label(15, position, String.valueOf(leaveSum / 60),excelCellFormat);
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
	
	private WritableCellFormat getExcelCellFormat(String excelCellTag) {
		WritableCellFormat excelCellFormat = null;
		
		switch (excelCellTag) {
		case EXCEL_CELL_HOLIDAY:
			return excelCellFormat;
		case EXCEL_CELL_DEFAULT:
			return excelCellFormat;
		case EXCEL_CELL_NOT_ENOUGH_WORK_TIME:
			return getWorkingTimeNoEnoughExcelCellSetting(false);	
		case EXCEL_CELL_IS_ABSENTEEISM:
			return excelCellFormat;	
		case EXCEL_CELL_NO_OF_DUTY_RECORD:
			return excelCellFormat;
		case EXCEL_CELL_ERROR_USER:
			return excelCellFormat;
		default:
			return excelCellFormat;
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
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			// 參閱 http://www.cnblogs.com/smilsy/articles/2126377.html
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
