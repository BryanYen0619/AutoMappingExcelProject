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
	
	private String attendancePath;
	private String leavePath;
	private String logPath;
	private String fillAttendanceRecordPath;
	
	private String attendanceGroupId;
	private String attendanceGroupName;
	private String attendanceId;
	private String attendanceName;
	private String attendanceDate;
	private String startWorkingTime;		// YYY-MM-dd HH:mm:ss
	private String endWorkingTime;		// YYY-MM-dd HH:mm:ss
	private String lateTotalTime;
	private String leaveEarlyTotalTime;
//	private String leaveStatus;
	private String complexWorkingTime;
	private String overtimeCategory;
	private String allowanceCategory;
	private String bookingRecord;
	private String workingItem;
	private String leaveId;
	private String leaveDate;
	private String currentId = null;
	private String currentName = null;
	private String path = "";	// 目前檔案路徑
	private String weekdayPath;		// Excel 休假日期 路徑
	
	private long workingMinute = 0;
	private int logPosition = 1;
	private int currentCount = 1;
	
	private boolean isWeekDay = false;
	private boolean isNoEndWorkingTime = false;
	
	public AutoMappingModel() {
		
	}
	
	public void run() {
		loadWeekdayExcelData();
		
		try {
			Workbook workbook = Workbook.getWorkbook(new File(attendancePath));
			Sheet attendanceSheet = workbook.getSheet(0);
			int attendanceSheetSize = attendanceSheet.getRows();	
//			System.out.println("Attendance Excel Size : " + attendanceSheetSize);		// 共幾筆
			
			Workbook leaveBook = Workbook.getWorkbook(new File(leavePath));
			Sheet leaveSheet = leaveBook.getSheet(0);
			int leaveSheetSize = leaveSheet.getRows();	
//			System.out.println("Leave Excel Size : " + leaveSheetSize);		// 共幾筆
			
			Workbook weekdayBook = Workbook.getWorkbook(new File(weekdayPath));
			Sheet weekdaySheet = weekdayBook.getSheet(0);
			int weekdaySheetSize = weekdaySheet.getRows();	
//			System.out.println("Weekday Excel Size : " + weekdaySheetSize);		// 共幾筆
			
			int fillAttendanceRecordSheetSize = 0;
			Sheet fillAttendanceRecordSheet = null;
			if (fillAttendanceRecordPath != null) {
				Workbook fillAttendanceRecordBook = Workbook.getWorkbook(new File(fillAttendanceRecordPath));
				fillAttendanceRecordSheet = fillAttendanceRecordBook.getSheet(0);
				fillAttendanceRecordSheetSize = fillAttendanceRecordSheet.getRows();
//				System.out.println("Fill Attendance Record Excel Size : " + fillAttendanceRecordSheetSize);		// 共幾筆
			}
			
			WritableWorkbook logBook = Workbook.createWorkbook(new File(logPath));
			WritableSheet logSheet = initLogExcel(logBook, OUTPUT_LOG_EXCEL_SHEET_NAME);
	           
			for (int position = 1; position < attendanceSheetSize; position++) {
				loadAttendenceDataFromExcel(attendanceSheet, position);		// 讀取門禁資料
				
				if (fillAttendanceRecordSheetSize > 0) {	// 讀取補刷卡資料
//					loadFillAttendenceRecordDataFromExcel(fillAttendanceRecordSheet, position);
				}
				
				if (!startWorkingTime.equals("") && endWorkingTime.length() < 3) {
					isNoEndWorkingTime = true;
				} else {
					isNoEndWorkingTime = false;
				}
				
				if (currentId == null) {
					currentId = attendanceId;
					currentName = attendanceName;
				}
				
				checkSearchEnd(logSheet);

				// 檢查是否在假日
				String weekday;
				isWeekDay = false;
				for (int weekdayPosition = 1; weekdayPosition < weekdaySheetSize; weekdayPosition++) {
					weekday = weekdaySheet.getCell(0, weekdayPosition).getContents();
//					System.out.println("week day : " + weekday);
					if (attendanceDate.equals(weekday)) {
						isWeekDay = true;
						break;
					}
				}
				
				setLogExcelData(logSheet, logPosition, isWeekDay, isNoEndWorkingTime);
				
				float leaveSum = workingMinute;
				List<LeaveDataModel> leaveDataModels = new ArrayList();
				
				for (int search = 1; search < leaveSheetSize; search++) {
					float tempLeaveSum = leaveSum;

					leaveId = leaveSheet.getCell(2, search ).getContents();
					leaveDate = leaveSheet.getCell(5,search).getContents();
					
//						System.out.println("search "+search+", id : " + leaveId);
//						System.out.println("search "+search+", date : " + leaveDate);
					
					if (attendanceId.equals(leaveId)) {
						if (attendanceDate.equals(leaveDate)) {
							LeaveDataModel leaveDataModel = new LeaveDataModel();
							
							leaveDataModel.setPosition(logPosition);
							leaveDataModel.isNoEndWorkingTime(isNoEndWorkingTime);
							
							String leaveCategory = leaveSheet.getCell(4,search).getContents();
							String leaveCount = leaveSheet.getCell(10,search).getContents();
						
							leaveDataModel.setCategory(leaveCategory);
							leaveDataModel.setCount(leaveCount);
							
//								System.out.println("leaveCategory : " + leaveCategory);
//								System.out.println("leaveCount : " + leaveCount);
							
							 String startTime = leaveSheet.getCell(7,search).getContents();
							 String endTime = leaveSheet.getCell(9, search).getContents();
							 
							// 用不到 暫時關閉 Begin.
							// long leaveMinute = getLeaveMinute(startTime, endTime);
							// End.
							 
							 leaveDataModel.setStartTime(startTime);
							 leaveDataModel.setEndTime(endTime);
							
//								System.out.println("search "+search+", start time : " + startTime);
//								System.out.println("search "+search+", end time : " + endTime);
//								System.out.println("search "+search+", leave minute : " + leaveMinute);
							
							leaveSum = tempLeaveSum + 60 +Float.parseFloat(leaveCount) * 60;
							leaveDataModel.setLeaveSum(leaveSum);
//								System.out.println("leave Sum : " + leaveSum);
							leaveDataModels.add(leaveDataModel);
							
							setLogExcelData(logSheet, logPosition, isWeekDay, isNoEndWorkingTime);
							
							logPosition++;
							currentCount++;
						}
					}
				}

				if (leaveDataModels.size() > 0) {
					for(int i = 0; i < leaveDataModels.size(); i++) {
						// 鎖定最後一筆才顯示加總不足提示
						boolean isEndPosition = false;
						if (i == leaveDataModels.size() - 1) {
							isEndPosition = true;
						}
						
						setLogExcelDataFromLeaveData(logSheet, leaveDataModels.get(i).getPosition(), leaveDataModels.get(i).getStartTime(),
								leaveDataModels.get(i).getEndTime(),leaveDataModels.get(i).getCategory(),leaveDataModels.get(i).getCount(), 
								leaveDataModels.get(i).getLeaveSum(), false, leaveDataModels.get(i).isNoEndWorkingTime(), isEndPosition);
					}
				} else {
					setLogExcelDataFromLeaveData(logSheet, logPosition, null, null, null, null, 0, true, isNoEndWorkingTime, true);
					logPosition++;
					currentCount++;
				}
			}

			logBook.write();
			logBook.close();
			
			weekdayBook.close();
			leaveBook.close();
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
	
	private void loadWeekdayExcelData() {
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
		
		weekdayPath = path+"107ygov_yfyshop_weekday.xls";
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
	private void loadAttendenceDataFromExcel(Sheet attendanceSheet, int position) {
		attendanceGroupId = attendanceSheet.getCell(0, position ).getContents();
		attendanceGroupName = attendanceSheet.getCell(1, position ).getContents();
		attendanceId = attendanceSheet.getCell(2, position ).getContents();
		attendanceName = attendanceSheet.getCell(3, position ).getContents();
		attendanceDate = attendanceSheet.getCell(4,position).getContents();
		startWorkingTime = attendanceSheet.getCell(5,position).getContents();
		endWorkingTime = attendanceSheet.getCell(6, position).getContents();
		lateTotalTime = attendanceSheet.getCell(7, position).getContents();	// 遲到
		leaveEarlyTotalTime = attendanceSheet.getCell(8, position).getContents();		// 早退
//		leaveStatus = attendanceSheet.getCell(9, position).getContents();
		complexWorkingTime = attendanceSheet.getCell(10, position).getContents();
		overtimeCategory = attendanceSheet.getCell(11, position).getContents();
		allowanceCategory = attendanceSheet.getCell(12, position).getContents();
		bookingRecord = attendanceSheet.getCell(13, position).getContents();
		workingItem = attendanceSheet.getCell(14, position).getContents();
	}
	
	private void checkSearchEnd(WritableSheet logSheet) {
		if (!currentId.equals(attendanceId)) {
//			System.out.println("get currentName : " + currentName);
//			System.out.println("get current : " + currentCount  - 1);
			
			Label labelEndName = new Label(3, logPosition, currentName + " 計數", getEndInfoExcelCellSetting());
			Label labelEndCount = new Label(20, logPosition, String.valueOf(currentCount - 1), getDateExcelCellSetting());
			
			logPosition++;
			
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
			
			currentId = attendanceId;
			currentName = attendanceName;
			currentCount = 1;
			isNoEndWorkingTime = false;
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
	private static WritableSheet initLogExcel(WritableWorkbook logBook, String logSheetName) {
		WritableSheet logSheet = null;
		try {
			logSheet = logBook.createSheet(logSheetName, 0);
			
			// 初始Title設定
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
		
		return logSheet;
	}
	
	private void setLogExcelData(WritableSheet logSheet, int logPosition, boolean isWeekday, boolean isNoEndWorkingTime) {
		Label labelGroupId;
		Label labelGroupName;
		Label labelId;
		Label labelName;
		Label labelDate;
		Label labelStartWorkingTime;
		Label labelendWorkingTime;
		Label labelWorkingMinute;
		Label labelWorkingHours;
		Label labelWorkingMin;
		Label labelLateTotalTime;
		Label labelLeaveEarlyTotalTime;
		Label labelComplexWorkingTime;
		Label labelOvertimeCategory;
		Label labelAllowanceCategory;
		Label labelBookingRecord;
		Label labelWorkingItem;
		
		workingMinute = getWorkingMinute(startWorkingTime, endWorkingTime);
		List<String> workTime = getHourTime(workingMinute * 60);
		
		if (isNoEndWorkingTime) {
			labelGroupId = new Label(0, logPosition, attendanceGroupId,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelGroupName = new Label(1, logPosition, attendanceGroupName,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelId = new Label(2, logPosition, attendanceId,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelName = new Label(3, logPosition, attendanceName,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelDate = new Label(4, logPosition, attendanceDate,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelStartWorkingTime = new Label(5, logPosition, fromatDate(startWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
			labelendWorkingTime = new Label(6, logPosition, fromatDate(endWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
			 
			labelWorkingMinute = new Label(7, logPosition, String.valueOf(workingMinute),getWorkingTimeNoEnoughExcelCellSetting(true, true));
			labelWorkingHours = new Label(8, logPosition, workTime.get(0),getWorkingTimeNoEnoughExcelCellSetting(true, true));
			labelWorkingMin = new Label(9, logPosition, workTime.get(1),getWorkingTimeNoEnoughExcelCellSetting(true, true));
			
			labelLateTotalTime = new Label(10, logPosition, lateTotalTime,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelLeaveEarlyTotalTime = new Label(11, logPosition, leaveEarlyTotalTime,getWorkingTimeNoEnoughExcelCellSetting(false));

			labelComplexWorkingTime = new Label(16, logPosition, complexWorkingTime,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelOvertimeCategory = new Label(17, logPosition, overtimeCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelAllowanceCategory = new Label(18, logPosition, allowanceCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelBookingRecord = new Label(19, logPosition, bookingRecord,getWorkingTimeNoEnoughExcelCellSetting(false));
			labelWorkingItem = new Label(20, logPosition, workingItem,getWorkingTimeNoEnoughExcelCellSetting(false));
		} else {
			labelGroupId = new Label(0, logPosition, attendanceGroupId);
			labelGroupName = new Label(1, logPosition, attendanceGroupName);
			labelId = new Label(2, logPosition, attendanceId);
			labelName = new Label(3, logPosition, attendanceName);
			
			if (isWeekday) {
				labelDate = new Label(4, logPosition, attendanceDate, getWeekdayExcelCellSetting());
			} else {
				labelDate = new Label(4, logPosition, attendanceDate);
			}
			
			labelStartWorkingTime = new Label(5, logPosition, fromatDate(startWorkingTime, "HH:mm:ss"),getDateExcelCellSetting());
			labelendWorkingTime = new Label(6, logPosition, fromatDate(endWorkingTime, "HH:mm:ss"),getDateExcelCellSetting());
			
			labelWorkingMinute = new Label(7, logPosition, String.valueOf(workingMinute),getLeaveExcelCellSetting(isWeekday));
			labelWorkingHours = new Label(8, logPosition, workTime.get(0),getLeaveExcelCellSetting(isWeekday));
			labelWorkingMin = new Label(9, logPosition, workTime.get(1),getLeaveExcelCellSetting(isWeekday));
			
			labelLateTotalTime = new Label(10, logPosition, lateTotalTime);
			labelLeaveEarlyTotalTime = new Label(11, logPosition, leaveEarlyTotalTime);

			labelComplexWorkingTime = new Label(16, logPosition, complexWorkingTime);
			labelOvertimeCategory = new Label(17, logPosition, overtimeCategory);
			labelAllowanceCategory = new Label(18, logPosition, allowanceCategory);
			labelBookingRecord = new Label(19, logPosition, bookingRecord);
			labelWorkingItem = new Label(20, logPosition, workingItem);
		}
		
		try {
			logSheet.addCell(labelGroupId);
			logSheet.addCell(labelGroupName); 
			logSheet.addCell(labelId); 
			logSheet.addCell(labelName); 
			logSheet.addCell(labelStartWorkingTime); 
			logSheet.addCell(labelendWorkingTime); 
			logSheet.addCell(labelLateTotalTime); 
			logSheet.addCell(labelWorkingMinute); 
			logSheet.addCell(labelWorkingHours);
			logSheet.addCell(labelWorkingMin);
			logSheet.addCell(labelLeaveEarlyTotalTime); 
			logSheet.addCell(labelComplexWorkingTime); 
			logSheet.addCell(labelOvertimeCategory); 
			logSheet.addCell(labelAllowanceCategory); 
			logSheet.addCell(labelBookingRecord); 
			logSheet.addCell(labelWorkingItem); 
			logSheet.addCell(labelDate);
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
	private static void setLogExcelDataFromLeaveData(WritableSheet logSheet, int logPosition, String startTime, String endTime, String leaveCategory, String leaveCount, float leaveSum, boolean isNullData, boolean isNoEndWorkingTime, boolean isWorkingTimeNotEnough) {
		Label labelLeaveStatus = null;
		Label labelLeaveCategory = null;
		Label labelLeaveCount = null;
		Label labelLeaveSum = null;
		
		if (isNoEndWorkingTime) {
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
					if (leaveSum < MAX_WORKING_MINUTE) {
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
	
	private static WritableCellFormat getWeekdayExcelCellSetting() {
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
