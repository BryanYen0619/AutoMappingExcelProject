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

public class AutoMappingModel {
	final long MAX_WORKING_MINUTE = 540; // 60 * 9
	
	private String attendancePath;
	private String leavePath;
	private String logPath;
	
	public AutoMappingModel() {
		
	}
	
	public void run() {
		File directory = new File(".");//设定为当前文件夹
		//System.out.println(directory.getCanonicalPath());//获取标准的路径
		//System.out.println(directory.getAbsolutePath());//获取绝对路径
		String path="";
		try {
			path = directory.getCanonicalPath()+"/src/";
			System.out.println("Project Path : " + path);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}	
		
		// Excel 出勤紀錄 路徑
//		String attendancePath = path+"/attendance_record.xls";
		// Excel 分頁名稱
		// String attendanceSheetName = "Table1";	
		
		// Excel 請假紀錄 路徑
//		String leavePath = path+"/leave_record.xls";
		// Excel 分頁名稱
		// String leavePathName = "Table";
		
		// Excel Log 路徑
//		String logPath = path+"/log_record.xls";
		// Excel 分頁名稱
		String logPathName = "Sheet";
		
		// Excel 休假日期 路徑
		String weekdayPath = path+"107ygov_yfyshop_weekday.xls";
		// Excel 分頁名稱
		// String weekdayPathName = "Sheet";
		
		if (attendancePath == null && leavePath == null) {
			
		}
		
		try {
			Workbook workbook = Workbook.getWorkbook(new File(attendancePath));
			Sheet attendanceSheet = workbook.getSheet(0);
			
			Workbook leaveBook = Workbook.getWorkbook(new File(leavePath));
			Sheet leaveSheet = leaveBook.getSheet(0);
			
			Workbook weekdayBook = Workbook.getWorkbook(new File(weekdayPath));
			Sheet weekdaySheet = weekdayBook.getSheet(0);
			
			WritableWorkbook logBook = Workbook.createWorkbook(new File(logPath));
			WritableSheet logSheet = logBook.createSheet(logPathName, 0);
			
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
	           
			int attendanceSheetSize = attendanceSheet.getRows();	
			System.out.println("Attendance Excel Size : " + attendanceSheetSize);		// 共幾筆
			
			int leaveSheetSize = leaveSheet.getRows();	
			System.out.println("Leave Excel Size : " + leaveSheetSize);		// 共幾筆
			
			int weekdaySheetSize = weekdaySheet.getRows();	
			System.out.println("Weekday Excel Size : " + weekdaySheetSize);		// 共幾筆
			
			String attendanceGroupId;
			String attendanceGroupName;
			String attendanceId;
			String attendanceName;
			String attendanceDate;
			String startWorkingTime;		// YYY-MM-dd HH:mm:ss
			String endWorkingTime;		// YYY-MM-dd HH:mm:ss
			String lateTotalTime;
			String leaveEarlyTotalTime;
			String leaveStatus;
			String complexWorkingTime;
			String overtimeCategory;
			String allowanceCategory;
			String bookingRecord;
			String workingItem;
			long workingMinute = 0;
			
			String leaveId;
			String leaveDate;
			String startTime;		// mm:ss
			String endTime;		// mm:ss
			String leaveCategory;
			String leaveCount;
			long leaveMinute = 0;
			
			int logPosition = 1;
			
			String currentId = null;
			String currentName = null;
			int currentCount = 1;
			
			for (int position = 1; position < attendanceSheetSize; position++) {
				boolean isNoEndWorkingTime = false;
				boolean isWorkingTimeNotEnough = false;
				
				attendanceGroupId = attendanceSheet.getCell(0, position ).getContents();
				attendanceGroupName = attendanceSheet.getCell(1, position ).getContents();
				attendanceId = attendanceSheet.getCell(2, position ).getContents();
				attendanceName = attendanceSheet.getCell(3, position ).getContents();
				attendanceDate = attendanceSheet.getCell(4,position).getContents();
				startWorkingTime = attendanceSheet.getCell(5,position).getContents();
				endWorkingTime = attendanceSheet.getCell(6, position).getContents();
				lateTotalTime = attendanceSheet.getCell(7, position).getContents();	// 遲到
				leaveEarlyTotalTime = attendanceSheet.getCell(8, position).getContents();		// 早退
				leaveStatus = attendanceSheet.getCell(9, position).getContents();
				complexWorkingTime = attendanceSheet.getCell(10, position).getContents();
				overtimeCategory = attendanceSheet.getCell(11, position).getContents();
				allowanceCategory = attendanceSheet.getCell(12, position).getContents();
				bookingRecord = attendanceSheet.getCell(13, position).getContents();
				workingItem = attendanceSheet.getCell(14, position).getContents();
				
				if (!startWorkingTime.equals("") && endWorkingTime.length() < 3) {
					isNoEndWorkingTime = true;
				}
				
				if (currentId == null) {
					currentId = attendanceId;
					currentName = attendanceName;
				}
				
				if (!currentId.equals(attendanceId)) {
//					System.out.println("get currentName : " + currentName);
//					System.out.println("get current : " + currentCount  - 1);
					
					Label labelEndName = new Label(3, logPosition, currentName + " 計數", getEndInfoExcelCellSetting());
					Label labelEndCount = new Label(20, logPosition, String.valueOf(currentCount - 1), getDateExcelCellSetting());
					
					logPosition++;
					
					logSheet.addCell(labelEndName); 
					logSheet.addCell(labelEndCount); 
					
					currentId = attendanceId;
					currentName = attendanceName;
					currentCount = 1;
					isNoEndWorkingTime = false;
				}
				
//				System.out.println("endWorkingTime : "+ endWorkingTime);
				
				Label labelGroupId = new Label(0, logPosition, attendanceGroupId);
				Label labelGroupName = new Label(1, logPosition, attendanceGroupName);
				Label labelId = new Label(2, logPosition, attendanceId);
				Label labelName = new Label(3, logPosition, attendanceName);
				Label labelDate = new Label(4, logPosition, attendanceDate);
				Label labelStartWorkingTime = new Label(5, logPosition, fromatDate(startWorkingTime, "HH:mm:ss"),getDateExcelCellSetting());
				Label labelendWorkingTime = new Label(6, logPosition, fromatDate(endWorkingTime, "HH:mm:ss"),getDateExcelCellSetting());
				Label labelLateTotalTime = new Label(10, logPosition, lateTotalTime);
				Label labelLeaveEarlyTotalTime = new Label(11, logPosition, leaveEarlyTotalTime);
				
				Label labelLeaveStatus;
				if (leaveStatus.length() > 11) {
					labelLeaveStatus = new Label(12, logPosition, leaveStatus.substring(0, 11));
				} else {
					labelLeaveStatus = new Label(12, logPosition, leaveStatus);
				}
				
				Label labelComplexWorkingTime = new Label(16, logPosition, complexWorkingTime);
				Label labelOvertimeCategory = new Label(17, logPosition, overtimeCategory);
				Label labelAllowanceCategory = new Label(18, logPosition, allowanceCategory);
				Label labelBookingRecord = new Label(19, logPosition, bookingRecord);
				Label labelWorkingItem = new Label(20, logPosition, workingItem);
				
				workingMinute = getWorkingMinute(startWorkingTime, endWorkingTime);
				List<String> workTime = getHourTime(workingMinute * 60);
				
				Label labelWorkingMinute = new Label(7, logPosition, String.valueOf(workingMinute),getLeaveExcelCellSetting());
				Label labelWorkingHours = new Label(8, logPosition, workTime.get(0),getLeaveExcelCellSetting());
				Label labelWorkingMin = new Label(9, logPosition, workTime.get(1),getLeaveExcelCellSetting());
		
				// 檢查是否在假日
				String weekday;
				for (int weekdayPosition = 1; weekdayPosition < weekdaySheetSize; weekdayPosition++) {
					weekday = weekdaySheet.getCell(0, weekdayPosition).getContents();
//					System.out.println("week day : " + weekday);
					if (attendanceDate.equals(weekday)) {
						labelDate = new Label(4, logPosition, attendanceDate, getWeekdayExcelCellSetting());
						isNoEndWorkingTime = false;
					}
				}
				
				if (isNoEndWorkingTime) {
					labelGroupId = new Label(0, logPosition, attendanceGroupId,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelGroupName = new Label(1, logPosition, attendanceGroupName,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelId = new Label(2, logPosition, attendanceId,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelName = new Label(3, logPosition, attendanceName,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelDate = new Label(4, logPosition, attendanceDate,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelStartWorkingTime = new Label(5, logPosition, fromatDate(startWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
					labelendWorkingTime = new Label(6, logPosition, fromatDate(endWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
					 
					labelLateTotalTime = new Label(10, logPosition, lateTotalTime,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelLeaveEarlyTotalTime = new Label(11, logPosition, leaveEarlyTotalTime,getWorkingTimeNoEnoughExcelCellSetting(false));

					if (leaveStatus.length() > 11) {
						labelLeaveStatus = new Label(12, logPosition, leaveStatus.substring(0, 11),getWorkingTimeNoEnoughExcelCellSetting(false));
					} else {
						labelLeaveStatus = new Label(12, logPosition, leaveStatus,getWorkingTimeNoEnoughExcelCellSetting(false));
					}
					
					labelComplexWorkingTime = new Label(16, logPosition, complexWorkingTime,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelOvertimeCategory = new Label(17, logPosition, overtimeCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelAllowanceCategory = new Label(18, logPosition, allowanceCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelBookingRecord = new Label(19, logPosition, bookingRecord,getWorkingTimeNoEnoughExcelCellSetting(false));
					labelWorkingItem = new Label(20, logPosition, workingItem,getWorkingTimeNoEnoughExcelCellSetting(false));
					
					labelWorkingMinute = new Label(7, logPosition, String.valueOf(workingMinute),getWorkingTimeNoEnoughExcelCellSetting(true, true));
					labelWorkingHours = new Label(8, logPosition, workTime.get(0),getWorkingTimeNoEnoughExcelCellSetting(true, true));
					labelWorkingMin = new Label(9, logPosition, workTime.get(1),getWorkingTimeNoEnoughExcelCellSetting(true, true));
				}

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
				logSheet.addCell(labelLeaveStatus); 
				logSheet.addCell(labelComplexWorkingTime); 
				logSheet.addCell(labelOvertimeCategory); 
				logSheet.addCell(labelAllowanceCategory); 
				logSheet.addCell(labelBookingRecord); 
				logSheet.addCell(labelWorkingItem); 
				logSheet.addCell(labelDate); 
				
				boolean isSearch = false;
				float leaveSum = workingMinute;
				
				Label labelLeaveCategory = null;
				Label labelLeaveCount = null;
				Label labelLeaveSum = null;
				
				int leaveRecordCount = 0;
				for (int search = 1; search < leaveSheetSize; search++) {
					leaveRecordCount++;
				
					leaveId = leaveSheet.getCell(2, search ).getContents();
					leaveDate = leaveSheet.getCell(5,search).getContents();
					
//						System.out.println("search "+search+", id : " + leaveId);
//						System.out.println("search "+search+", date : " + leaveDate);
					
					float tempLeaveSum = leaveSum;
					
					if (attendanceId.equals(leaveId)) {
						if (attendanceDate.equals(leaveDate)) {
							leaveCategory = leaveSheet.getCell(4,search).getContents();
							leaveCount = leaveSheet.getCell(10,search).getContents();
							
							
//								System.out.println("leaveCategory : " + leaveCategory);
//								System.out.println("leaveCount : " + leaveCount);
							
							startTime = leaveSheet.getCell(7,search).getContents();
							endTime = leaveSheet.getCell(9, search).getContents();
							leaveMinute = getLeaveMinute(startTime, endTime);
							
//								System.out.println("search "+search+", start time : " + startTime);
//								System.out.println("search "+search+", end time : " + endTime);
//								System.out.println("search "+search+", leave minute : " + leaveMinute);
							leaveSum = tempLeaveSum + Float.parseFloat(leaveCount) * 60;
							
//							System.out.println("leave Sum : " + leaveSum);
							
							if (attendanceId.equals("201700590")) {
								System.out.println("logPosition : " + logPosition);
								System.out.println("leaveCount : " + leaveCount);
							}
							
							if (leaveRecordCount > 1) {
								if (isNoEndWorkingTime) {
									labelGroupId = new Label(0, logPosition, attendanceGroupId,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelGroupName = new Label(1, logPosition, attendanceGroupName,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelId = new Label(2, logPosition, attendanceId,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelName = new Label(3, logPosition, attendanceName,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelDate = new Label(4, logPosition, attendanceDate,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelStartWorkingTime = new Label(5, logPosition, fromatDate(startWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
									labelendWorkingTime = new Label(6, logPosition, fromatDate(endWorkingTime, "HH:mm:ss"),getWorkingTimeNoEnoughExcelCellSetting(true));
									
									 labelLateTotalTime = new Label(10, logPosition, lateTotalTime,getWorkingTimeNoEnoughExcelCellSetting(false));
									 labelLeaveEarlyTotalTime = new Label(11, logPosition, leaveEarlyTotalTime,getWorkingTimeNoEnoughExcelCellSetting(false));
									
									if (leaveStatus.length() > 11) {
										labelLeaveStatus = new Label(12, logPosition, leaveStatus.substring(0, 11),getWorkingTimeNoEnoughExcelCellSetting(false));
									} else {
										labelLeaveStatus = new Label(12, logPosition, leaveStatus,getWorkingTimeNoEnoughExcelCellSetting(false));
									}
									
									labelLeaveCategory = new Label(13, logPosition, leaveCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelLeaveCount = new Label(14, logPosition, leaveCount,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelLeaveSum = new Label(15, logPosition, String.valueOf(leaveSum / 60),getWorkingTimeNoEnoughExcelCellSetting(false));
									
									labelComplexWorkingTime = new Label(16, logPosition, complexWorkingTime,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelOvertimeCategory = new Label(17, logPosition, overtimeCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelAllowanceCategory = new Label(18, logPosition, allowanceCategory,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelBookingRecord = new Label(19, logPosition, bookingRecord,getWorkingTimeNoEnoughExcelCellSetting(false));
									labelWorkingItem = new Label(20, logPosition, workingItem,getWorkingTimeNoEnoughExcelCellSetting(false));
									
									labelWorkingMinute = new Label(7, logPosition, String.valueOf(workingMinute),getWorkingTimeNoEnoughExcelCellSetting(true, true));
									labelWorkingHours = new Label(8, logPosition, workTime.get(0),getWorkingTimeNoEnoughExcelCellSetting(true, true));
									labelWorkingMin = new Label(9, logPosition, workTime.get(1),getWorkingTimeNoEnoughExcelCellSetting(true, true));
								} else {
									labelGroupId = new Label(0, logPosition, attendanceGroupId);
									labelGroupName = new Label(1, logPosition, attendanceGroupName);
									labelId = new Label(2, logPosition, attendanceId);
									labelName = new Label(3, logPosition, attendanceName);
									labelDate = new Label(4, logPosition, attendanceDate);
									labelStartWorkingTime = new Label(5, logPosition, fromatDate(startWorkingTime, "HH:mm:ss"),getDateExcelCellSetting());
									labelendWorkingTime = new Label(6, logPosition, fromatDate(endWorkingTime, "HH:mm:ss"),getDateExcelCellSetting());
									 
									 labelLateTotalTime = new Label(10, logPosition, lateTotalTime);
									 labelLeaveEarlyTotalTime = new Label(11, logPosition, leaveEarlyTotalTime);
									
									if (leaveStatus.length() > 11) {
										labelLeaveStatus = new Label(12, logPosition, leaveStatus.substring(0, 11));
									} else {
										labelLeaveStatus = new Label(12, logPosition, leaveStatus);
									}
									
									labelLeaveCategory = new Label(13, logPosition, leaveCategory);
									labelLeaveCount = new Label(14, logPosition, leaveCount);
									labelLeaveSum = new Label(15, logPosition, String.valueOf(leaveSum / 60));
									
									labelComplexWorkingTime = new Label(16, logPosition, complexWorkingTime);
									labelOvertimeCategory = new Label(17, logPosition, overtimeCategory);
									labelAllowanceCategory = new Label(18, logPosition, allowanceCategory);
									labelBookingRecord = new Label(19, logPosition, bookingRecord);
									labelWorkingItem = new Label(20, logPosition, workingItem);
									
									labelWorkingMinute = new Label(7, logPosition, String.valueOf(workingMinute),getLeaveExcelCellSetting());
									labelWorkingHours = new Label(8, logPosition, workTime.get(0),getLeaveExcelCellSetting());
									labelWorkingMin = new Label(9, logPosition, workTime.get(1),getLeaveExcelCellSetting());
								}
								
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
								logSheet.addCell(labelLeaveStatus); 
								logSheet.addCell(labelComplexWorkingTime); 
								logSheet.addCell(labelOvertimeCategory); 
								logSheet.addCell(labelAllowanceCategory); 
								logSheet.addCell(labelBookingRecord); 
								logSheet.addCell(labelWorkingItem); 
								logSheet.addCell(labelDate);
								logSheet.addCell(labelLeaveCategory);
								logSheet.addCell(labelLeaveCount); 
								logSheet.addCell(labelLeaveSum); 
							}
							
							logPosition++;
							currentCount++;
							isSearch = true;
						}
					}
				}

				if (!isSearch) {
					if (isNoEndWorkingTime) {
						labelLeaveCategory = new Label(13, logPosition, "",getWorkingTimeNoEnoughExcelCellSetting(false));
						labelLeaveCount = new Label(14, logPosition, "",getWorkingTimeNoEnoughExcelCellSetting(false));
						labelLeaveSum = new Label(15, logPosition, "",getWorkingTimeNoEnoughExcelCellSetting(false));
					
						logSheet.addCell(labelLeaveCategory);
						logSheet.addCell(labelLeaveCount); 
						logSheet.addCell(labelLeaveSum); 
					}
					
					logPosition++;
					currentCount++;
				} 
			}

			logBook.write();
			logBook.close();
			
			weekdayBook.close();
			leaveBook.close();
			workbook.close();
			
			Alert alert = new Alert(AlertType.INFORMATION);
	        alert.setTitle("提示");
	        alert.setHeaderText("完成");
	        alert.setContentText("已產出結果!");
	        alert.showAndWait();
	        
		} catch (BiffException e) {
			e.printStackTrace();
			
			Alert alert = new Alert(AlertType.INFORMATION);
	        alert.setTitle("提示");
	        alert.setHeaderText("失敗");
	        alert.setContentText( e.getMessage());
	        alert.showAndWait();
		} catch (IOException e) {
			e.printStackTrace();
			
			Alert alert = new Alert(AlertType.INFORMATION);
	        alert.setTitle("提示");
	        alert.setHeaderText("失敗");
	        alert.setContentText( e.getMessage());
	        alert.showAndWait();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			
			Alert alert = new Alert(AlertType.INFORMATION);
	        alert.setTitle("提示");
	        alert.setHeaderText("失敗");
	        alert.setContentText( e.getMessage());
	        alert.showAndWait();
		}
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
	private void loadAttendence() {
		
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
	 * 12 請假狀況
	 * 13 假別 (請假紀錄)
	 * 14 時數 (請假紀錄)
	 * 15 加總 (請假紀錄)
	 * 16 綜合工時
	 * 17 加班類別/時數
	 * 18 津貼類別/時數
	 * 19 刷卡紀錄
	 * 20 班別名稱
	 */
	private void createLog() {
			
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
	
	private static WritableCellFormat getLeaveExcelCellSetting() {
		// Cell的格式，如下
		WritableCellFormat cellFormat = new WritableCellFormat ();

		try {
			// 參閱 http://www.cnblogs.com/smilsy/articles/2126377.html
			cellFormat.setBackground(Colour.YELLOW);
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
}
