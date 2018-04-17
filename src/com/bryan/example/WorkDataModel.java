package com.bryan.example;

import java.util.List;

import jxl.write.Label;

public class WorkDataModel {
	private int position;

    private String attendanceGroupId;
    private String attendanceGroupName;
    private String attendanceId;
    private String attendanceName;
    private String attendanceDate;
    private String startWorkingTime;
    private String endWorkingTime;
    //
    private String workingMinute;		// 總工時
    private String workingHours;		// 工時	(時)
    private String workingMin;		// 工時	(分)
    //
    private String lateTotalTime;		// 遲到
    private String leaveEarlyTotalTime;		// 早退
    //
    private List<LeaveDataModel> leaveData;		// 請假紀錄
    //
    private String complexWorkingTime;
    private String overtimeCategory;
    private String allowanceCategory;
    private String bookingRecord;
    private String workingItem;

    private String labelAttribute;

    public int getPosition() {
        return position;
    }

    public String getAttendanceDate() {
        return attendanceDate;
    }

    public String getAttendanceGroupId() {
        return attendanceGroupId;
    }

    public String getAttendanceGroupName() {
        return attendanceGroupName;
    }

    public String getAttendanceId() {
        return attendanceId;
    }

    public String getAttendanceName() {
        return attendanceName;
    }

    public List<LeaveDataModel> getLeaveData() {
        return leaveData;
    }

    public String getComplexWorkingTime() {
        return complexWorkingTime;
    }

    public String getAllowanceCategory() {
        return allowanceCategory;
    }

    public String getEndWorkingTime() {
        return endWorkingTime;
    }

    public String getLateTotalTime() {
        return lateTotalTime;
    }

    public String getLeaveEarlyTotalTime() {
        return leaveEarlyTotalTime;
    }

    public String getOvertimeCategory() {
        return overtimeCategory;
    }

    public String getBookingRecord() {
        return bookingRecord;
    }

    public String getLabelAttribute() {
        return labelAttribute;
    }

    public String getStartWorkingTime() {
        return startWorkingTime;
    }

    public String getWorkingHours() {
        return workingHours;
    }

    public String getWorkingItem() {
        return workingItem;
    }

    public String getWorkingMin() {
        return workingMin;
    }

    public String getWorkingMinute() {
        return workingMinute;
    }

    public void setAllowanceCategory(String allowanceCategory) {
        this.allowanceCategory = allowanceCategory;
    }

    public void setAttendanceDate(String attendanceDate) {
        this.attendanceDate = attendanceDate;
    }

    public void setAttendanceGroupId(String attendanceGroupId) {
        this.attendanceGroupId = attendanceGroupId;
    }

    public void setAttendanceGroupName(String attendanceGroupName) {
        this.attendanceGroupName = attendanceGroupName;
    }

    public void setAttendanceId(String attendanceId) {
        this.attendanceId = attendanceId;
    }

    public void setAttendanceName(String attendanceName) {
        this.attendanceName = attendanceName;
    }

    public void setBookingRecord(String bookingRecord) {
        this.bookingRecord = bookingRecord;
    }

    public void setComplexWorkingTime(String complexWorkingTime) {
        this.complexWorkingTime = complexWorkingTime;
    }

    public void setEndWorkingTime(String endWorkingTime) {
        this.endWorkingTime = endWorkingTime;
    }

    public void setLabelAttribute(String labelAttribute) {
        this.labelAttribute = labelAttribute;
    }

    public void setLateTotalTime(String lateTotalTime) {
        this.lateTotalTime = lateTotalTime;
    }

    public void setLeaveData(List<LeaveDataModel> leaveData) {
        this.leaveData = leaveData;
    }

    public void setLeaveEarlyTotalTime(String leaveEarlyTotalTime) {
        this.leaveEarlyTotalTime = leaveEarlyTotalTime;
    }

    public void setOvertimeCategory(String overtimeCategory) {
        this.overtimeCategory = overtimeCategory;
    }

    public void setPosition(int position) {
        this.position = position;
    }

    public void setStartWorkingTime(String startWorkingTime) {
        this.startWorkingTime = startWorkingTime;
    }

    public void setWorkingHours(String workingHours) {
        this.workingHours = workingHours;
    }

    public void setWorkingItem(String workingItem) {
        this.workingItem = workingItem;
    }

    public void setWorkingMin(String workingMin) {
        this.workingMin = workingMin;
    }

    public void setWorkingMinute(String workingMinute) {
        this.workingMinute = workingMinute;
    }
	
}
