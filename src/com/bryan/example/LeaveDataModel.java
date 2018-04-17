package com.bryan.example;

public class LeaveDataModel {
	int position;
	
	String category;
	String count;
	String startTime;
	String endTime;
	
	boolean isNoEndWorkingTime;
	String labelAttribute;
	
	public LeaveDataModel() {
		
	}
	
	public void setPosition(int position) {
		this.position = position;
	}
	
	public void setCategory(String category) {
		this.category = category;
	}
	
	public void setCount(String count) {
		this.count = count;
	}
	
	public void setStartTime(String startTime) {
		this.startTime = startTime;
	}
	
	public void setEndTime(String endTime) {
		this.endTime = endTime;
	}
	
	public void isNoEndWorkingTime(boolean status) {
		this.isNoEndWorkingTime = status;
	}
	
	public int getPosition() {
		return position;
	}
	
	public String getCategory() {
		return category;
	}
	
	public String getCount() {
		return count;
	}
	
	public boolean isNoEndWorkingTime() {
		return isNoEndWorkingTime;
	}
	
	public String getStartTime() {
		return startTime;
	}
	
	public String getEndTime() {
		return endTime;
	}
	
}
