package com.bryan.example;

public class LeaveDataModel {
	int position;
	String category;
	String count;
	float sum;
	boolean isNoEndWorkingTime;
	
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
	
	public void setLeaveSum(float sum) {
		this.sum = sum;
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
	
	public float getLeaveSum() {
		return sum;
	}
	
	public boolean isNoEndWorkingTime() {
		return isNoEndWorkingTime;
	}
	
}
