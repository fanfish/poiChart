package com.manyiyun.poi.data;

import java.io.Serializable;
import java.util.Date;

import com.manyiyun.poi.util.DocxTblColumn;




public class Hospital implements Serializable{
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getAddress() {
		return address;
	}
	public void setAddress(String address) {
		this.address = address;
	}
	public int getBedNum() {
		return bedNum;
	}
	public void setBedNum(int bedNum) {
		this.bedNum = bedNum;
	}
	public boolean isSj() {
		return isSj;
	}
	public void setSj(boolean isSj) {
		this.isSj = isSj;
	}
	public Date getEnd() {
		return end;
	}
	public void setEnd(Date end) {
		this.end = end;
	}
	public double getArea() {
		return area;
	}
	public void setArea(double area) {
		this.area = area;
	}
	//医院名称
	@DocxTblColumn(columnName="name",order=0)
	private String name;
	//医院地址
	@DocxTblColumn(columnName="address",order=1)
	private String address;
	//床位数	
	private int bedNum;
	//是否三甲
	@DocxTblColumn(columnName="isSj",order=3)
	private boolean isSj;
	//截止日期
	@DocxTblColumn(columnName="end",order=4,format="yyyy-MM-dd")
	private Date end;
	//占地面积
	@DocxTblColumn(columnName="area",order=2,format="#.#")
	private double area;
	
	
}
