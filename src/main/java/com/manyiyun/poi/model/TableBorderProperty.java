package com.manyiyun.poi.model;

import java.math.BigInteger;


public class TableBorderProperty{

	private String val ="single";
	private BigInteger sz =new BigInteger("1");
	private String color="000000";
	
	public TableBorderProperty(String val ,BigInteger sz,String color) {
		this.val=val;
		this.sz=sz;
		this.color=color;
	}
	public TableBorderProperty(String val) {
		this.val=val;
	}
	public TableBorderProperty() {

	}

	public String getVal() {
		return val;
	}
	public void setVal(String val) {
		this.val = val;
	}
	public BigInteger getSz() {
		return sz;
	}
	public void setSz(BigInteger sz) {
		this.sz = sz;
	}
	public String getColor() {
		return color;
	}
	public void setColor(String color) {
		this.color = color;
	}

}
