package com.manyiyun.poi.model;

import java.math.BigInteger;
import java.util.List;

import com.manyiyun.poi.model.TableBorderProperty;
import com.manyiyun.poi.model.TableDataRowProperty;
import com.manyiyun.poi.model.TableHeader;


public class TableProperty<T> {
	//表格表头对象
	private TableHeader header;
	//表格数据列表
	private List<T>data;
	//表格宽度，默认8000
	private BigInteger width = new BigInteger("8000");
	//水平边框
	private TableBorderProperty hBorder;
	//竖向边框
	private TableBorderProperty vBorder;
	//左边框
	private TableBorderProperty lBorder;
	//右边框
	private TableBorderProperty rBorder;
	//上边框
	private TableBorderProperty tBorder;
	//下边框
	private TableBorderProperty bBorder;
	//表格行属性对象(主要是为控制奇、偶数行的背景色)
	private TableDataRowProperty rowProperty;
	

	public TableProperty(TableHeader header,List<T>data) {
		this.header=header;
		this.data=data;
		this.hBorder = new TableBorderProperty("none");
		this.vBorder = new TableBorderProperty("none");
		this.lBorder = new TableBorderProperty("none");
		this.rBorder = new TableBorderProperty("none");
		this.tBorder = new TableBorderProperty();
		this.bBorder = new TableBorderProperty();
		this.rowProperty = new TableDataRowProperty();
	}		


	public BigInteger getWidth() {
		return width;
	}
	public void setWidth(BigInteger width) {
		this.width = width;
	}
	public TableHeader getHeader() {
		return header;
	}
	public void setHeader(TableHeader header) {
		this.header = header;
	}
	public List<T> getData() {
		return data;
	}
	public void setData(List<T> data) {
		this.data = data;
	}
	public TableBorderProperty gethBorder() {
		return hBorder;
	}

	public void sethBorder(TableBorderProperty hBorder) {
		this.hBorder = hBorder;
	}

	public TableBorderProperty getvBorder() {
		return vBorder;
	}

	public void setvBorder(TableBorderProperty vBorder) {
		this.vBorder = vBorder;
	}

	public TableBorderProperty getlBorder() {
		return lBorder;
	}

	public void setlBorder(TableBorderProperty lBorder) {
		this.lBorder = lBorder;
	}

	public TableBorderProperty getrBorder() {
		return rBorder;
	}

	public void setrBorder(TableBorderProperty rBorder) {
		this.rBorder = rBorder;
	}

	public TableBorderProperty gettBorder() {
		return tBorder;
	}

	public void settBorder(TableBorderProperty tBorder) {
		this.tBorder = tBorder;
	}

	public TableBorderProperty getbBorder() {
		return bBorder;
	}

	public void setbBorder(TableBorderProperty bBorder) {
		this.bBorder = bBorder;
	}

	
	public TableDataRowProperty getRowProperty() {
		return rowProperty;
	}

	public void setRowProperty(TableDataRowProperty rowProperty) {
		this.rowProperty = rowProperty;
	}
}
