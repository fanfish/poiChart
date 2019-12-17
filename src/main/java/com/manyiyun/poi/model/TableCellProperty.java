package com.manyiyun.poi.model;

import java.math.BigInteger;

import com.manyiyun.poi.enums.TCTextAlignment;
import com.manyiyun.poi.enums.TableCellBorderTypes;

public class TableCellProperty {
	
	//单元格背景色
	private String bgColor="A7BFDE";
	//单元格行号
	private int row;
	//单元格列号
	private int column;
	//private TableCellBorder border;
	//单元格是否显示外边框
	private TableCellBorderTypes have;
	//单元格文本对象(文本字体、大小、颜色等)
	private TableCellText text;
	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public int getColumn() {
		return column;
	}

	public void setColumn(int column) {
		this.column = column;
	}
	
	public TableCellProperty(int row,int column,TableCellBorderTypes border,TableCellText text) {
		this.have=border;
		this.text=text;
		this.row=row;
		this.column=column;
	}
	
	/*
	 * public TableCellBorder getBorder() { return border; }
	 * 
	 * public void setBorder(TableCellBorder border) { this.border = border; }
	 */

	public TableCellText getText() {
		return text;
	}

	public void setText(TableCellText text) {
		this.text = text;
	}


    public String getBgColor() {
		return bgColor;
	}

	public void setBgColor(String bgColor) {
		this.bgColor = bgColor;
	}
	
	public TableCellBorderTypes getHave() {
		return have;
	}

	public void setHave(TableCellBorderTypes have) {
		this.have = have;
	}

    
}
