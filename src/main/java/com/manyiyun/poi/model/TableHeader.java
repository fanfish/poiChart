package com.manyiyun.poi.model;

public class TableHeader {
	//表头行数
	private int rows;
	//表头列数
	private int columns;
	//表头各列标题(单行表头)
	private String[] title;
	//表头各列标题(两行表头)
	private String[][] title2d;
	//表头背景色
	private String bgColor ="A7BFDE";
	//表头文本颜色
	private String fontColor="000000";
	//表头文本大小
	private int fontSize=10;
	//表头文本字体
	private String fontFamily="Courier";
	//表头文本是否粗体
	private boolean isBold;
	//是否是两行表头
	private boolean h2d;
	
	public int getRows() {
		return rows;
	}
	public void setRows(int rows) {
		this.rows = rows;
	}
	public int getColumns() {
		return columns;
	}
	public void setColumns(int columns) {
		this.columns = columns;
	}
	
	public TableHeader(String[] title) {
		this.title=title;
	}
	public TableHeader(String[][] title2d) {
		this.title2d=title2d;
		this.h2d=true;
		this.rows=title2d.length;
		this.columns=title2d[0].length;
	}
	
	public TableHeader(int rows,int columns) {
		this.rows=rows;
		this.columns=columns;
		this.h2d=true;
		this.title2d = new String[rows][columns];
		for(int i=0;i<rows;i++) {
			for(int j=0;j<columns;j++) {
				this.title2d[i][j]="";
			}
		}
	}
	
	public boolean isH2d() {
		return h2d;
	}

	public String getBgColor() {
		return bgColor;
	}

	public void setBgColor(String bgColor) {
		this.bgColor = bgColor;
	}

	public String getFontColor() {
		return fontColor;
	}

	public void setFontColor(String fontColor) {
		this.fontColor = fontColor;
	}

	public int getFontSize() {
		return fontSize;
	}

	public void setFontSize(int fontSize) {
		this.fontSize = fontSize;
	}

	public String getFontFamily() {
		return fontFamily;
	}

	public void setFontFamily(String fontFamily) {
		this.fontFamily = fontFamily;
	}

	public boolean isBold() {
		return isBold;
	}

	public void setBold(boolean isBold) {
		this.isBold = isBold;
	}

	
	public String[] getTitle() {
		return title;
	}

	public void setTitle(String[] title) {
		this.title = title;
	}

	
	public String[][] getTitle2d() {
		return title2d;
	}

	public void setTitle2d(String[][] title2d) {
		this.title2d = title2d;
	}

}
