package com.manyiyun.poi.model;

import com.manyiyun.poi.enums.TCTextAlignment;

public class TableDataRowProperty {
	private String fontColor="000000";
	private int fontSize=10;
	private String fontFamily="Courier";
	private boolean isBold;
	private String evenBgColor="D3DFEE";
	private TCTextAlignment fontAlign=TCTextAlignment.CENTER;
	private String oddBgColor="ffffff";
	
	
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

	public TCTextAlignment getFontAlign() {
		return fontAlign;
	}
	public void setFontAlign(TCTextAlignment fontAlign) {
		this.fontAlign = fontAlign;
	}

	public String getEvenBgColor() {
		return evenBgColor;
	}
	public void setEvenBgColor(String evenBgColor) {
		this.evenBgColor = evenBgColor;
	}
	public String getOddBgColor() {
		return oddBgColor;
	}
	public void setOddBgColor(String oddBgColor) {
		this.oddBgColor = oddBgColor;
	}

}
