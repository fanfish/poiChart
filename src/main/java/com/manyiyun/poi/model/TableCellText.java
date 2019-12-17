package com.manyiyun.poi.model;

import com.manyiyun.poi.enums.TCTextAlignment;

public class TableCellText {
	public TableCellText(String text) {
		this.text=text;
	}
	
	public String getFontFamily() {
		return fontFamily;
	}
	public void setFontFamily(String fontFamily) {
		this.fontFamily = fontFamily;
	}
	public int getFontSize() {
		return fontSize;
	}
	public void setFontSize(int fontSize) {
		this.fontSize = fontSize;
	}
	public String getFontColor() {
		return fontColor;
	}
	public void setFontColor(String fontColor) {
		this.fontColor = fontColor;
	}
	public boolean isFontBold() {
		return fontBold;
	}
	public void setFontBold(boolean fontBold) {
		this.fontBold = fontBold;
	}
	public TCTextAlignment getFontAlign() {
		return fontAlign;
	}
	public void setFontAlign(TCTextAlignment fontAlign) {
		this.fontAlign = fontAlign;
	}
	public String getText() {
		return text;
	}
	public void setText(String text) {
		this.text = text;
	}
	private boolean fontBold;
	private String text;
	private TCTextAlignment fontAlign=TCTextAlignment.CENTER;	
	private String fontColor="000000";
	private int fontSize=10;
	private String fontFamily="Courier";

}
