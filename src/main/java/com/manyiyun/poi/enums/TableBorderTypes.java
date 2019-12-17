package com.manyiyun.poi.enums;

public enum TableBorderTypes {
	NIL("nil"), 
	NONE("none"), 
	SINGLE("single"), 
	THICK("thick"), 
	DOUBLE("double"), 
	DOTTED("dotted"), 
	DASHED("dashed");
	private String type;
	TableBorderTypes(String type){
		this.type=type;
	}
}
