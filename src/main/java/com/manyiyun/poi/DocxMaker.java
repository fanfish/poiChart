package com.manyiyun.poi;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;


import com.manyiyun.poi.model.Chart;
import com.manyiyun.poi.model.TableProperty;
import com.manyiyun.poi.model.TableColumnData;


public interface DocxMaker {
	public void write(String outDocxPath) throws IOException;
	public void write(XWPFDocument doc, String outDocxPath) throws Exception;
	public void createRegion(ParagraphAlignment align, int firstLineIndent, String text, boolean bold, String rgbStr,String fontFamily, int fontSize);
	public void createPicture(File imgFile, int width, int height)throws Exception;
	public void createChart(Chart chart)throws Exception;
	public XWPFTable createTable(TableProperty<?> tbl,List<TableColumnData> columnsOrder) throws Exception;
	public XWPFTable createTable(TableProperty<?> tbl) throws Exception;
	public void addTitle(int level, String titleText,boolean isBreak);
	public void createToc();
	public void createHeader(String headerText) throws Exception;;
	public void createFooter() throws Exception;
	
}
