package com.manyiyun.poi.model;

import java.util.List;

import org.apache.poi.xddf.usermodel.chart.XDDFChart;

import com.manyiyun.poi.enums.ChartTypes;


public abstract class Chart {	
	public String getChartTitle() {
		return chartTitle;
	}
	public void setChartTitle(String chartTitle) {
		this.chartTitle = chartTitle;
	}
	public String[] getSeries() {
		return series;
	}
	public void setSeries(String[] series) {
		this.series = series;
	}
	public String[] getCategories() {
		return categories;
	}
	public void setCategories(String[] categories) {
		this.categories = categories;
	}
	public List<Double[]> getValues() {
		return values;
	}
	public void setValues(List<Double[]> values) {
		this.values = values;
	}
	public int getWidth() {
		return width;
	}
	public void setWidth(int width) {
		this.width = width;
	}
	public int getHeight() {
		return height;
	}
	public void setHeight(int height) {
		this.height = height;
	}
	protected String chartTitle;
	protected String[] series;
	protected String[] categories;
	protected List<Double[]> values;
	protected int width;
	protected int height;
	private static final int WIDTH = XDDFChart.DEFAULT_WIDTH;
	private static final int HEIGHT = XDDFChart.DEFAULT_HEIGHT;
	
	public abstract ChartTypes type();
	
	public  Chart(String chartTitle, String[] series, String[] categories,List<Double[]> values,int width,int height) {
		this.chartTitle=chartTitle;
		this.series=series;
		this.categories=categories;
		this.values = values;
		this.width = width*WIDTH;
		this.height = height*HEIGHT;
	}
	
}
