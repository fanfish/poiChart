package com.manyiyun.poi.model;

import java.util.List;

import com.manyiyun.poi.enums.ChartTypes;

public class LineChart extends Chart{

	public LineChart(String chartTitle, String[] series, String[] categories, List<Double[]> values,int width,int height) {
		super(chartTitle, series, categories, values, width, height);
		// TODO Auto-generated constructor stub
	}

	@Override
	public ChartTypes type() {
		// TODO Auto-generated method stub
		return ChartTypes.LINE;
	}

}
