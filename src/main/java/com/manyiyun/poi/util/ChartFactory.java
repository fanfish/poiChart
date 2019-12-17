package com.manyiyun.poi.util;

import java.util.List;

import com.manyiyun.poi.enums.ChartTypes;
import com.manyiyun.poi.model.BarChart;
import com.manyiyun.poi.model.BarLineChart;
import com.manyiyun.poi.model.Chart;
import com.manyiyun.poi.model.LineChart;
import com.manyiyun.poi.model.PieChart;
import com.manyiyun.poi.model.RadarChart;
import com.manyiyun.poi.model.VBarChart;

public class ChartFactory {

	  public static Chart getChart(ChartTypes type,String chartTitle, String[] series, String[] categories,List<Double[]> values,int width,int height)
	  {
		Chart instance = null;
		if(type==ChartTypes.VBAR)
		  return instance =new VBarChart(chartTitle, series, categories,values, width, height);
		else if ( type==ChartTypes.BAR)
	      return instance =new BarChart(chartTitle, series, categories,values, width, height);
	    else if ( type==ChartTypes.PIE )
	      return instance =new PieChart(chartTitle, series, categories,values, width, height);
	    else if ( type==ChartTypes.LINE)
	      return instance =new LineChart(chartTitle, series, categories,values, width, height);
	    else if ( type==ChartTypes.RADAR)
		      return instance =new RadarChart(chartTitle, series, categories,values, width, height);
	    else if ( type==ChartTypes.BAR_LINE)
		      return instance =new BarLineChart(chartTitle, series, categories,values, width, height);
	    return instance;
	  }
}
