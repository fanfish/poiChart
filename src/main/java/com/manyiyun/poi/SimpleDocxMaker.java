package com.manyiyun.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.net.URL;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.AxisCrossBetween;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.AxisTickMark;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.RadarStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import com.manyiyun.poi.exception.SeriesValuesException;
import com.manyiyun.poi.exception.UnequalException;
import com.manyiyun.poi.model.Chart;
import com.manyiyun.poi.model.TableProperty;

import com.manyiyun.poi.model.TableCellProperty;
import com.manyiyun.poi.model.TableCellText;
import com.manyiyun.poi.model.TableColumnData;
import com.manyiyun.poi.model.TableDataRowProperty;

import com.manyiyun.poi.util.CustomXWPFDocument;
import com.manyiyun.poi.util.DocxTblColumn;
import com.manyiyun.poi.enums.ChartTypes;
import com.manyiyun.poi.enums.TableCellBorderTypes;

public class SimpleDocxMaker implements DocxMaker {

	private CustomXWPFDocument wxpf = new CustomXWPFDocument();
	
	/** 
	* @Description .docx文件输出 
	* @param outDocxPath 文件输出路径 
	* @return
	* @throws 
	*/
	public void write(String outDocxPath) throws IOException {
		// save the result
		try (OutputStream out = new FileOutputStream(outDocxPath)) {
			wxpf.write(out);
		}
	}

	/** 
	* @Description .docx文件输出 
	* @param doc XWPFDocument对象引用 
	* @param outDocxPath 文件输出路径 
	* @return
	* @throws 
	*/
	public void write(XWPFDocument doc, String outDocxPath) throws IOException, UnequalException {
		// save the result
		if (!doc.equals(this.wxpf))
			throw new UnequalException("parameter object doc must obtain from the method of getWxpf()");
		try (OutputStream out = new FileOutputStream(outDocxPath)) {
			doc.write(out);
		}
	}

	/** 
	* @Description 创建文本
	* @param align 对齐方式
	* @param firstLineIndent 缩进大小(0代表无缩进)
	* @param text 文本内容
	* @param bold 是否粗体显示(true,是;false,否)
	* @param rgbStr 文本颜色RGB字符串
	* @return
	* @throws 
	*/
	public void createRegion(ParagraphAlignment align, int firstLineIndent, String text, boolean bold, String rgbStr,
			String fontFamily, int fontSize) {
		XWPFParagraph para = wxpf.createParagraph();
		para.setAlignment(align);
		// para.setIndentationFirstLine(480);//首行缩进24磅
		para.setIndentationFirstLine(firstLineIndent);
		XWPFRun run = para.createRun();
		run.setText(text);
		run.setColor(rgbStr);
		run.setBold(bold);
		run.setFontFamily(fontFamily);
		run.setFontSize(fontSize);
	}

	/** 
	* @Description 创建图片
	* @param pics 图片文件名列表
	* @param width 图片宽度
	* @param height 图片高度
	* @return
	* @throws 
	*/
	public void createPicture(File imgFile, int width, int height) throws InvalidFormatException, IOException {
		XWPFParagraph p = wxpf.createParagraph();
		p.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun r = p.createRun();			
		int format=XWPFDocument.PICTURE_TYPE_PNG;
		System.out.println(imgFile.getName());
		if (imgFile.getName().endsWith(".emf")) {
			format = XWPFDocument.PICTURE_TYPE_EMF;
		} else if (imgFile.getName().endsWith(".wmf")) {
			format = XWPFDocument.PICTURE_TYPE_WMF;
		} else if (imgFile.getName().endsWith(".pict")) {
			format = XWPFDocument.PICTURE_TYPE_PICT;
		} else if (imgFile.getName().endsWith(".jpeg") || imgFile.getName().endsWith(".jpg")) {
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		} else if (imgFile.getName().endsWith(".png")) {
				format = XWPFDocument.PICTURE_TYPE_PNG;
		} else if (imgFile.getName().endsWith(".dib")) {
				format = XWPFDocument.PICTURE_TYPE_DIB;
		} else if (imgFile.getName().endsWith(".gif")) {
				format = XWPFDocument.PICTURE_TYPE_GIF;
		} else if (imgFile.getName().endsWith(".tiff")) {
				format = XWPFDocument.PICTURE_TYPE_TIFF;
		} else if (imgFile.getName().endsWith(".eps")) {
				format = XWPFDocument.PICTURE_TYPE_EPS;
		} else if (imgFile.getName().endsWith(".bmp")) {
				format = XWPFDocument.PICTURE_TYPE_BMP;
		} else if (imgFile.getName().endsWith(".wpg")) {
				format = XWPFDocument.PICTURE_TYPE_WPG;
		} else {
			System.err.println("Unsupported picture: " + imgFile
						+ ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");				
		}			
		r.addBreak();			
		try (FileInputStream is = new FileInputStream(imgFile)) {
			r.addPicture(is, format, imgFile.getName(), Units.toEMU(width), Units.toEMU(width)); 
		}		
	}

	/** 
	* @Description 创建表格
	* @param tbl TableProperty对象引用
	* @return XWPFTable对象实例
	* @throws 
	*/
	public XWPFTable createTable(TableProperty<?> tbl) throws Exception {
		if (tbl == null)
			throw new NullPointerException("the parameter is null!");		
		Map<Integer, TableColumnData> map = new TreeMap<Integer, TableColumnData>();
		Object po = tbl.getData().get(0);
		for (Field f : po.getClass().getDeclaredFields()) {
			DocxTblColumn dtc = f.getAnnotation(DocxTblColumn.class);
			if (dtc != null) {
				TableColumnData tcd = new TableColumnData();
				tcd.setColumnName(dtc.columnName());
				if (f.getType().equals(float.class) || f.getType().equals(Float.TYPE)
						|| f.getType().equals(double.class) || f.getType().equals(Double.TYPE)) {
					DecimalFormat df = new DecimalFormat(dtc.format());
					tcd.setFormat(df);
				} else if (f.getType().equals(Date.class)) {
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					tcd.setFormat(sdf);
				} else {
					tcd.setFormat(null);
				}
				map.put(dtc.order(), tcd);
			}
		}
		List<TableColumnData> columnsOrder = new ArrayList<TableColumnData>(map.values());
		return createTable(tbl, columnsOrder);
	}

	/** 
	* @Description 创建表格
	* @param tbl TableProperty对象引用
	* @param columnsOrder TableColumnData对象列表
	* @return XWPFTable对象实例
	* @throws 
	*/
	public XWPFTable createTable(TableProperty<?> tbl, List<TableColumnData> columnsOrder) throws Exception {
		if (tbl == null || columnsOrder==null)
			throw new NullPointerException("the parameters  is null!");
		XWPFTable table = null;
		if (tbl.getHeader().isH2d()) {
			int nRows = tbl.getData().size() + tbl.getHeader().getRows();
			int nCols = tbl.getHeader().getColumns();
			table = wxpf.createTable(nRows, nCols);
		} else {
			int nRows = tbl.getData().size() + 1;
			int nCols = tbl.getHeader().getTitle().length;
			table = wxpf.createTable(nRows, nCols);
		}
		table.setTableAlignment(TableRowAlign.CENTER);
		// Set the table style. If the style is not defined, the table style
		// will become "Normal".
		CTTblPr tblPr = table.getCTTbl().getTblPr();

		tblPr.getTblW().setType(STTblWidth.DXA);
		tblPr.getTblW().setW(tbl.getWidth());
		// 设置表格边框
		setTableBorders(table, tbl);
		// CTString styleStr = tblPr.addNewTblStyle();
		// styleStr.setVal("StyledTable");

		// Get a list of the rows in the table
		List<XWPFTableRow> rows = table.getRows();
		int rowCt = 0;
		int colCt = 0;
		for (XWPFTableRow row : rows) {
			// get table row properties (trPr)
			CTTrPr trPr = row.getCtRow().addNewTrPr();
			// set row height; units = twentieth of a point, 360 = 0.25"
			CTHeight ht = trPr.addNewTrHeight();
			ht.setVal(BigInteger.valueOf(360));
			// get the cells in this row
			List<XWPFTableCell> cells = row.getTableCells();
			// add content to each cell
			for (XWPFTableCell cell : cells) {
				// get a table cell properties element (tcPr)
				CTTcPr tcpr = cell.getCTTc().addNewTcPr();
				// set vertical alignment to "center"
				CTVerticalJc va = tcpr.addNewVAlign();
				va.setVal(STVerticalJc.CENTER);

				// create cell color element
				CTShd ctshd = tcpr.addNewShd();
				ctshd.setColor("auto");
				ctshd.setVal(STShd.CLEAR);
				int hRows = 0;
				String hBgColor = "ffffff";

				if (tbl.getHeader().isH2d())
					hRows = tbl.getHeader().getRows() - 1;
				hBgColor = tbl.getHeader().getBgColor();
				// get 1st paragraph in cell's paragraph list
				XWPFParagraph para = cell.getParagraphs().get(0);
				// create a run to contain the content
				XWPFRun rh = para.createRun();
				if (rowCt <= hRows) {
					// header row
					ctshd.setFill(hBgColor);
					para.setAlignment(ParagraphAlignment.CENTER);
					rh.setFontSize(tbl.getHeader().getFontSize());
					rh.setFontFamily(tbl.getHeader().getFontFamily());
					rh.setBold(tbl.getHeader().isBold());
					if (tbl.getHeader().isH2d())
						rh.setText(tbl.getHeader().getTitle2d()[rowCt][colCt]);
					else
						rh.setText(tbl.getHeader().getTitle()[colCt]);
				} else {
					// other rows
					TableDataRowProperty rowProperty = tbl.getRowProperty();
					if (rowCt % 2 == 0) {
						// even row
						ctshd.setFill(rowProperty.getEvenBgColor());
					} else {
						// odd row
						ctshd.setFill(rowProperty.getOddBgColor());
					}
					para.setAlignment(ParagraphAlignment.CENTER);
					Object po = tbl.getData().get(rowCt - tbl.getHeader().getRows());
					String fieldName = columnsOrder.get(colCt).getColumnName();
					java.text.Format df = columnsOrder.get(colCt).getFormat();
					// To get private fields use
					Field field = po.getClass().getDeclaredField(fieldName);
					field.setAccessible(true);
					String value = "";
					if (field.getType().equals(boolean.class) || field.getType().equals(Boolean.TYPE)) {
						boolean b = field.getBoolean(po);
						value = b ? "是" : "否";
					} else if (field.getType().equals(String.class)) {
						value = (field.get(po) == null) ? value : String.valueOf(field.get(po));
					} else if (field.getType().equals(Date.class)) {
						if (df != null)
							value = df.format(field.get(po));
					} else if (field.getType().equals(float.class) || field.getType().equals(Float.TYPE)
							|| field.getType().equals(double.class) || field.getType().equals(Double.TYPE)) {
						if (df != null)
							value = df.format(field.get(po));
					} else {
						value = String.valueOf(field.get(po));
					}
					rh.setText(value);
				}
				colCt++;
			} // for cell
			colCt = 0;
			rowCt++;
		} // for row
		return table;
	}

	private void setTableBorders(XWPFTable table, TableProperty<?> tbl) {
		// 添加边框
		CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
		CTBorder hBorder = borders.addNewInsideH();
		hBorder.setVal(STBorder.Enum.forString(tbl.gethBorder().getVal()));
		hBorder.setSz(tbl.gethBorder().getSz());
		hBorder.setColor(tbl.gethBorder().getColor());

		CTBorder vBorder = borders.addNewInsideV();
		vBorder.setVal(STBorder.Enum.forString(tbl.getvBorder().getVal()));
		vBorder.setSz(tbl.getvBorder().getSz());
		vBorder.setColor(tbl.getvBorder().getColor());

		CTBorder lBorder = borders.addNewLeft();
		lBorder.setVal(STBorder.Enum.forString(tbl.getlBorder().getVal()));
		lBorder.setSz(tbl.getlBorder().getSz());
		lBorder.setColor(tbl.getlBorder().getColor());

		CTBorder rBorder = borders.addNewRight();
		rBorder.setVal(STBorder.Enum.forString(tbl.getrBorder().getVal()));
		rBorder.setSz(tbl.getrBorder().getSz());
		rBorder.setColor(tbl.getrBorder().getColor());

		CTBorder tBorder = borders.addNewTop();
		tBorder.setVal(STBorder.Enum.forString(tbl.gettBorder().getVal()));
		tBorder.setSz(tbl.gettBorder().getSz());
		tBorder.setColor(tbl.gettBorder().getColor());

		CTBorder bBorder = borders.addNewBottom();
		bBorder.setVal(STBorder.Enum.forString(tbl.getbBorder().getVal()));
		bBorder.setSz(tbl.getbBorder().getSz());
		bBorder.setColor(tbl.getbBorder().getColor());
	}

	/** 
	* @Description 创建目录标题
	* @param level 标题级别
	* @param titleText 标题文本
	* @param isBreak 是否断行
	* @return 
	* @throws 
	*/
	public void addTitle(int level, String titleText, boolean isBreak) {
		XWPFParagraph paragraph = wxpf.createParagraph();
		XWPFRun run = paragraph.createRun();
		run = paragraph.createRun();
		if (isBreak)
			run.addBreak(BreakType.PAGE);
		run.setText(titleText);
		run.setBold(true);
		switch (level) {
		case 1:		
			run.setFontFamily("Calibri (西文正文)");
			run.setFontSize(22);
			break;
		case 2:			
			run.setFontFamily("Cambria");
			run.setFontSize(16);
			break;
		case 3:			
			run.setFontFamily("Calibri (西文正文)");
			run.setFontSize(16);
			break;
		case 4:			
			run.setFontFamily("Cambria (西文标题)");
			run.setFontSize(14);
			break;
		case 5:			
			run.setFontFamily("Calibri (西文正文)");
			run.setFontSize(14);
			break;
		}

		StringBuffer sb = new StringBuffer("heading ");
		sb.append(level);
		paragraph.setStyle(sb.toString());

	}
	
	/** 
	* @Description 创建目录大纲(默认有5级下级标题)
	* @return 
	* @throws 
	*/
	public void createToc() {
		wxpf.createTOC();
		addCustomHeadingStyle(wxpf, "heading 1", 1);
		addCustomHeadingStyle(wxpf, "heading 2", 2);
		addCustomHeadingStyle(wxpf, "heading 3", 3);
		addCustomHeadingStyle(wxpf, "heading 4", 4);
		addCustomHeadingStyle(wxpf, "heading 5", 5);
		// the body content
		XWPFParagraph paragraph = wxpf.createParagraph();
		CTP ctP = paragraph.getCTP();
		CTSimpleField toc = ctP.addNewFldSimple();
		toc.setInstr("TOC \\h");
		toc.setDirty(STOnOff.TRUE);	
	}

	/** 
	* @Description 创建图表
	* @param chart Chart对象引用
	* @return 
	* @throws IOException,SeriesValuesException,InvalidFormatException
	*/
	public void createChart(Chart chart) throws IOException, SeriesValuesException, InvalidFormatException {
		int width = chart.getWidth();
		int height = chart.getHeight();
		if(chart.type() == ChartTypes.VBAR) {
			XWPFChart xc = wxpf.createChart(width, height);
			setVBarData(xc, chart.getChartTitle(), chart.getSeries(), chart.getCategories(), chart.getValues());
		}else if (chart.type() == ChartTypes.BAR) {
			XWPFChart xc = wxpf.createChart(width, height);
			setBarData(xc, chart.getChartTitle(), chart.getSeries(), chart.getCategories(), chart.getValues());
		} else if (chart.type() == ChartTypes.PIE) {
			List<Double[]> values = chart.getValues();
			if (values != null && values.size() != 1)
				throw new SeriesValuesException("pie chart values size must be 1");
			XWPFChart xc = wxpf.createChart(width, height);
			setPieData(xc, chart.getChartTitle(), chart.getSeries(), chart.getCategories(), chart.getValues());
		} else if (chart.type() == ChartTypes.LINE) {
			XWPFChart xc = wxpf.createChart(width, height);
			setLineData(xc, chart.getChartTitle(), chart.getSeries(), chart.getCategories(), chart.getValues());
		} else if (chart.type() == ChartTypes.RADAR) {
			XWPFChart xc = wxpf.createChart(width, height);
			setRadarData(xc, chart.getChartTitle(), chart.getSeries(), chart.getCategories(), chart.getValues());
		} else if (chart.type() == ChartTypes.BAR_LINE) {
			List<Double[]> values = chart.getValues();
			if (values != null && values.size() != 2)
				throw new SeriesValuesException("barLine chart values size must be 2");
			XWPFChart xc = wxpf.createChart(width, height);
			setBarLineData(xc, chart.getChartTitle(), chart.getSeries(), chart.getCategories(), chart.getValues());
		}
	}
	
    private void setVBarData(XWPFChart chart, String chartTitle, String[] series, String[] categories, List<Double[]> list) {
        // Use a category axis for the bottom axis.
        XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);      
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);        
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setMajorTickMark(AxisTickMark.OUT);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

        final int numOfPoints = categories.length;
		XDDFBarChartData bar = (XDDFBarChartData) chart.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.BAR,
				bottomAxis, leftAxis);
		final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
		final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);		
		
		for (int i = 0; i < list.size(); i++) {
			final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, i + 1, i + 1));
			final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(list.get(i),
					valuesDataRange, i);
			valuesData.setFormatCode("General");
			XDDFChartData.Series s = bar.addSeries(categoriesData, valuesData);
			s.setTitle(series[i], chart.setSheetTitle(series[i], i));
		}        
        bar.setBarGrouping(BarGrouping.CLUSTERED);
        bar.setVaryColors(true);
        bar.setBarDirection(BarDirection.COL);
        chart.plot(bar);
		if (list.size() > 1) {
			XDDFChartLegend legend = chart.getOrAddLegend();
			legend.setPosition(LegendPosition.BOTTOM);
			legend.setOverlay(false);
		}
        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);       
    }


	private void setBarData(XWPFChart chart, String chartTitle, String[] series, String[] categories,
			List<Double[]> list) {
		XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);		
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);		
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
		XDDFBarChartData bar = (XDDFBarChartData) chart.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.BAR,
				bottomAxis, leftAxis);
		final int numOfPoints = categories.length;
		final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
		final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);

		for (int i = 0; i < list.size(); i++) {
			final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, i + 1, i + 1));
			final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(list.get(i),
					valuesDataRange, i);
			XDDFChartData.Series s = bar.addSeries(categoriesData, valuesData);
			s.setTitle(series[i], chart.setSheetTitle(series[i], i));
		}
		chart.plot(bar);
		if (list.size() > 1) {
			XDDFChartLegend legend = chart.getOrAddLegend();
			legend.setPosition(LegendPosition.BOTTOM);
			legend.setOverlay(false);
		}
		chart.setTitleText(chartTitle); // https://stackoverflow.com/questions/30532612
		chart.setTitleOverlay(false);
	}

	private void setBarLineData(XWPFChart chart, String chartTitle, String[] series, String[] categories,
			List<Double[]> list) {
		// create data sources
		int numOfPoints = categories.length;
		String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
		String valuesDataRange1 = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
		String valuesDataRange2 = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
		XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
		XDDFNumericalDataSource<? extends Number> valuesData1 = XDDFDataSourcesFactory.fromArray(list.get(0),
				valuesDataRange1, 1);
		XDDFNumericalDataSource<? extends Number> valuesData2 = XDDFDataSourcesFactory.fromArray(list.get(1),
				valuesDataRange2, 2);		
		// first bar chart
		XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
		leftAxis.setMinimum(0);
		XDDFBarChartData bar = (XDDFBarChartData) chart.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.BAR,
				bottomAxis, leftAxis);
		XDDFChartData.Series s1 = bar.addSeries(categoriesData, valuesData1);
		s1.setTitle(series[0], chart.setSheetTitle(series[0], 0));
		bar.setVaryColors(true);
		bar.setBarDirection(BarDirection.COL);
		chart.plot(bar);		
		// second line chart
		bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setVisible(false);
		XDDFValueAxis rightAxis = chart.createValueAxis(AxisPosition.RIGHT);
		rightAxis.setCrosses(AxisCrosses.MAX);
		rightAxis.setMinimum(0);
		rightAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
		// set correct cross axis
		bottomAxis.crossAxis(rightAxis);
		rightAxis.crossAxis(bottomAxis);
		XDDFChartData data = chart.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.LINE,
				bottomAxis, rightAxis);
		XDDFChartData.Series s2 = data.addSeries(categoriesData, valuesData2);
		s2.setTitle(series[1], chart.setSheetTitle(series[1], 1));
		chart.plot(data);
		
		chart.setTitleText(chartTitle);
		chart.setTitleOverlay(false);
		XDDFChartLegend legend = chart.getOrAddLegend();
		legend.setPosition(LegendPosition.BOTTOM);
		legend.setOverlay(false);

		chart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getIdx().setVal(1);
		chart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getOrder().setVal(1);
	}

	private void setPieData(XWPFChart chart, String chartTitle, String[] series, String[] categories,
			List<Double[]> list) {
		final int numOfPoints = categories.length;
		final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
		final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
		XDDFChartData pie = new XDDFPieChartData(chart.getCTChart().getPlotArea().addNewPieChart());
		for (int i = 0; i < list.size(); i++) {
			final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, i + 1, i + 1));
			final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(list.get(i),
					valuesDataRange, i);
			XDDFChartData.Series s = pie.addSeries(categoriesData, valuesData);
			s.setTitle(series[i], chart.setSheetTitle(series[i], i));
		}

		pie.setVaryColors(true);
		chart.plot(pie);

		XDDFChartLegend legend = chart.getOrAddLegend();
		legend.setPosition(LegendPosition.RIGHT);
		legend.setOverlay(false);

		chart.setTitleText(chartTitle);
		chart.setTitleOverlay(false);
	}

	private void setLineData(XWPFChart chart, String chartTitle, String[] series, String[] categories,
			List<Double[]> list) {
		XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);		
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
		XDDFLineChartData line = (XDDFLineChartData) chart
				.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.LINE, bottomAxis, leftAxis);
		final int numOfPoints = categories.length;
		final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
		final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
		for (int i = 0; i < list.size(); i++) {
			final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, i + 1, i + 1));
			final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(list.get(i),
					valuesDataRange, i);
			XDDFChartData.Series s = line.addSeries(categoriesData, valuesData);
			s.setTitle(series[i], chart.setSheetTitle(series[i], i));
		}
		chart.plot(line);
		if (list.size() > 1) {
			XDDFChartLegend legend = chart.getOrAddLegend();
			legend.setPosition(LegendPosition.BOTTOM);
			legend.setOverlay(false);
		}
		chart.setTitleText(chartTitle); // https://stackoverflow.com/questions/30532612
		chart.setTitleOverlay(false);
	}

	private void setRadarData(XWPFChart chart, String chartTitle, String[] series, String[] categories,
			List<Double[]> list) {
		XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		XDDFRadarChartData radar = (XDDFRadarChartData) chart
				.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.RADAR, bottomAxis, leftAxis);
		final int numOfPoints = categories.length;
		final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
		final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);

		for (int i = 0; i < list.size(); i++) {
			final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, i + 1, i + 1));
			final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(list.get(i),
					valuesDataRange, i);
			XDDFChartData.Series s = radar.addSeries(categoriesData, valuesData);
			s.setTitle(series[i], chart.setSheetTitle(series[i], i));
		}
		radar.setStyle(RadarStyle.STANDARD);
		chart.plot(radar);
		if (list.size() > 1) {
			XDDFChartLegend legend = chart.getOrAddLegend();
			legend.setPosition(LegendPosition.BOTTOM);
			legend.setOverlay(false);
		}
		chart.setTitleText(chartTitle); // https://stackoverflow.com/questions/30532612
		chart.setTitleOverlay(false);

	}

	public XWPFDocument getWxpf() {
		return wxpf;
	}

	public void setWxpf(CustomXWPFDocument wxpf) {
		this.wxpf = wxpf;
	}

	/**  
	* @Description: 单元格横向合并 
	* @param table  XWPFTable对象引用 
	* @param row  单元格所在行 
	* @param fromCol 合并起始列
	* @param toCol 合并终止列 
	* @return XWPFTable对象引用 
	* @throws 
	*/
	public XWPFTable mergeHCells(XWPFTable table, int row, int fromCol, int toCol) {
		// First Col
		CTHMerge hMerge = CTHMerge.Factory.newInstance();
		hMerge.setVal(STMerge.RESTART);
		CTHMerge hMerge1 = CTHMerge.Factory.newInstance();
		hMerge1.setVal(STMerge.CONTINUE);
		for (int i = fromCol; i <= toCol; i++) {
			if (i == fromCol)
				table.getRow(row).getCell(i).getCTTc().getTcPr().setHMerge(hMerge);
			else
				table.getRow(row).getCell(i).getCTTc().getTcPr().setHMerge(hMerge1);
		}
		return table;
	}

	/**  
	* @Description: 单元格竖向合并 
	* @param table  XWPFTable对象引用 
	* @param col  单元格所在列
	* @param fromRow 合并起始行
	* @param toRow 合并终止行 
	* @return XWPFTable对象引用 
	* @throws 
	*/
	public XWPFTable mergeVCells(XWPFTable table, int col, int fromRow, int toRow) {
		// First Row
		CTVMerge vMerge = CTVMerge.Factory.newInstance();
		vMerge.setVal(STMerge.RESTART);
		// other Rows
		CTVMerge vMerge1 = CTVMerge.Factory.newInstance();
		vMerge1.setVal(STMerge.CONTINUE);
		for (int i = fromRow; i <= toRow; i++) {
			if (i == fromRow)
				table.getRow(i).getCell(col).getCTTc().getTcPr().setVMerge(vMerge);
			else
				table.getRow(i).getCell(col).getCTTc().getTcPr().setVMerge(vMerge1);
		}
		return table;
	}
	
	/**  
	* @Description: 设置单元格属性 
	* @param table  XWPFTable对象引用 
	* @param tc TableCellProperty对象引用
	* @return 
	* @throws 
	*/
	public void setTableCell(XWPFTable table, TableCellProperty tc) {
		List<XWPFTableRow> rows = table.getRows();
		int rowCt = 0;
		int colCt = 0;
		for (XWPFTableRow row : rows) {
			if (rowCt == tc.getRow()) {
				// get the cells in this row
				List<XWPFTableCell> cells = row.getTableCells();
				for (XWPFTableCell cell : cells) {
					if (colCt == tc.getColumn()) {
						// get a table cell properties element (tcPr)
						CTTcPr tcpr = cell.getCTTc().addNewTcPr();
						// set vertical alignment to "center"
						CTVerticalJc va = tcpr.addNewVAlign();
						va.setVal(STVerticalJc.CENTER);
						//TableCellBorder tborder = tc.getBorder();
						if (tc.getHave() == TableCellBorderTypes.YES) {
							CTTcBorders tcBorders = tcpr.addNewTcBorders();
							CTBorder bBorder = tcBorders.addNewBottom();
							bBorder.setColor("000000");
							bBorder.setVal(STBorder.Enum.forString("single"));
							bBorder.setSz(new BigInteger("1"));
							CTBorder tBorder = tcBorders.addNewTop();
							tBorder.setColor("000000");
							tBorder.setVal(STBorder.Enum.forString("single"));
							tBorder.setSz(new BigInteger("1"));
							CTBorder lBorder = tcBorders.addNewLeft();
							lBorder.setColor("000000");
							lBorder.setVal(STBorder.Enum.forString("single"));
							lBorder.setSz(new BigInteger("1"));
							CTBorder rBorder = tcBorders.addNewRight();
							rBorder.setColor("000000");
							rBorder.setVal(STBorder.Enum.forString("single"));
							rBorder.setSz(new BigInteger("1"));
						} else {
							CTTcBorders tcBorders = tcpr.addNewTcBorders();
							CTBorder bBorder = tcBorders.addNewBottom();
							bBorder.setColor("000000");
							bBorder.setVal(STBorder.Enum.forString("single"));
							bBorder.setSz(new BigInteger("1"));
							CTBorder tBorder = tcBorders.addNewTop();
							tBorder.setColor("000000");
							tBorder.setVal(STBorder.Enum.forString("single"));
							tBorder.setSz(new BigInteger("1"));
							CTBorder lBorder = tcBorders.addNewLeft();
							// lBorder.setColor("000000");
							lBorder.setVal(STBorder.Enum.forString("none"));
							// lBorder.setSz(new BigInteger("1"));
							CTBorder rBorder = tcBorders.addNewRight();
							// rBorder.setColor("000000");
							rBorder.setVal(STBorder.Enum.forString("none"));
							// rBorder.setSz(new BigInteger("1"));
						}
						// get 1st paragraph in cell's paragraph list
						XWPFParagraph para = cell.getParagraphs().get(0);
						// create a run to contain the content
						XWPFRun rh = para.createRun();
						TableCellText text = tc.getText();						
						para.setAlignment(ParagraphAlignment.CENTER);
						rh.setFontSize(text.getFontSize());
						rh.setFontFamily(text.getFontFamily());
						rh.setBold(text.isFontBold());
						rh.setText(text.getText());
						// create cell color element
						CTShd ctshd = tcpr.addNewShd();
						ctshd.setColor("auto");
						ctshd.setVal(STShd.CLEAR);
						// header row
						ctshd.setFill(tc.getBgColor());
						break;
					}
					colCt++;
				}
				break;
			}
			colCt = 0;
			rowCt++;
		}
	}

	private void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {

		CTStyle ctStyle = CTStyle.Factory.newInstance();
		ctStyle.setStyleId(strStyleId);

		CTString styleName = CTString.Factory.newInstance();
		styleName.setVal(strStyleId);
		ctStyle.setName(styleName);

		CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
		indentNumber.setVal(BigInteger.valueOf(headingLevel));

		// lower number > style is more prominent in the formats bar
		ctStyle.setUiPriority(indentNumber);

		CTOnOff onoffnull = CTOnOff.Factory.newInstance();
		ctStyle.setUnhideWhenUsed(onoffnull);

		// style shows up in the formats bar
		ctStyle.setQFormat(onoffnull);

		// style defines a heading of the given level
		CTPPr ppr = CTPPr.Factory.newInstance();
		ppr.setOutlineLvl(indentNumber);
		ctStyle.setPPr(ppr);

		XWPFStyle style = new XWPFStyle(ctStyle);

		// is a null op if already defined
		XWPFStyles styles = docxDocument.createStyles();

		style.setType(STStyleType.PARAGRAPH);
		styles.addStyle(style);

	}

	/**  
	* @Description: 创建页眉
	* @param headerText 页眉文本
	* @return 
	* @throws 
	*/
	@Override
	public void createHeader(String headerText) throws Exception {
		// TODO Auto-generated method stub
		CTSectPr sectPr = wxpf.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(wxpf, sectPr);

		// write header content
		CTP ctpHeader = CTP.Factory.newInstance();
		CTR ctrHeader = ctpHeader.addNewR();
		CTText ctHeader = ctrHeader.addNewT();		
		ctHeader.setStringValue(headerText);
		XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, wxpf);
		XWPFParagraph[] parsHeader = new XWPFParagraph[1];
		parsHeader[0] = headerParagraph;
		policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

	}

	/**  
	* @Description: 创建页脚(页码)
	* @return 
	* @throws 
	*/
	@Override
	public void createFooter() throws Exception {
		// TODO Auto-generated method stub
		CTP ctp = CTP.Factory.newInstance();
//this add page number incremental
		ctp.addNewR().addNewPgNum();
		XWPFParagraph codePara = new XWPFParagraph(ctp, wxpf);
		XWPFParagraph[] paragraphs = new XWPFParagraph[1];
		paragraphs[0] = codePara;
//position of number
		codePara.setAlignment(ParagraphAlignment.CENTER);
		CTSectPr sectPr = wxpf.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(wxpf, sectPr);
		headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, paragraphs);
	}

}
