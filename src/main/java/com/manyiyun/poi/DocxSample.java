package com.manyiyun.poi;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.manyiyun.poi.data.Hospital;
import com.manyiyun.poi.enums.ChartTypes;

import com.manyiyun.poi.enums.TableCellBorderTypes;
import com.manyiyun.poi.model.Chart;

import com.manyiyun.poi.model.TableCellProperty;
import com.manyiyun.poi.model.TableCellText;

import com.manyiyun.poi.model.TableHeader;
import com.manyiyun.poi.model.TableProperty;
import com.manyiyun.poi.util.ChartFactory;

public class DocxSample {

	private static final String CHART_DATA_FILE = "bar-chart-data.txt";
	private static final String DOCX_OUT_FILE = "e:/manyiyun.docx";

	// get file from classpath, resources folder
	private File readFileFromClasspath(String fileName) {
		ClassLoader classLoader = getClass().getClassLoader();
		URL resource = classLoader.getResource(fileName);
		if (resource == null) {
			throw new IllegalArgumentException("file is not found!");
		} else {
			return new File(resource.getFile());
		}
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		DocxSample ds = new DocxSample();
		File df = ds.readFileFromClasspath(CHART_DATA_FILE);
		try (BufferedReader modelReader = new BufferedReader(new FileReader(df))) {

			String chartTitle = modelReader.readLine(); // first line is chart title
			String[] series = modelReader.readLine().split(",");

			// Category Axis Data
			List<String> listLanguages = new ArrayList<>(10);

			// Values
			List<Double> listCountries = new ArrayList<>(10);
			List<Double> listSpeakers = new ArrayList<>(10);

			// set model
			String ln;
			while ((ln = modelReader.readLine()) != null) {
				String[] vals = ln.split(",");
				listCountries.add(Double.valueOf(vals[0]));
				listSpeakers.add(Double.valueOf(vals[1]));
				listLanguages.add(vals[2]);
			}
			String[] categories = listLanguages.toArray(new String[listLanguages.size()]);
			Double[] values1 = listCountries.toArray(new Double[listCountries.size()]);
			Double[] values2 = listSpeakers.toArray(new Double[listSpeakers.size()]);

			List<Double[]> pieValues = new ArrayList<Double[]>();
			pieValues.add(values1);
			// values.add(values2);

			List<Double[]> barValues = new ArrayList<Double[]>();
			barValues.add(values1);
			barValues.add(values2);

			List<Double[]> lineValues = new ArrayList<Double[]>();
			lineValues.add(values1);
			lineValues.add(values2);

			List<Double[]> radarValues = new ArrayList<Double[]>();
			radarValues.add(values1);

			List<Hospital> hos = new ArrayList<Hospital>();
			Hospital h1 = new Hospital();
			h1.setAddress("北京市西城区南礼士路8号");
			h1.setArea(12897.2d);
			h1.setBedNum(356);
			h1.setEnd(new Date());
			h1.setName("儿童医院");
			h1.setSj(true);
			Hospital h2 = new Hospital();
			h2.setAddress("北京市西长安街248号");
			h2.setArea(18280.50d);
			h2.setBedNum(545);
			h2.setEnd(new Date());
			h2.setName("北京301医院");
			h2.setSj(true);
			Hospital h3 = new Hospital();
			h3.setAddress("北京市石景山区西五环外路28号");
			h3.setArea(16280.50d);
			h3.setBedNum(480);
			h3.setEnd(new Date());
			h3.setName("北京朝阳医院(石景山分院)");
			h3.setSj(false);
			hos.add(h1);
			hos.add(h2);
			hos.add(h3);

			// 创建表格表头对象(2行5列)
			TableHeader tHeader = new TableHeader(2, 5);
			// 创建表格属性对象
			TableProperty<Hospital> tbl = new TableProperty<Hospital>(tHeader, hos);

			// 创建DOCX文档生成器对象
			SimpleDocxMaker instance = new SimpleDocxMaker();
			// 创建DOCX文档页眉
			instance.createHeader("2019年公立医院社会效益调查患者满意度报告");
			// 创建DOCX文档页脚
			instance.createFooter();
			// 获取DOCX文档对象引用
			XWPFDocument doc = instance.getWxpf();
			// 创建段落
			XWPFParagraph para = doc.createParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
			// 创建文本
			XWPFRun run = para.createRun();
			run.setBold(true);
			run.setFontFamily("仿宋");
			run.setFontSize(24);
			// run.setTextPosition(20);
			run.setText("委属医院平均调查项目");
			// 断行
			run.addBreak();
			run.addBreak();
			para = doc.createParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
			run = para.createRun();
			run.setBold(true);
			run.setFontFamily("仿宋");
			run.setFontSize(24);
			// run.setTextPosition(20);
			run.setText("公立三级医院社会效益调查患者满意度报告");
			run.addBreak();
			para = doc.createParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
			run = para.createRun();
			run.setBold(true);
			run.setFontFamily("仿宋");
			run.setFontSize(24);
			// run.setTextPosition(20);
			run.setText("（2019年）");
			for (int i = 0; i < 13; i++) {
				// run.addBreak();
				run.addCarriageReturn();
			}
			para = doc.createParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
			run = para.createRun();
			run.setBold(true);
			run.setFontFamily("仿宋");
			run.setFontSize(20);
			// run.setTextPosition(20);
			run.setText("XXX信息科技有限公司");
			para = doc.createParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);
			run = para.createRun();
			run.setBold(true);
			run.setFontFamily("仿宋");
			run.setFontSize(20);
			// run.setTextPosition(20);
			run.setText("二零一九年八月");
			// 创建换页符(下一页)
			run.addBreak(BreakType.PAGE);
			// 创建文档目录大纲
			instance.createToc();
			// 标题 (1级)
			instance.addTitle(1, "标题1", true);
			// 正文部分
			// 标题 (2级)
			instance.addTitle(2, "标题1.1", false);
			// 正文部分
			String text = "开始测试第一段文字：Apache POI包括一系列的API，它们可以操作基于MicroSoft OLE 2 Compound Document Format的各种格式文件，可以通过这些API在Java中读写Excel、Word等文件。\r\n"
					+ "\r\n" + "　　优点：跨平台支持windows、unix和linux。\r\n" + "\r\n"
					+ "　　缺点：ache.poi.poifs.crypt.dsig.ExpiredCertificateSecurityException\r\n"
					+ "org.apache.poi.poifs.crypt.dsig.RevokedCertificateSecurityException\r\n"
					+ "org.apache.poi.poifs.crypt.dsig.TrustCertificateSecurityException\r\n"
					+ "org.apache.poi.xdgf.usermodel.shape.exceptions.StopVisiting\r\n"
					+ "org.apache.poi.xdgf.usermodel.shape.exceptions.StopVisitingThisBranch相对与对word文件的处理来说，POI更适合excel处理，对于word实现一些简单文件的操作凑合，不能设置样式且生成的word文件格式不够规范。";
			// 创建文本
			instance.createRegion(ParagraphAlignment.LEFT, 480, text, false, "000000", "宋体 (中文正文)", 11);
			instance.createRegion(ParagraphAlignment.LEFT, 480, "开始测试第二段文字：I包括一系列的API，它们可以操作基于MicroSoft ", true,
					"000000", "宋体 (中文正文)", 11);
			// 标题 (3级)
			instance.addTitle(3, "标题1.1.1", false);
			// 正文部分
			instance.createRegion(ParagraphAlignment.CENTER, 0, "表-1", true, "000000", "宋体 (中文正文)", 11);
			// 创建表格
			XWPFTable table = instance.createTable(tbl);
			// 横向表格合并
			table = instance.mergeHCells(table, 0, 0, 2);
			// 设置表头标题
			TableCellText cellText = new TableCellText("概况");
			TableCellProperty tc = new TableCellProperty(0, 0, TableCellBorderTypes.NO, cellText);
			// 设置单元格属性
			instance.setTableCell(table, tc);
			cellText.setText("评级");
			TableCellProperty tc1 = new TableCellProperty(0, 3, TableCellBorderTypes.NO, cellText);
			instance.setTableCell(table, tc1);
			table = instance.mergeVCells(table, 4, 0, 1);
			cellText.setText("日期");
			TableCellProperty tc2 = new TableCellProperty(0, 4, TableCellBorderTypes.NO, cellText);
			// 设置单元格背景色
			tc2.setBgColor("C3599D");
			instance.setTableCell(table, tc2);
			// 设置表头标题
			String[] headerTitle = { "医院名称", "医院地址", "占地面积", "是否三甲" };
			for (int i = 0; i < 4; i++) {
				TableCellText ct = new TableCellText(headerTitle[i]);
				TableCellProperty tcp = new TableCellProperty(1, i, TableCellBorderTypes.NO, ct);
				tc2.setBgColor("C3599D");
				instance.setTableCell(table, tcp);
			}

			// 标题 (2级)
			instance.addTitle(2, "标题1.2", false);
			// 正文部分
			File imgFile = ds.readFileFromClasspath("hema.jpeg");
			instance.createPicture(imgFile, 250, 150);
			instance.createRegion(ParagraphAlignment.CENTER, 0, "图-1 动物园的河马", true, "000000", "宋体 (中文正文)", 11);

			// 标题 (1级)
			instance.addTitle(1, "标题2", true);
			// 正文部分
			// 标题 (2级)
			instance.addTitle(2, "标题2.1", false);
			// 正文部分
			// 创建图表(柱图-垂直方向)
			Chart vbar = ChartFactory.getChart(ChartTypes.VBAR, chartTitle, series, categories, barValues, 11, 7);
			instance.createChart(vbar);
			instance.createRegion(ParagraphAlignment.CENTER, 0, "图表-柱图", true, "000000", "宋体 (中文正文)", 11);
			// 创建图表(饼图)
			Chart pie = ChartFactory.getChart(ChartTypes.PIE, chartTitle, series, categories, pieValues, 11, 7);
			instance.createChart(pie);
			instance.createRegion(ParagraphAlignment.CENTER, 0, "图表-饼图", true, "000000", "宋体 (中文正文)", 11);
			// 创建图表(柱图-水平方向)
			Chart bar = ChartFactory.getChart(ChartTypes.BAR, chartTitle, series, categories, barValues, 11, 7);
			instance.createChart(bar);
			instance.createRegion(ParagraphAlignment.CENTER, 0, "图表-柱图", true, "000000", "宋体 (中文正文)", 11);
			// 创建图表(雷达图)
			Chart radar = ChartFactory.getChart(ChartTypes.RADAR, chartTitle, series, categories, radarValues, 11, 7);
			instance.createChart(radar);
			instance.createRegion(ParagraphAlignment.CENTER, 0, "图表-雷达图", true, "000000", "宋体 (中文正文)", 11);
			// 标题 (2级)
			instance.addTitle(2, "标题2.2", false);
			// 正文部分
			// 创建图表(线图)
			Chart line = ChartFactory.getChart(ChartTypes.LINE, chartTitle, series, categories, lineValues, 11, 7);
			instance.createChart(line);
			instance.createRegion(ParagraphAlignment.CENTER, 0, "图表-线图", true, "000000", "宋体 (中文正文)", 11);
			// 创建图表(柱线图-双坐标轴)
			Chart barLine = ChartFactory.getChart(ChartTypes.BAR_LINE, chartTitle, series, categories, lineValues, 11,
					7);
			instance.createChart(barLine);
			instance.createRegion(ParagraphAlignment.CENTER, 0, "图表-柱线图", true, "000000", "宋体 (中文正文)", 11);
			// 输出文档
			instance.write(DOCX_OUT_FILE);
			System.out.println("done!");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
