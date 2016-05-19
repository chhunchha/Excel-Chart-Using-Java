package com.programmingfree.excelexamples;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.LineChartSeries;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.charts.XSSFChartAxis;
import org.apache.poi.xssf.usermodel.charts.XSSFChartLegend;
import org.apache.poi.xssf.usermodel.charts.XSSFLineChartData;
import org.apache.poi.xssf.usermodel.charts.XSSFValueAxis;
import org.apache.xmlbeans.SchemaType;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFChart;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.jfree.data.general.DefaultPieDataset;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.ChartUtilities;
//import org.json.JSONArray;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLayout;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.impl.CTBooleanImpl;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

import com.smartxls.ChartFormat;
import com.smartxls.ChartShape;
import com.smartxls.WorkBook;
import com.sun.corba.se.spi.orbutil.threadpool.Work;

public class CreateExcelFile {
	public static void main(String args[]) throws IOException, InvalidFormatException, ParseException {
		generateExcelChart();
	}

	private static Map<String, CellStyle> createStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

		// http://poi.apache.org/spreadsheet/quick-guide.html#DataFormats

		short borderColor = IndexedColors.GREY_50_PERCENT.getIndex();

		CellStyle style;
		Font companyNameFont = wb.createFont();
		companyNameFont.setFontHeightInPoints((short) 48);
		companyNameFont.setColor(IndexedColors.DARK_BLUE.getIndex());
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setFont(companyNameFont);
		styles.put("company_name", style);

		Font label_font = wb.createFont();
		label_font.setFontHeightInPoints((short) 14);
		label_font.setBold(true);
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		style.setFont(label_font);
		styles.put("label_font", style);

		Font data_font = wb.createFont();
		data_font.setFontHeightInPoints((short) 14);
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		style.setFont(data_font);
		styles.put("data_font", style);

		Font percent_data_font = wb.createFont();
		percent_data_font.setFontHeightInPoints((short) 14);
		style = wb.createCellStyle();
		style.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		style.setFont(percent_data_font);
		styles.put("percent_data_font", style);

		Font survey_description = wb.createFont();
		survey_description.setFontHeightInPoints((short) 14);
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		style.setFont(survey_description);
		style.setWrapText(true);
		styles.put("survey_description", style);

		Font titleFont = wb.createFont();
		titleFont.setFontHeightInPoints((short) 48);
		titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setFont(titleFont);
		styles.put("title", style);

		// Font monthFont = wb.createFont();
		// monthFont.setFontHeightInPoints((short) 12);
		// monthFont.setColor(IndexedColors.WHITE.getIndex());
		// monthFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		// style = wb.createCellStyle();
		// style.setAlignment(CellStyle.ALIGN_CENTER);
		// style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		// style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
		// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		// style.setFont(monthFont);
		// styles.put("month", style);
		//
		// Font dayFont = wb.createFont();
		// dayFont.setFontHeightInPoints((short) 14);
		// dayFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		// style = wb.createCellStyle();
		// style.setAlignment(CellStyle.ALIGN_LEFT);
		// style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		// style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		// style.setBorderLeft(CellStyle.BORDER_THIN);
		// style.setLeftBorderColor(borderColor);
		// style.setBorderBottom(CellStyle.BORDER_THIN);
		// style.setBottomBorderColor(borderColor);
		// style.setFont(dayFont);
		// styles.put("weekend_left", style);
		//
		// style = wb.createCellStyle();
		// style.setAlignment(CellStyle.ALIGN_CENTER);
		// style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		// style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		// style.setBorderRight(CellStyle.BORDER_THIN);
		// style.setRightBorderColor(borderColor);
		// style.setBorderBottom(CellStyle.BORDER_THIN);
		// style.setBottomBorderColor(borderColor);
		// styles.put("weekend_right", style);
		//
		// style = wb.createCellStyle();
		// style.setAlignment(CellStyle.ALIGN_LEFT);
		// style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		// style.setBorderLeft(CellStyle.BORDER_THIN);
		// style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		// style.setLeftBorderColor(borderColor);
		// style.setBorderBottom(CellStyle.BORDER_THIN);
		// style.setBottomBorderColor(borderColor);
		// style.setFont(dayFont);
		// styles.put("workday_left", style);
		//
		// style = wb.createCellStyle();
		// style.setAlignment(CellStyle.ALIGN_CENTER);
		// style.setVerticalAlignment(CellStyle.VERTICAL_TOP);
		// style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		// style.setBorderRight(CellStyle.BORDER_THIN);
		// style.setRightBorderColor(borderColor);
		// style.setBorderBottom(CellStyle.BORDER_THIN);
		// style.setBottomBorderColor(borderColor);
		// styles.put("workday_right", style);
		//
		// style = wb.createCellStyle();
		// style.setBorderLeft(CellStyle.BORDER_THIN);
		// style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		// style.setBorderBottom(CellStyle.BORDER_THIN);
		// style.setBottomBorderColor(borderColor);
		// styles.put("grey_left", style);
		//
		// style = wb.createCellStyle();
		// style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		// style.setBorderRight(CellStyle.BORDER_THIN);
		// style.setRightBorderColor(borderColor);
		// style.setBorderBottom(CellStyle.BORDER_THIN);
		// style.setBottomBorderColor(borderColor);
		// styles.put("grey_right", style);

		return styles;
	}

	// Will add line chart by default and excel macro will do the remaining task of converting it to appropriate chart.
	public static void createChart(XSSFWorkbook wb, XSSFSheet sheet, String question, int dataRangeStartRow, int dataRangeEndRow) throws IOException {		
		XSSFDrawing drawing = sheet.createDrawingPatriarch();
		
//		col1 - the column (0 based) of the first cell.
//		row1 - the row (0 based) of the first cell.
//		col2 - the column (0 based) of the second cell.
//		row2 - the row (0 based) of the second cell.
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 3, 12, 30);

        XSSFChart chart  = drawing.createChart(anchor);
        XSSFChartLegend legend = chart.getOrCreateLegend();
        legend.setPosition(LegendPosition.RIGHT);
        
        XSSFLineChartData data = chart.getChartDataFactory().createLineChartData();

        // Use a category axis for the bottom axis.
        XSSFChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        XSSFValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        ChartDataSource<Number> xs = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(dataRangeStartRow, dataRangeEndRow, 0,  0));
        ChartDataSource<Number> ys = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(dataRangeStartRow, dataRangeEndRow, 1,  1));

        data.addSeries(xs, ys);
        
        CTChart ctChart = chart.getCTChart();
        CTTitle title = ctChart.addNewTitle();
        title.addNewOverlay().setVal(false);
        CTTx tx = title.addNewTx();
        CTTextBody rich = tx.addNewRich();
        rich.addNewBodyPr();  // body properties must exist, but can be empty
        CTTextParagraph para = rich.addNewP();
        CTRegularTextRun r = para.addNewR();
        r.setT(question);
        
        chart.plot(data, bottomAxis, leftAxis);
	}

	// public static void createPieChart(XSSFWorkbook wb,XSSFSheet sheet, String
	// question, int dataRangeStartRow, int dataRangeEndRow) throws IOException
	// {
	// DefaultPieDataset my_pie_chart_data = new DefaultPieDataset();
	// for(int rowIndex = dataRangeStartRow; rowIndex <= dataRangeEndRow;
	// rowIndex++) {
	// String chart_label =
	// sheet.getRow(rowIndex).getCell(0).getStringCellValue();
	// Number chart_data =
	// sheet.getRow(rowIndex).getCell(1).getNumericCellValue();
	// my_pie_chart_data.setValue(chart_label,chart_data);
	// }
	//
	// JFreeChart
	// myPieChart=ChartFactory.createPieChart(question,my_pie_chart_data,true,true,false);
	// int width=640;
	// int height=480;
	// float quality=1;
	//
	// ByteArrayOutputStream chart_out = new ByteArrayOutputStream();
	// ChartUtilities.writeChartAsJPEG(chart_out,quality,myPieChart,width,height);
	//
	// int my_picture_id = wb.addPicture(chart_out.toByteArray(),
	// Workbook.PICTURE_TYPE_JPEG);
	// chart_out.close();
	// XSSFDrawing drawing = sheet.createDrawingPatriarch();
	// ClientAnchor my_anchor = new XSSFClientAnchor();
	// my_anchor.setCol1(4);
	// my_anchor.setRow1(dataRangeStartRow);
	// XSSFPicture my_picture = drawing.createPicture(my_anchor, my_picture_id);
	// my_picture.resize();
	// }

	// public static void createColumnChart(WorkBook wb,Sheet sheet, String
	// question, int dataRangeStartRow, int dataRangeEndRow) throws IOException
	// {
	//
	// ChartShape chart = wb.addChart(1, 7, 13, 31);
	// chart.setChartType(ChartShape.Column);
	//
	// chart.setLinkRange(sheet.getSheetName() + "!$a$" + dataRangeStartRow +
	// ":$b$" + dataRangeEndRow, false);
	// chart.setAxisTitle(ChartShape.XAxis, 0, "X-axis data");
	// chart.setAxisTitle(ChartShape.YAxis, 0, "Y-axis data");
	// chart.setTitle(question);
	//
	// for(int rowIndex = dataRangeStartRow; rowIndex <= dataRangeEndRow;
	// rowIndex++) {
	// int i = 0;
	// chart.setSeriesName(i,
	// sheet.getRow(rowIndex).getCell(0).getStringCellValue());
	// i++;
	//// String chart_label =
	// sheet.getRow(rowIndex).getCell(0).getStringCellValue();
	//// Number chart_data =
	// sheet.getRow(rowIndex).getCell(1).getNumericCellValue();
	//// my_pie_chart_data.setValue(chart_label,chart_data);
	// }
	//
	// ChartFormat chartFormat = chart.getPlotFormat();
	// chartFormat.setSolid();
	// chartFormat.setForeColor(java.awt.Color.RED.getRGB());
	// chart.setPlotFormat(chartFormat);
	//
	// ChartFormat titleformat = chart.getTitleFormat();
	// titleformat.setFontSize(14*20);
	// titleformat.setFontUnderline(true);
	// titleformat.setTextRotation(90);
	// chart.setTitleFormat(titleformat);
	// }

	private static void generateExcelChart() throws IOException, InvalidFormatException, ParseException {

		InputStream report_excel_in = CreateExcelFile.class.getClass()
				.getResourceAsStream("/com/programmingfree/excelexamples/ReportTemplate.xlsm");
		XSSFWorkbook workbook = new XSSFWorkbook(report_excel_in);
		// XSSFWorkbook workbook = new XSSFWorkbook();

		Map<String, CellStyle> styles = createStyles(workbook);

		CreationHelper createHelper = workbook.getCreationHelper();

		JSONParser parser = new JSONParser();
		InputStream io = CreateExcelFile.class.getClass()
				.getResourceAsStream("/com/programmingfree/excelexamples/data.json");
		BufferedReader br = new BufferedReader(new InputStreamReader(io, "UTF-8"));

		Object jsonData = parser.parse(br);
		JSONObject surveyData = (JSONObject) jsonData;

		String survey_name = (String) surveyData.get("surveyName");
		String survey_description = (String) surveyData.get("surveyDescription");

		// XSSFSheet mainSheet = workbook.createSheet(survey_name);
		XSSFSheet mainSheet = workbook.createSheet(survey_name);

		// create header
		Row headerRow = mainSheet.createRow(0);
		headerRow.setHeightInPoints(80);
		Cell companyNameCell = headerRow.createCell(0);
		companyNameCell.setCellValue("Company Name");
		companyNameCell.setCellStyle(styles.get("company_name"));
		mainSheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$N$1"));

		// System.out.println(survey_name);

		Row surveyNameRow = mainSheet.createRow(1);
		Cell surveyNameTitleCell = surveyNameRow.createCell(0);
		Cell surveyNameValueCell = surveyNameRow.createCell(1);

		surveyNameTitleCell.setCellValue("Survey");
		surveyNameTitleCell.setCellStyle(styles.get("label_font"));

		surveyNameValueCell.setCellValue(survey_name);
		surveyNameValueCell.setCellStyle(styles.get("label_font"));

		mainSheet.addMergedRegion(CellRangeAddress.valueOf("$B$2:$N$2"));

		Row surveyDescRow = mainSheet.createRow(2);
		Cell surveyDescTitleCell = surveyDescRow.createCell(0);
		Cell surveyDescValueCell = surveyDescRow.createCell(1);

		surveyDescTitleCell.setCellValue("Description");
		surveyDescTitleCell.setCellStyle(styles.get("label_font"));

		surveyDescValueCell.setCellValue(survey_description);
		surveyDescValueCell.setCellStyle(styles.get("survey_description"));

		surveyDescRow.setHeightInPoints(100);
		mainSheet.addMergedRegion(CellRangeAddress.valueOf("$B$3:$N$3"));

		mainSheet.autoSizeColumn(0);

		String chartDataRangeName = "";
		String chartNameRangeName = "";

		JSONArray question_answers = (JSONArray) surveyData.get("questionAnswers");

		for (Object questionData : question_answers) {
			// System.out.println(questionData.toString());
			JSONObject questionDataJson = (JSONObject) questionData;

			String question = (String) questionDataJson.get("question");
			System.out.println(question);

			String chartType = (String) questionDataJson.get("chartType");
			System.out.println(chartType);

			boolean hasChart = false;
			XSSFSheet questionSheet;
			if (chartType.equalsIgnoreCase("pie") || chartType.equalsIgnoreCase("bar")
					|| chartType.equalsIgnoreCase("column")) {
				questionSheet = workbook.cloneSheet(workbook.getSheetIndex(workbook.getSheet("chart_sheet")));
				workbook.setSheetName(workbook.getSheetIndex(questionSheet), question);
				hasChart = true;
			} else {
				questionSheet = workbook.createSheet(question);
			}

			Row row = questionSheet.createRow(0);
			Cell titleCell = row.createCell(0);
			Cell valueCell = row.createCell(1);

			titleCell.setCellValue("Question");
			titleCell.setCellStyle(styles.get("label_font"));

			valueCell.setCellValue(question);
			valueCell.setCellStyle(styles.get("data_font"));

			Cell chartTypeCell = row.createCell(25);
			chartTypeCell.setCellValue(chartType);

			JSONObject answers = (JSONObject) questionDataJson.get("answers");
			// System.out.println(answers);

			row = questionSheet.createRow(1);
			titleCell = row.createCell(0);
			titleCell.setCellValue("Responses");
			titleCell.setCellStyle(styles.get("label_font"));

			Iterator<?> iter = answers.keySet().iterator();
			int dataStartRow = 2;
			int dataEndRow = 2;
			int rowIndex = dataStartRow;
			while (iter.hasNext()) {
				String option = iter.next().toString();
				String value = answers.get(option).toString();
				row = questionSheet.createRow(rowIndex);
				titleCell = row.createCell(0);
				valueCell = row.createCell(1);

				titleCell.setCellValue(option);
				titleCell.setCellStyle(styles.get("label_font"));

				if (value.contains("%")) {
					valueCell.setCellValue(Double.parseDouble(value.substring(0, value.indexOf('%'))) / 100);
					valueCell.setCellStyle(styles.get("percent_data_font"));
				} else {
					try {
						valueCell.setCellValue(Double.parseDouble(value));
					} catch (NumberFormatException e) {
						valueCell.setCellValue(value);
					}
					valueCell.setCellStyle(styles.get("data_font"));
				}

				System.out.println(option + " : " + value);
				rowIndex++;
			}
			dataEndRow = rowIndex - 1;
			questionSheet.autoSizeColumn(0);

			if (hasChart) {
				createChart(workbook, questionSheet, question, dataStartRow, dataEndRow);
			}

		}

		workbook.removeSheetAt(workbook.getSheetIndex(workbook.getSheet("chart_sheet")));
		FileOutputStream f = new FileOutputStream("GeneratedReport.xlsm");
		workbook.write(f);
		f.close();

	}
}