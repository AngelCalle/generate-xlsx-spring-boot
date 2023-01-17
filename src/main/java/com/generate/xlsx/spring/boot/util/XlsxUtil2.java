package com.generate.xlsx.spring.boot.util;

import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.generate.xlsx.spring.boot.model.ResultDto;

public class XlsxUtil2 {

	private XSSFWorkbook xSSFWorkbook;
	private XSSFSheet xSSFSheet;
	private List<String> headerLiterals;
	private String nameSheet;
	private List<ResultDto> listResultDto;

	public XlsxUtil2(List<String> headerLiterals, String nameSheet, List<ResultDto> listResultDto) {
		this.headerLiterals = headerLiterals;
		this.nameSheet = nameSheet;
		this.listResultDto = listResultDto;
		xSSFWorkbook = new XSSFWorkbook();
	}

	public void export(HttpServletResponse httpServletResponse) {
		writeHeaderLine();
		writeDataLines();
		try (ServletOutputStream outputStream = httpServletResponse.getOutputStream()) {
			xSSFWorkbook.write(outputStream);
			xSSFWorkbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void writeHeaderLine() {
		xSSFSheet = xSSFWorkbook.createSheet(nameSheet);
		Row row = xSSFSheet.createRow(0);

		XSSFFont xSSFFont = xSSFWorkbook.createFont();
		xSSFFont.setBold(true);
		xSSFFont.setFontHeight(16);
		xSSFFont.setColor(IndexedColors.BLACK.index);
		xSSFFont.setFontName("Arial");

		CellStyle cellStyle = xSSFWorkbook.createCellStyle();
		cellStyle.setFont(xSSFFont);
		cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		AtomicInteger index = new AtomicInteger();
		headerLiterals.forEach(literal -> createCell(row, index.getAndIncrement(), literal, cellStyle));

		for (int i = 0; i < xSSFSheet.getRow(0).getLastCellNum(); i++) {
			xSSFSheet.autoSizeColumn(i);
		}
	}

	private void createCell(Row row, int columnCount, Object value, CellStyle cellStyle) {
		xSSFSheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
		if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue(value.toString());
		} else {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(cellStyle);
	}

	private void createCellHyperlink(Row row, int columnCount, Object name, Object linkValue, CellStyle cellStyle) {
		xSSFSheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
		XSSFHyperlink link = xSSFWorkbook.getCreationHelper().createHyperlink(HyperlinkType.FILE);
		link.setAddress((String) linkValue);
		cell.setHyperlink(link);
		cell.setCellValue((String) name);
		cell.setCellStyle(cellStyle);
	}

	private void writeDataLines() {
		int rowCount = 1;
		int columnCount = 0;

		XSSFFont xSSFFont = xSSFWorkbook.createFont();
		xSSFFont.setFontHeight(14);

		CellStyle cellStyle = xSSFWorkbook.createCellStyle();
		cellStyle.setFont(xSSFFont);

		XSSFFont xSSFFontHyperlink = xSSFWorkbook.createFont();
		xSSFFontHyperlink.setFontHeight(14);
		xSSFFontHyperlink.setColor(IndexedColors.BLUE.index);

		CellStyle cellStyleHyperlink = xSSFWorkbook.createCellStyle();
		cellStyleHyperlink.setFont(xSSFFontHyperlink);

		for (ResultDto dto : listResultDto) {
			Row row = xSSFSheet.createRow(rowCount++);
			createCell(row, columnCount++, dto.getName(), cellStyle);
			createCell(row, columnCount++, dto.getType(), cellStyle);
			createCellHyperlink(row, columnCount++, dto.getName(), dto.getHyperlink(), cellStyleHyperlink);
		}
	}

}