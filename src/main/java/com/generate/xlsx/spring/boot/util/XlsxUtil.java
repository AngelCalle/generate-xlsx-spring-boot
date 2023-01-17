package com.generate.xlsx.spring.boot.util;

import java.io.IOException;

import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.tomcat.util.http.fileupload.ByteArrayOutputStream;

import org.springframework.stereotype.Component;

import com.generate.xlsx.spring.boot.model.ResultDto;

import java.io.ByteArrayInputStream;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Component
public class XlsxUtil {

	Sheet xSSFSheet;

	protected XlsxUtil() {
	}

	public ByteArrayInputStream create(List<String> headerLiterals, String nameSheet, List<ResultDto> listResultDto) {
		try (XSSFWorkbook xSSFWorkbook = new XSSFWorkbook();
				ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();) {
			writeHeaderLine(xSSFWorkbook, nameSheet, headerLiterals);
			writeDataLines(xSSFWorkbook, listResultDto);
			xSSFWorkbook.write(byteArrayOutputStream);
			return new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
		} catch (IOException e) {
			e.printStackTrace();
			return null;
		}
	}

	private void writeHeaderLine(XSSFWorkbook xSSFWorkbook, String nameSheet, List<String> headerLiterals) {
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

	private void createCellHyperlink(XSSFWorkbook xSSFWorkbook, Row row, int columnCount, Object name, Object linkValue,
			CellStyle cellStyle) {
		xSSFSheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
		XSSFHyperlink link = xSSFWorkbook.getCreationHelper().createHyperlink(HyperlinkType.FILE);
		link.setAddress((String) linkValue);
		cell.setHyperlink(link);
		cell.setCellValue((String) name);
		cell.setCellStyle(cellStyle);
	}

	private void writeDataLines(XSSFWorkbook xSSFWorkbook, List<ResultDto> listResultDto) {
		int rowCount = 1;

		XSSFFont xSSFFont = xSSFWorkbook.createFont();
		xSSFFont.setFontHeight(14);

		CellStyle cellStyle = xSSFWorkbook.createCellStyle();
		cellStyle.setFont(xSSFFont);

		XSSFFont xSSFFontHyperlink = xSSFWorkbook.createFont();
		xSSFFontHyperlink.setFontHeight(14);
		xSSFFontHyperlink.setColor(IndexedColors.BLUE.index);

		CellStyle cellStyleHyperlink = xSSFWorkbook.createCellStyle();
		cellStyleHyperlink.setFont(xSSFFontHyperlink);

		int columnCount = 0;
		for (ResultDto dto : listResultDto) {
			Row row = xSSFSheet.createRow(rowCount++);
			createCell(row, columnCount++, dto.getName(), cellStyle);
			createCell(row, columnCount++, dto.getType(), cellStyle);
			createCellHyperlink(xSSFWorkbook, row, columnCount++, dto.getName(), dto.getHyperlink(),
					cellStyleHyperlink);
		}
	}

}