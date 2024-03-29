package com.generate.xlsx.spring.boot.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.generate.xlsx.spring.boot.model.User;

public class UserExcelExporter {

	// http://poi.apache.org/components/spreadsheet/examples.html
	// https://svn.apache.org/repos/asf/poi/trunk/poi-examples/src/main/java/org/apache/poi/examples/ss/LoanCalculator.java

	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private List<User> listUsers;

	public UserExcelExporter(List<User> listUsers) {
		this.listUsers = listUsers;
		workbook = new XSSFWorkbook();
	}

	private void writeHeaderLine() {
		sheet = workbook.createSheet("Users");

		Row row = sheet.createRow(0);

		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(16);
		font.setColor(IndexedColors.ORANGE.index);
		font.setFontName("Arial");
		style.setFont(font);

		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		createCell(row, 0, "User ID", style);
		createCell(row, 1, "E-mail", style);
		createCell(row, 2, "Full Name", style);
		createCell(row, 3, "Roles", style);
		createCell(row, 4, "Enabled", style);

	}

	private void createCell(Row row, int columnCount, Object value, CellStyle style) {
		sheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
		if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue(value.toString());
		} else {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(style);
	}

	private void writeDataLines() {
		int rowCount = 1;

		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setFontHeight(14);
		style.setFont(font);

		for (User user : listUsers) {
			Row row = sheet.createRow(rowCount++);
			int columnCount = 0;
			createCell(row, columnCount++, user.getId(), style);
			createCell(row, columnCount++, user.getEmail(), style);
			createCell(row, columnCount++, user.getFullName(), style);
			createCell(row, columnCount++, user.getRoles().toString(), style);
			createCell(row, columnCount++, user.isEnabled(), style);
		}
	}

	public void export(HttpServletResponse response) {
		writeHeaderLine();
		writeDataLines();

		ServletOutputStream outputStream;
		try {
			outputStream = response.getOutputStream();
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void saveFile() {
		writeHeaderLine();
		writeDataLines();

		String ruta = ".//src//main//resources//archivo.xlsx";
		FileOutputStream out;
		try {
			out = new FileOutputStream(ruta);
			workbook.write(out);
			out.close();
			workbook.close();

			File file = new File(ruta);
			System.out.println(file.length());
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}