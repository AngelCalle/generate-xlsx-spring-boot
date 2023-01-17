package com.generate.xlsx.spring.boot.controller;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.generate.xlsx.spring.boot.service.IUserServices;

@RestController
@RequestMapping("/report")
public class ReportController {

	@Autowired
	private HttpServletResponse httpServletResponse;

	@Autowired
	private IUserServices service;

	// http://localhost:8080/report/download-excel
	@GetMapping("/download-excel")
	public void downloadToExcel() {
		service.downloadToExcel(httpServletResponse);
	}

	// http://localhost:8080/report/download-excel1
	@GetMapping("/download-excel1")
	public ResponseEntity<InputStreamResource> downloadToExcel1() {
		DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
		String currentDateTime = dateFormatter.format(new Date());
		String fileName = currentDateTime + ".xlsx";
		String headerValue = "attachment; filename=Report__";

		HttpHeaders headers = new HttpHeaders();
		headers.add(HttpHeaders.CONTENT_DISPOSITION, headerValue + fileName);
		return ResponseEntity.ok().headers(headers).contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
				.body(service.downloadToExcel1());
	}

	// http://localhost:8080/report/download-excel2
	@GetMapping("/download-excel2")
	public void downloadToExcel2(HttpServletResponse httpServletResponse) {
		DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
		String currentDateTime = dateFormatter.format(new Date());
		String fileName = currentDateTime + ".xlsx";
		String headerValue = "attachment; filename=Report__";

		httpServletResponse.setContentType("application/octet-stream");
		httpServletResponse.setHeader(HttpHeaders.CONTENT_DISPOSITION, headerValue + fileName);
		service.downloadToExcel2(httpServletResponse);
	}

	// http://localhost:8080/report/save-excel
	@GetMapping("/save-excel")
	public void saveToExcel() {
		service.saveToExcel();
	}

}
