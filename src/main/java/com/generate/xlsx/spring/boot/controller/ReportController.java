package com.generate.xlsx.spring.boot.controller;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.generate.xlsx.spring.boot.service.UserServices;

@RestController
@RequestMapping("/report")
public class ReportController {
	
	@Autowired
	private HttpServletResponse httpServletResponse;

	@Autowired
	private UserServices userServices;

	// http://localhost:8080/report/download-excel
	@GetMapping("/download-excel")
	public void downloadToExcel() {
		userServices.downloadToExcel(httpServletResponse);
	}
	
	// http://localhost:8080/report/save-excel
	@GetMapping("/save-excel")
	public void saveToExcel() {		
		userServices.saveToExcel();
	}

}
