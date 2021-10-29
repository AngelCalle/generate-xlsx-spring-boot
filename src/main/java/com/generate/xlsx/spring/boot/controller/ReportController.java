package com.generate.xlsx.spring.boot.controller;

import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
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
	private IUserServices userServices;

	// http://localhost:8080/report/download-excel
	@GetMapping("/download-excel")
	public void downloadToExcel() {
		try {
			userServices.downloadToExcel(httpServletResponse);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	// http://localhost:8080/report/save-excel
	@GetMapping("/save-excel")
	public void saveToExcel() {		
		try {
			userServices.saveToExcel();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
