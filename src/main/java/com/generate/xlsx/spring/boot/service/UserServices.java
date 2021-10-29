package com.generate.xlsx.spring.boot.service;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;
import javax.transaction.Transactional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import com.generate.xlsx.spring.boot.model.User;
import com.generate.xlsx.spring.boot.repository.UserRepository;
import com.generate.xlsx.spring.boot.util.UserExcelExporter;

@Service
@Transactional
public class UserServices {

	@Autowired
	private UserRepository userRepository;

	public List<User> listAll() {
		return userRepository.findAll(Sort.by("email").ascending());
	}

	public void downloadToExcel(HttpServletResponse response) {
		response.setContentType("application/octet-stream");
		DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
		String currentDateTime = dateFormatter.format(new Date());

		String headerKey = "Content-Disposition";
		String headerValue = "attachment; filename=users_" + currentDateTime + ".xlsx";
		response.setHeader(headerKey, headerValue);

		List<User> listUsers = this.listAll();

		UserExcelExporter excelExporter = new UserExcelExporter(listUsers);

		try {
			excelExporter.export(response);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public void saveToExcel() {
		List<User> listUsers = this.listAll();
		UserExcelExporter excelExporter = new UserExcelExporter(listUsers);
		try {
			excelExporter.saveFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}