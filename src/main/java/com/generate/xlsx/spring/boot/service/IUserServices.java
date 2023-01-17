package com.generate.xlsx.spring.boot.service;

import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.core.io.InputStreamResource;

import com.generate.xlsx.spring.boot.model.User;

public interface IUserServices {

	List<User> listAll();

	void downloadToExcel(HttpServletResponse response);

	void saveToExcel();

	InputStreamResource downloadToExcel1();

	void downloadToExcel2(HttpServletResponse response);

}