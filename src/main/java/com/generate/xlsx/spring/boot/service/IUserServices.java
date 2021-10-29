package com.generate.xlsx.spring.boot.service;

import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import com.generate.xlsx.spring.boot.model.User;

public interface IUserServices {

	public List<User> listAll();

	public void downloadToExcel(HttpServletResponse response) throws IOException;
	
	public void saveToExcel() throws IOException;

}