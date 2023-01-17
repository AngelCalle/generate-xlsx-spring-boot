package com.generate.xlsx.spring.boot.service.impl;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;
import javax.transaction.Transactional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import com.generate.xlsx.spring.boot.model.ResultDto;
import com.generate.xlsx.spring.boot.model.User;
import com.generate.xlsx.spring.boot.repository.UserRepository;
import com.generate.xlsx.spring.boot.service.IUserServices;
import com.generate.xlsx.spring.boot.util.UserExcelExporter;
import com.generate.xlsx.spring.boot.util.XlsxUtil;
import com.generate.xlsx.spring.boot.util.XlsxUtil2;

@Service
@Transactional
public class UserServicesImpl implements IUserServices {

	@Autowired
	private UserRepository userRepository;

	@Autowired
	private XlsxUtil xlsxUtil;

	public List<User> listAll() {
		return userRepository.findAll(Sort.by("email").ascending());
	}

	public void downloadToExcel(HttpServletResponse httpServletResponse) {
		httpServletResponse.setContentType("application/octet-stream");
		DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH_mm_ss");
		String currentDateTime = dateFormatter.format(new Date());

		String headerKey = "Content-Disposition";
		String headerValue = "attachment; filename=users_" + currentDateTime + ".xlsx";
		httpServletResponse.setHeader(headerKey, headerValue);

		List<User> listUsers = this.listAll();

		UserExcelExporter excelExporter = new UserExcelExporter(listUsers);

		excelExporter.export(httpServletResponse);
	}

	public void saveToExcel() {
		List<User> listUsers = this.listAll();
		UserExcelExporter excelExporter = new UserExcelExporter(listUsers);
		excelExporter.saveFile();
	}

	@Override
	public InputStreamResource downloadToExcel1() {
		List<String> headerLiterals = Arrays.asList("Ruta relativa en disco", "Nombre del archivo", "Tipo",
				"Código Línea Inversión", "UUFF", "CINI", "Type of System", "Objeto control", "Pedido Posición",
				"Importe total de la línea de inversión", "Importe por cada UUFF", "Número documento GL",
				"Fecha contabilización GL", "Año contabilización", "CIF Proveedor",
				"Importe total de la factura (sin IVA)", "Hipervínculo al fichero local para su consulta individual");

		String nameSheet = "name Sheet";

		ResultDto dto = new ResultDto();
		dto.setName("Name");
		dto.setType("Type");
		dto.setHyperlink("C:/Users/acalles.ext/Desktop/ENEL/logback.xml");
		List<ResultDto> listResultDto = new ArrayList<>();
		listResultDto.add(dto);

		return new InputStreamResource(xlsxUtil.create(headerLiterals, nameSheet, listResultDto));
	}

	@Override
	public void downloadToExcel2(HttpServletResponse httpServletResponse) {
		List<String> headerLiterals = Arrays.asList("Ruta relativa en disco", "Nombre del archivo", "Tipo",
				"Código Línea Inversión", "UUFF", "CINI", "Type of System", "Objeto control", "Pedido Posición",
				"Importe total de la línea de inversión", "Importe por cada UUFF", "Número documento GL",
				"Fecha contabilización GL", "Año contabilización", "CIF Proveedor",
				"Importe total de la factura (sin IVA)", "Hipervínculo al fichero local para su consulta individual");

		String nameSheet = "name Sheet";

		ResultDto dto = new ResultDto();
		dto.setName("Name");
		dto.setType("Type");
		dto.setHyperlink("C:/Users/acalles.ext/Desktop/ENEL/logback.xml");
		List<ResultDto> listResultDto = new ArrayList<>();
		listResultDto.add(dto);

		XlsxUtil2 xlsxUtil2 = new XlsxUtil2(headerLiterals, nameSheet, listResultDto);
		try {
			xlsxUtil2.export(httpServletResponse);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}