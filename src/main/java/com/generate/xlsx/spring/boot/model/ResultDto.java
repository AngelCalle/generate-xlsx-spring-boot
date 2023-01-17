package com.generate.xlsx.spring.boot.model;

import java.io.Serializable;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class ResultDto implements Serializable{

	private static final long serialVersionUID = -3200111502428935502L;
	
	private String name;
	private String type;
	private String hyperlink;
	
}
