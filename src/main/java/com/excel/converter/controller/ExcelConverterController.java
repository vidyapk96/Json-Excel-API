package com.excel.converter.controller;

import java.io.IOException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.excel.converter.dto.JsonToExcelDto;
import com.excel.converter.service.ExcelConverterService;
/**
 * 
 * @author Vidya
 *
 */
@RestController
@RequestMapping("/excel-converter")
public class ExcelConverterController {
	private final Logger log = LoggerFactory.getLogger(ExcelConverterController.class);

	@Autowired
	private ExcelConverterService excelConverterService;

	@PostMapping
	public ResponseEntity<String> jsonToExcelConverter(@RequestBody JsonToExcelDto jsonToExcelDto) throws IOException {

		log.info("Creating excel file");
		excelConverterService.jsonToExcelConverter(jsonToExcelDto);
		return ResponseEntity.ok().body("SuccessFully created");
	}

}
