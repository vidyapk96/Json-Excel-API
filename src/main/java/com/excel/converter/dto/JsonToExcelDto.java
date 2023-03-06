package com.excel.converter.dto;

import java.util.List;
import java.util.Map;

import lombok.Data;
/**
 * 
 * @author Vidya
 *
 */
@Data
public class JsonToExcelDto {
	Map<String, String> header;
	List<Map<String, String>> values;

}
