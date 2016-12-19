package com.lcy.service.impl;

import java.util.List;

import org.apache.poi.hssf.record.formula.functions.T;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Service;

import com.lcy.service.ExportExcelService;
import com.lcy.util.ExcelUtil;

@Service
public class ExportExcelServiceImpl implements ExportExcelService<T> {

	private static final String SHEET_NAME = "data";
	
	/**
	 * 将表数据dataList封装到excel
	 */
	public HSSFWorkbook exportExcel(List<T> dataList) {
		HSSFWorkbook workbook = new HSSFWorkbook();// 创建工作簿
		HSSFSheet sheet = ExcelUtil.createSheet(workbook, SHEET_NAME);// 创建工作表
		//填充数据
		ExcelUtil.setDataFromDTO(workbook, sheet, dataList);
		
		return workbook;
	}
}
