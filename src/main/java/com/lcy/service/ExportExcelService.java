package com.lcy.service;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public interface ExportExcelService<T> {

	/**
	 * 导出excel
	 * @param dataList
	 */
	HSSFWorkbook exportExcel(List<T> dataList);
}
