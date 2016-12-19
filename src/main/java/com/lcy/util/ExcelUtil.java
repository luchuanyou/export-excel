package com.lcy.util;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelUtil {

	/**
	 * 创建HSSFSheet工作簿
	 * @param wb
	 * @param sheetName
	 * @return
	 */
	public static HSSFSheet createSheet(HSSFWorkbook wb, String sheetName) {
		HSSFSheet sheet = wb.createSheet(sheetName);
//		sheet.setDefaultColumnWidth(12);
//		sheet.setGridsPrinted(false);
//		sheet.setDisplayGridlines(false);
		return sheet;
	}
	/**
	 * 创建HSSFRow
	 * @param sheet
	 * @param rowNum
	 * @param height
	 * @return
	 */
	public static HSSFRow createRow(HSSFSheet sheet,int rowNum){
        HSSFRow row=sheet.createRow(rowNum);
//        row.setHeight((short)height);
        return row;
    }
	/**
	 * 创建CELL
	 * @param row
	 * @param cellNum
	 * @return
	 */
	public static HSSFCell createCell0(HSSFRow row,int cellNum){
        HSSFCell cell=row.createCell(cellNum);
         return cell;
    }
	/**
     * 创建CELL
     * @param     row        HSSFRow    
     * @param     cellNum    int
     * @param     style    HSSFStyle
     * @return    HSSFCell
     */
    public static HSSFCell createCell(HSSFRow row,int cellNum,CellStyle style){
        HSSFCell cell=row.createCell(cellNum);
        cell.setCellStyle(style);
        return cell;
    }
    /**
     * 遍历数据并写入excel表格中
     * @param sheet
     * @param dataList
     */
    public static <T> void setDataFromDTO(HSSFWorkbook workbook, HSSFSheet sheet, List<T> dataList){
    	if(null == dataList || dataList.size() == 0)
    		return;
    	
    	//创建标题行
    	HSSFRow titleRow = ExcelUtil.createRow(sheet, 0);// 创建新行
    	Iterator<T> labIt = dataList.iterator();
		int rowNum =1;
		while (labIt.hasNext()) {
			HSSFRow row = ExcelUtil.createRow(sheet, rowNum);// 创建新行
			T t = labIt.next();
			Field[] fields = t.getClass().getDeclaredFields();
			for (short i = 0; i < fields.length; i++) {
				HSSFCell cell = ExcelUtil.createCell0(row, i);// 创建新列
				Field field = fields[i];
                String fieldName = field.getName();
                //System.out.println(fieldName);
                
                String getMethodName = "get"
                        + fieldName.substring(0, 1).toUpperCase()
                        + fieldName.substring(1);
                    Class tCls = t.getClass();
                    try {
						Method getMethod = tCls.getMethod(getMethodName,
						        new Class[] {});
						Object val = getMethod.invoke(t, new Object[] {});
						if("createTime".equals(fieldName) || "updateTime".equals(fieldName)){
							setCellValue(workbook, cell, val);
						}else{
							cell.setCellValue(val+"");
						}
						if(1== rowNum){
							HSSFCell titleCell = ExcelUtil.createCell0(titleRow, i);// 创建新列
							titleCell.setCellValue(fieldName);
						}
						System.out.println(fields[i].getName()+":"+val);
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} 
			}
			rowNum++;
		}
    }
	private static void setCellValue(HSSFWorkbook workbook, HSSFCell cell, Object val) {
		if(null == val)
			return;
		//添加单元格模板样式
		HSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(workbook.createDataFormat().getFormat("yy/mm/dd hh:mm:ss"));
		
		cell.setCellValue((Date)val);
		cell.setCellStyle(cellStyle);//应用模板格式
	}
}
