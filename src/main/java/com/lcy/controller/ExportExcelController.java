package com.lcy.controller;

import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.util.List;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.lcy.dao.UserDTOMapper;
import com.lcy.model.UserDTO;
import com.lcy.service.ExportExcelService;

@Controller
@RequestMapping("/download")
public class ExportExcelController {

	@Resource
	ExportExcelService<UserDTO> exportExcelService;
	@Autowired
	UserDTOMapper userDTOMapper;
	
	/**
	 * 导出excel接口
	 * @param request
	 * @param response
	 */
	@RequestMapping(value="/exportExcel.do")
	public void exportExcel(HttpServletRequest request,HttpServletResponse response){
		try {
			String filename = "user.xls";
			List<UserDTO> dataList = userDTOMapper.selectAll();
			System.out.println("=============size===="+dataList);
			
			HSSFWorkbook workbook = exportExcelService.exportExcel(dataList );
			// 清空response  
			response.reset();  
			// 设置response的Header  
			response.addHeader("Content-Disposition", "attachment;filename=" + new String(filename.getBytes()));  
			response.setContentType("application/octet-stream");  
			
			OutputStream  fos = new BufferedOutputStream(response.getOutputStream());  
			workbook.write(fos);
			fos.flush();
			fos.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
