package com.dy.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.springframework.web.multipart.MultipartFile;

/**
 * excel表格适配器
 * 
 * @author dy
 *
 */
public class ExcelUtil {

	/**
	 * @method export
	 * @Description 导出excel
	 * @param data     要导出的数据
	 * @param fileName 文件名称
	 * @param rowNum   每页行数
	 * @param request
	 * @param response
	 * @author dy
	 */
	public static void export(List<ExcelCell[]> data, String fileName, Integer rowNum, HttpServletRequest request,
			HttpServletResponse response) {
		if (StringUtils.isNotBlank(fileName))
			fileName = LocalDateTime.now().getSecond() + ".xls";
		try {
			String userAgent = request.getHeader("User-Agent");
			// 针对IE或者以IE为内核的浏览器：
			if (userAgent.contains("MSIE") || userAgent.contains("Trident"))
				fileName = java.net.URLEncoder.encode(fileName, "UTF-8");
			else // 非IE浏览器的处理：
				fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");
			// 获取输出流
			OutputStream os = response.getOutputStream();
			// 设置导出Excel报表的导出形式
			response.setContentType("application/x-excel");
			response.setCharacterEncoding("GBK");
			response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
			write(null, data, fileName, null, null, os);
			// 刷新输出流、关闭输出流
			os.flush();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 读取excel表格
	 * @param file      上传的文件
	 * @param sheetName 要读取的工作簿
	 * @param columns   读取的列
	 * @param row       从第几行开始读
	 * @return
	 * @throws IOException
	 */
	ArrayList<String[]> read(MultipartFile file, String sheetName, Integer columns, Integer row) throws IOException {
		if (file.getOriginalFilename().endsWith(".xls"))
			return ExcelXls.read(file.getInputStream(), sheetName, columns, row);
		else
			return ExcelXlsx.read(file.getInputStream(), sheetName, columns, row);
	};

	/**
	 * 导出excel表格
	 * @param headers   表头
	 * @param data      数据
	 * @param sheetName 工作簿名称
	 * @param rowNum    每一个工作簿的行数
	 * @param os        输出流
	 */
	static void write(List<ExcelCell[]> headers, List<ExcelCell[]> data, String fileName, String sheetName,
			Integer rowNum, OutputStream os) {
		if (StringUtils.isNotBlank(fileName)) {
			if (fileName.endsWith(".xls"))
				ExcelXls.write(headers, data, sheetName, rowNum, os);
			else
				ExcelXlsx.write(headers, data, sheetName, rowNum, os);
		}
	}

}
