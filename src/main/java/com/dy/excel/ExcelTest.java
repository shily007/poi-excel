package com.dy.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class ExcelTest {
	
	public static void main(String[] args) {
		List<ExcelCell[]> headers = new ArrayList<ExcelCell[]>();
//		headers.add(new ExcelCell[] { new ExcelCell().setText("Excel测试表\r\n2019年").setAlignment(HSSFCellStyle.ALIGN_CENTER)
//				.setFontColor(HSSFColor.GREEN.index).setColNum(5).setHeight((short) 1000).setRowNum(2)
//				.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD).setWrapText(true) });
//		headers.add(new ExcelCell[] {
//				new ExcelCell().setText("").setWrapText(true).setAlignment(HSSFCellStyle.ALIGN_CENTER)});
//		headers.add(new ExcelCell[] {
//				new ExcelCell().setText("表头1行1列").setWrapText(true).setAlignment(HSSFCellStyle.ALIGN_CENTER),
//				new ExcelCell().setText("表头2行2列").setRowNum(2).setColNum(2).setAlignment(HSSFCellStyle.ALIGN_CENTER),
//				null, new ExcelCell().setText("表头1行3列").setAlignment(HSSFCellStyle.ALIGN_CENTER),
//				new ExcelCell().setText("表头1行4列").setAlignment(HSSFCellStyle.ALIGN_CENTER) });
//		headers.add(new ExcelCell[] { new ExcelCell().setText("表头2行1列").setAlignment(HSSFCellStyle.ALIGN_CENTER), null,
//				null, new ExcelCell().setText("表头2行3列").setWidth((short) 9000).setAlignment(HSSFCellStyle.ALIGN_CENTER),
//				new ExcelCell().setText("表头2行4列").setAlignment(HSSFCellStyle.ALIGN_CENTER) });
//		headers.add(new ExcelCell[] { new ExcelCell().setText("Excel测试表\r\n2019年").setAlignment(HSSFCellStyle.ALIGN_CENTER)
//				.setFontColor(HSSFColor.GREEN.index).setColNum(5).setHeight((short) 1000).setRowNum(2)
//				.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD).setWrapText(true) });
		ExcelCell[] excelCells = new ExcelCell[17];
		excelCells[0] = new ExcelCell().setRowNum(2);
		excelCells[1] = new ExcelCell().setText("星期一").setAlignment(HSSFCellStyle.ALIGN_CENTER)
				.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER).setFontColor(HSSFColor.GREEN.index).setColNum(8)
				.setHeight((short) 600).setBoldweight(HSSFFont.BOLDWEIGHT_BOLD).setWrapText(true);
		excelCells[9] = new ExcelCell().setText("星期二").setAlignment(HSSFCellStyle.ALIGN_CENTER)
				.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER).setFontColor(HSSFColor.GREEN.index).setColNum(8)
				.setHeight((short) 600).setBoldweight(HSSFFont.BOLDWEIGHT_BOLD).setWrapText(true);
		headers.add(excelCells);
		excelCells = new ExcelCell[17];
		for (int i = 0; i < 2; i++) {
			for (int j = 0; j < 8; j++) {
				String text = "";
				if (j < 4)
					text = "上午" + (j + 1);
				else
					text = "下午" + (j + 1);
				excelCells[i * 8 + j + 1] = new ExcelCell().setText(text).setAlignment(HSSFCellStyle.ALIGN_CENTER)
						.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER).setFontColor(HSSFColor.RED.index)
						.setHeight((short) 600).setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			}
		}
		headers.add(excelCells);
		List<ExcelCell[]> data = new ArrayList<ExcelCell[]>();
		for (int i = 0; i < 10; i++) {
			ExcelCell[] cells = new ExcelCell[17];
			for (int j = 0; j < 16; j++) {
				cells[j] = new ExcelCell().setText("第" + (i + 1) + "行\r\n第" + (j + 1) + "列")
						.setBgColor(HSSFColor.WHITE.index).setFillPattern(FillPatternType.SOLID_FOREGROUND)
						.setFillPattern2(HSSFCellStyle.FINE_DOTS).setWrapText(true);
				if(i == 9) 
					cells[j] = new ExcelCell().setImgPath("C:\\Users\\Administrator\\Desktop\\java架构学习\\timg.jpg")
					.setWidth((short) 3800)
					.setHeight((short) 1200);
			} 
			if (i == 0)
				cells[16] = new ExcelCell().setText("第17列").setRowNum(10).setAlignment(XSSFCellStyle.ALIGN_CENTER)
						.setFillPattern(FillPatternType.SOLID_FOREGROUND)// xlsx的填充模式
						.setFillPattern2(HSSFCellStyle.SOLID_FOREGROUND)// xls的填充模式
						.setWrapText(true);			
			data.add(cells);
		}
		String filePath = "D:/122213.xlsx";
		File file = new File(filePath);
		if (!file.exists()) {
			file.setWritable(true, false);// 获取权限
			try {
				file.createNewFile();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		try {
			OutputStream os = new FileOutputStream(file);
			ExcelUtil.write(headers, data, filePath, null, null, os);
			os.close();
			System.out.println("导出完成");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
