package com.dy.excel;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @Title ExcelHelp
 * @Description Excel表格read or write（本方法适用于Excel表格97/2003/2007/2010版本）
 * @author dy
 * @date 2016年11月1日
 */
public class ExcelXlsx {

	/**
	 * @method readXlsx(InputStream in,String sheetName,Integer column)
	 * @Description 读取Excel表格2007/2010版本
	 * @param in        Excel表格二进制输入流
	 * @param sheetName 需要读取的工作薄名称
	 * @param column    需要读取的列数
	 * @return
	 * @throws IOException
	 * @author dy
	 */
	public static ArrayList<String[]> read(InputStream in, String sheetName, Integer columns, Integer row)
			throws IOException {

		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(in);
		ArrayList<String[]> list = new ArrayList<String[]>();
		// Read the Sheet 1
		XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
		if (row == null || row < 0)
			row = 1;
		// Read the Row 2
		for (int rowNum = row; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
			XSSFRow xssfRow = xssfSheet.getRow(rowNum);
			String[] strs = new String[columns];
			if (xssfRow != null) {
				for (int i = 0; i < columns; i++) {
					if (xssfRow.getCell(i) != null && (xssfRow.getCell(i)).toString() != null) {
						xssfRow.getCell(i).setCellType(XSSFCell.CELL_TYPE_STRING);
						strs[i] = xssfRow.getCell(i).toString();
						if (strs[i] != null)
							strs[i] = strs[i].trim();
					}
				}
				if (strs != null) {
					list.add(strs);
				}
			}
		}
//        xssfWorkbook.close();
		return list;
	}

	/**
	 * @method write
	 * @Description 导出xlsx类型的表格
	 * @param headers   表头
	 * @param data      表内容
	 * @param sheetName 工作簿名称
	 * @param rowNum    每个工作簿的行数
	 * @param os        输出流
	 * @author dy
	 */
	public static void write(List<ExcelCell[]> headers, List<ExcelCell[]> data, String sheetName, Integer rowNum,
			OutputStream os) {
		XSSFWorkbook workBook = new XSSFWorkbook();
		if (data != null && data.size() > 0) {
			int totalRow = data.size();// 总行数
			int sheetNum = 1;// 工作薄的数量
			if (rowNum == null || rowNum > 1000000 || rowNum <= 0)
				rowNum = 1000000;
			if (totalRow % rowNum == 0)
				sheetNum = totalRow / rowNum;
			else
				sheetNum = totalRow / rowNum + 1;
			try {
				for (int i = 0; i < sheetNum; i++) {
					XSSFSheet sheet;
					List<ExcelCell[]> list = new ArrayList<ExcelCell[]>();
					int last = (i + 1) * rowNum;
					if (last > data.size())
						last = data.size();
					for (int j = i * rowNum; j < last; j++) {
						list.add(data.get(j));
					}
					if (sheetName == null) {
						if (i == 0)
							sheet = workBook.createSheet("Sheet");
						else
							sheet = workBook.createSheet("Sheet" + i);
					} else {
						if (i == 0)
							sheet = workBook.createSheet(sheetName);
						else
							sheet = workBook.createSheet(sheetName + i);
					}
					if (headers != null && headers.size() > 0)
						sheet = writeDataInXSSFSheet(sheet, workBook, headers, 0);
					if (data != null && data.size() > 0) {
						if (headers != null)
							sheet = writeDataInXSSFSheet(sheet, workBook, list, headers.size());
						else
							sheet = writeDataInXSSFSheet(sheet, workBook, list, 0);
					}
				}
				workBook.write(os);
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				try {
					os.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * @method writeDataInXSSFSheet
	 * @Description 将数据写入工作薄
	 * @param sheet    工作簿
	 * @param workBook
	 * @param data     要插入的数据
	 * @param start    开始行数
	 * @return
	 * @author dy
	 */
	private static XSSFSheet writeDataInXSSFSheet(XSSFSheet sheet, XSSFWorkbook workBook, List<ExcelCell[]> data,
			int start) {
		if (sheet != null && workBook != null && data != null && data.size() > 0) {
			XSSFCellStyle cellStyle = null;
			ExcelCell lastEcell = null;
			for (int j = 0; j < data.size(); j++) {
				XSSFRow row = sheet.getRow(start);
				if (row == null)
					row = sheet.createRow(start);
				short height = 0;
				for (int k = 0; k < data.get(j).length; k++) {
					ExcelCell ecell = data.get(j)[k];
					if (ecell != null) {
						if (ecell.getHeight() != null && ecell.getHeight() > height)
							height = ecell.getHeight();// 获取单元格高度以最高为准
						if (ecell.getWidth() != null && ecell.getWidth() > 0)
							sheet.setColumnWidth(k, ecell.getWidth());// 设置列宽
						if (!ecell.equals(lastEcell)) {
							// 创建样式(使用工作本的对象创建)
							cellStyle = workBook.createCellStyle();
							// 设置线框颜色
							cellStyle.setBorderTop(ecell.getBorderTop());
							cellStyle.setBorderRight(ecell.getBorderRight());
							cellStyle.setBorderBottom(ecell.getBorderBottom());
							cellStyle.setBorderLeft(ecell.getBorderLeft());
							// 创建字体的对象
							XSSFFont font = workBook.createFont();
							if (ecell.getFontColor() != null)
								font.setColor(ecell.getFontColor());// 设置字体的颜色
							if (ecell.getFontHeightInPoints() != null)
								font.setFontHeightInPoints(ecell.getFontHeightInPoints());// 字体大小
							if (ecell.getBoldweight() != null)
								font.setBoldweight(ecell.getBoldweight());// 将字体加粗
							if (ecell.getFillPattern() != null)
								cellStyle.setFillPattern(ecell.getFillPattern());// 填充方案
							if (ecell.getBgColor() != null)
								cellStyle.setFillForegroundColor(ecell.getBgColor());// 设置背景色
							if (ecell.getAlignment() != null)
								cellStyle.setAlignment(ecell.getAlignment());// 水平居中
							if (ecell.getVerticalAlignment() != null)
								cellStyle.setVerticalAlignment(ecell.getVerticalAlignment());// 垂直居中
							// 将新设置的字体属性放置到样式中
							cellStyle.setFont(font);
							cellStyle.setWrapText(ecell.isWrapText());// 是否自动换行
						}
						int firstCol = k;
						int lastCol = k;
						if (ecell.getColNum() != null && ecell.getColNum() > 0) {
							lastCol += ecell.getColNum() - 1;
						}
						int firstRow = start;
						int lastRow = start;
						if (ecell.getRowNum() != null && ecell.getRowNum() > 0) {
							lastRow += ecell.getRowNum() - 1;
						}
						if (lastCol != firstCol || lastRow != firstRow) {
							CellRangeAddress cellRange = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
							sheet.addMergedRegion(cellRange);
							setRegionBorder(ecell, cellRange, sheet, workBook);
						}
						XSSFCell cell = row.getCell(k);
						if (cell == null)
							cell = row.createCell(k);
						if (ecell.getCellType() != null)
							cell.setCellType(ecell.getCellType());// 单元格格式
						if (ecell.getText() != null) {
							cell.setCellValue(new XSSFRichTextString(ecell.getText().toString()));// 单元格内容
							if (ecell.getText().toString().indexOf("\r") > -1) {
								String[] strs = ecell.getText().toString().split("\r");
								if (strs.length > 0)
									height = (short) (row.getHeight() * strs.length);
							}
						}
						if(StringUtils.isNotBlank(ecell.getImgPath())) {
							try {
								writeImg(workBook, sheet, start, k, ecell);
							} catch (IOException e) {
								e.printStackTrace();
							}
						}
						if (cellStyle != null)
							cell.setCellStyle(cellStyle);
					}
				}
				if (height > 0)
					row.setHeight(height);
				start++;
			}
		}
		return sheet;
	}

	/**
	 * @method setRegionBorder
	 * @Description 设置合并表格的边框
	 * @param ecell
	 * @param region
	 * @param sheet
	 * @param wb
	 * @author dy
	 */
	private static void setRegionBorder(ExcelCell ecell, CellRangeAddress region, Sheet sheet, Workbook wb) {
		RegionUtil.setBorderBottom(ecell.getBorderBottom(), region, sheet, wb);
		RegionUtil.setBorderLeft(ecell.getBorderLeft(), region, sheet, wb);
		RegionUtil.setBorderRight(ecell.getBorderRight(), region, sheet, wb);
		RegionUtil.setBorderTop(ecell.getBorderTop(), region, sheet, wb);
	}
	
	/**
	 * 向excell中插入图片
	 * @param workBook
	 * @param sheet
	 * @param row 开始行
	 * @param k 开始列
	 * @param cell
	 * @throws IOException
	 */
	private static void writeImg(XSSFWorkbook workBook, XSSFSheet sheet, Integer row, int k, ExcelCell cell)
			throws IOException {
		XSSFDrawing patriarch = sheet.createDrawingPatriarch();
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		// anchor主要用于设置图片的属性
		if(cell.getRowNum() == null)
			cell.setRowNum(1);
		if(cell.getColNum() == null)
			cell.setColNum(1);
		XSSFClientAnchor anchor = new XSSFClientAnchor(cell.getX1(), cell.getY1(), cell.getX2(), cell.getY2(), (short) k, row,
				Short.valueOf((k+cell.getColNum())+""), row+cell.getRowNum());
		BufferedImage bufferImg = ImageIO.read(new File(cell.getImgPath()));
		ImageIO.write(bufferImg, cell.getImgPath().substring(cell.getImgPath().lastIndexOf(".") + 1), byteArrayOut);
		// 画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
		anchor.setAnchorType(3);
		// 插入图片
		patriarch.createPicture(anchor, workBook.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
	}

}