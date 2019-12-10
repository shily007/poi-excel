package com.dy.excel;

import java.io.Serializable;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * @Title ExcelCell
 * @Description  excel单元格
 * @author dy
 * @date 2019年3月22日
 */
@SuppressWarnings("serial")
@Data
@Accessors(chain=true)
public class ExcelCell implements Serializable {
	
	//字体颜色	HSSFColor.BLACK.index
	private Short fontColor;
	//字体大小
	private Short fontHeightInPoints;
	//填充方案	FillPatternType.SOLID_FOREGROUND
	private FillPatternType fillPattern = FillPatternType.SOLID_FOREGROUND;//用于xlsx
	private Short fillPattern2 = HSSFCellStyle.SOLID_FOREGROUND;//用于xls
	//背景颜色	HSSFColor.BLACK.index
	private Short bgColor = HSSFColor.WHITE.index;
	//水平居中	HSSFCellStyle.ALIGN_CENTER
	private Short alignment;
	//垂直居中	XSSFCellStyle.VERTICAL_CENTER
	private Short verticalAlignment = XSSFCellStyle.VERTICAL_CENTER;
	//字体加粗	HSSFFont.BOLDWEIGHT_BOLD
	private Short boldweight;
	//是否换行
	private boolean wrapText = false;
	//单元格内容
	private Object text;
	//行高
	private Short height;
	//列宽
	private Short width;
	//单元格格式	XSSFCell.CELL_TYPE_STRING
	private Integer cellType;
	//行数
	private Integer rowNum;
	//列数
	private Integer colNum;
	//线框颜色
	private short borderTop = CellStyle.BORDER_THIN;
	private short borderRight = CellStyle.BORDER_THIN;
	private short borderBottom = CellStyle.BORDER_THIN;
	private short borderLeft = CellStyle.BORDER_THIN;
	//图片路径	以下x1,x2,y1,y2的值都是调整OK的图片刚好再单元格内露出边框，可以根据实际需要进行调整（对xlsx好像不起作用）
	private String imgPath;
	private short x1 = 10;	
	private short x2 = 0;
	private short y1 = 3;
	private short y2 = 0;

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		ExcelCell other = (ExcelCell) obj;
		if (alignment == null) {
			if (other.alignment != null)
				return false;
		} else if (!alignment.equals(other.alignment))
			return false;
		if (bgColor == null) {
			if (other.bgColor != null)
				return false;
		} else if (!bgColor.equals(other.bgColor))
			return false;
		if (boldweight == null) {
			if (other.boldweight != null)
				return false;
		} else if (!boldweight.equals(other.boldweight))
			return false;
		if (borderBottom != other.borderBottom)
			return false;
		if (borderLeft != other.borderLeft)
			return false;
		if (borderRight != other.borderRight)
			return false;
		if (borderTop != other.borderTop)
			return false;
		if (cellType == null) {
			if (other.cellType != null)
				return false;
		} else if (!cellType.equals(other.cellType))
			return false;
		if (colNum == null) {
			if (other.colNum != null)
				return false;
		} else if (!colNum.equals(other.colNum))
			return false;
		if (fillPattern != other.fillPattern)
			return false;
		if (fillPattern2 == null) {
			if (other.fillPattern2 != null)
				return false;
		} else if (!fillPattern2.equals(other.fillPattern2))
			return false;
		if (fontColor == null) {
			if (other.fontColor != null)
				return false;
		} else if (!fontColor.equals(other.fontColor))
			return false;
		if (fontHeightInPoints == null) {
			if (other.fontHeightInPoints != null)
				return false;
		} else if (!fontHeightInPoints.equals(other.fontHeightInPoints))
			return false;
		if (rowNum == null) {
			if (other.rowNum != null)
				return false;
		} else if (!rowNum.equals(other.rowNum))
			return false;
		if (verticalAlignment == null) {
			if (other.verticalAlignment != null)
				return false;
		} else if (!verticalAlignment.equals(other.verticalAlignment))
			return false;
		if (wrapText != other.wrapText)
			return false;
		return true;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((alignment == null) ? 0 : alignment.hashCode());
		result = prime * result + ((bgColor == null) ? 0 : bgColor.hashCode());
		result = prime * result + ((boldweight == null) ? 0 : boldweight.hashCode());
		result = prime * result + borderBottom;
		result = prime * result + borderLeft;
		result = prime * result + borderRight;
		result = prime * result + borderTop;
		result = prime * result + ((cellType == null) ? 0 : cellType.hashCode());
		result = prime * result + ((colNum == null) ? 0 : colNum.hashCode());
		result = prime * result + ((fillPattern == null) ? 0 : fillPattern.hashCode());
		result = prime * result + ((fillPattern2 == null) ? 0 : fillPattern2.hashCode());
		result = prime * result + ((fontColor == null) ? 0 : fontColor.hashCode());
		result = prime * result + ((fontHeightInPoints == null) ? 0 : fontHeightInPoints.hashCode());
		result = prime * result + ((height == null) ? 0 : height.hashCode());
		result = prime * result + ((rowNum == null) ? 0 : rowNum.hashCode());
		result = prime * result + ((text == null) ? 0 : text.hashCode());
		result = prime * result + ((verticalAlignment == null) ? 0 : verticalAlignment.hashCode());
		result = prime * result + ((width == null) ? 0 : width.hashCode());
		result = prime * result + (wrapText ? 1231 : 1237);
		result = prime * result + ((imgPath == null) ? 0 : imgPath.hashCode());
		return result;
	}

}
