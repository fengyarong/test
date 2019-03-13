package Util;

import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class SaveUtil {
	
	public void SaveMethod(String object, XSSFCellStyle xssfCellStyle, XSSFCell xssfCell) {
		if("".equals(object)){
			xssfCell.setCellValue("客户评价：满意");
		}else if("1".equals(object)){
			xssfCell.setCellValue("");
		}else{
			xssfCell.setCellValue(object);
		}
		xssfCell.setCellStyle(xssfCellStyle);
	}
	
	public void SaveMethod(Date object, XSSFCellStyle xssfCellStyle, XSSFCell xssfCell) {
		xssfCell.setCellValue(object);
		XSSFCellStyle myxssf = (XSSFCellStyle) xssfCellStyle.clone();
		if(DateUtil.jdugeBeforeDate(object)){
			//添加背景色为红色
			myxssf.setFillForegroundColor(IndexedColors.RED.getIndex());
			myxssf.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		xssfCell.setCellStyle(myxssf);
	}
	
	public void SaveMethod(double object, XSSFCellStyle xssfCellStyle, XSSFCell xssfCell) {
		xssfCell.setCellValue(object);
		xssfCell.setCellStyle(xssfCellStyle);
	}
	
	public void SaveMethod(int object, XSSFCellStyle xssfCellStyle, XSSFCell xssfCell) {
		xssfCell.setCellValue(object);
		xssfCell.setCellStyle(xssfCellStyle);
	}
}
