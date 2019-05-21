package Util;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;

public class StringUtil {

	public static String getCellValue(Cell cell) {
		if (cell == null) {
			return "";
		}
		// 如果是boolean
		if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(cell.getBooleanCellValue());
		}
		// 如果是数字类型
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			return String.valueOf(cell.getNumericCellValue());
		}
		// 如果是String类型
		return String.valueOf(cell.getStringCellValue());
	}
	
	//判断是否为数字
	public static boolean isNumeric(String str){
		try{
			Double.parseDouble(str);
		}catch(Exception e){
			return false;
		}
		return true;
	}
	
	public static void main(String[] args) {
		String str ="9";
		double parseDouble = Double.parseDouble(str);
		System.out.println(parseDouble);
		System.out.println(isNumeric(str));
	}
}
