package Util;

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

}
