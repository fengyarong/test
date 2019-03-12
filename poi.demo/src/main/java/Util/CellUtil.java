package Util;

import org.apache.poi.xssf.usermodel.XSSFRow;

import poi.CellType;

public class CellUtil {
	
	public CellType getTpye(XSSFRow row){
		CellType type =new CellType();
		type.setXssf0(row.getCell(0).getCellStyle());
		type.setXssf1(row.getCell(1).getCellStyle());
		type.setXssf2(row.getCell(2).getCellStyle());
		type.setXssf3(row.getCell(3).getCellStyle());
		type.setXssf4(row.getCell(4).getCellStyle());
		type.setXssf5(row.getCell(5).getCellStyle());
		type.setXssf6(row.getCell(6).getCellStyle());
		type.setXssf7(row.getCell(7)==null?null:row.getCell(7).getCellStyle());
		type.setXssf8(row.getCell(8)==null?null:row.getCell(8).getCellStyle());
		return type;
	}
}
