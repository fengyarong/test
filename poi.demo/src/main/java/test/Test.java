package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Util.DateUtil;
import poi.dto.WorkTimeDto;

public class Test {

	public static void main(String[] args) throws Exception {
//		String addresspath = "D:\\project_doc\\project-docs\\3_项目跟踪与监控\\50_个人周报\\2019\\201903\\20190308";
		String addresspath="D:/Config";
		Map<String, WorkTimeDto> map = new HashMap<String, WorkTimeDto>();
		File file = new File(addresspath);
		File[] listFiles = file.listFiles();
		for (File myfile : listFiles) {
			if (myfile.isFile()) {
				String path = myfile.getPath();
				String name = path.substring(path.lastIndexOf("-")+1,path.lastIndexOf("."));
				InputStream is = new FileInputStream(path);
				WorkTimeDto workDto =new WorkTimeDto();
				XSSFWorkbook hw = new XSSFWorkbook(is);
				Sheet sheet = hw.getSheetAt(3);
				long num = DateUtil.getDaysFromStar()+4;
				for (long i = num-7; i <num; i++) {
					int ca=(int) (num-i);
					Row row = sheet.getRow((int)i);
					Cell cell = row.getCell(1);
					switch (ca) {
					case 7:
						workDto.setSunday(getCellValue(cell));
						break;
					case 6:
						workDto.setSaturday(getCellValue(cell));
						break;
					case 5:
						workDto.setFriday(getCellValue(cell));
						break;
					case 4:
						workDto.setThursday(getCellValue(cell));	
						break;
					case 3:
						workDto.setWednesday(getCellValue(cell));		
						break;
					case 2:
						workDto.setTuesday(getCellValue(cell));		
						break;
					case 1:
						workDto.setMonday(getCellValue(cell));			
						break;	
					}
				}
				map.put(name, workDto);
				is.close();
			}
		}
		System.out.println(map);
		
	}
	
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
