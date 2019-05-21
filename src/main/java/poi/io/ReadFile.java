package poi.io;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.LinkedHashSet;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Util.DateUtil;
import Util.StringUtil;
import poi.dto.WorkTimeDto;
import poi.dto.XlsDto;

public class ReadFile {

	public static void readSecondPage(String path, LinkedHashSet<XlsDto> list, int page) throws Exception {
		InputStream is = new FileInputStream(path);
		XSSFWorkbook hw = new XSSFWorkbook(is);
		Sheet sheet = hw.getSheetAt(page);
		for (int i = 2; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (!"".equals(StringUtil.getCellValue(row.getCell(2)))) {
				XlsDto xls = new XlsDto();
				for (int j = 1; j <= 8; j++) {
					Cell cell = row.getCell(j);
					if (cell == null)
						continue;
					switch (j) {
					case 1:
						xls.setType(cell.getStringCellValue());
						break;
					case 2:
						xls.setName(cell.getStringCellValue());
						break;
					case 3:
						xls.setStardate(cell.getDateCellValue());
						break;
					case 4:
						xls.setEnddate(cell.getDateCellValue());
						break;
					case 5:
						xls.setDays(new Double(cell.getNumericCellValue()));
						break;
					case 6:
						xls.setDutyperson(cell.getStringCellValue());
						break;
					case 7:
						xls.setLv(cell.getNumericCellValue());
						break;
					case 8:
						xls.setDesc(cell.getStringCellValue());
						break;
					}
				}
				list.add(xls);
			}
		}
		is.close();
	}

	public static void readThirdPage(String path, LinkedHashSet<XlsDto> list, int page) throws Exception {
		InputStream is = new FileInputStream(path);
		XSSFWorkbook hw = new XSSFWorkbook(is);
		Sheet sheet = hw.getSheetAt(page);
		System.out.println("目前一共有："+sheet.getLastRowNum());
		for (int i = 2; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (!"".equals(StringUtil.getCellValue(row.getCell(6)))) {
				XlsDto xls = new XlsDto();
				for (int j = 1; j <= 6; j++) {
					Cell cell = row.getCell(j);
					if (cell == null) {
						continue;
					}
					switch (j) {
					case 1:
						xls.setType(cell.getStringCellValue());
						break;
					case 2:
						xls.setName(cell.getStringCellValue());
						break;
					case 3:
						xls.setStardate(cell.getDateCellValue());
						break;
					case 4:
						xls.setEnddate(cell.getDateCellValue());
						break;
					case 5:
						xls.setDays(new Double(cell.getNumericCellValue()));
						break;
					case 6:
						xls.setDutyperson(cell.getStringCellValue());
						break;
					}
				}
				list.add(xls);
			}
		}
		is.close();
	}

	public static void readFourthPage(String path, Map<String, WorkTimeDto> map, int page) throws Exception {
		String name = path.substring(path.lastIndexOf("-") + 1, path.lastIndexOf("."));
		InputStream is = new FileInputStream(path);
		WorkTimeDto workDto = new WorkTimeDto();
		XSSFWorkbook hw = new XSSFWorkbook(is);
		Sheet sheet = hw.getSheetAt(3);
		int num = DateUtil.getDaysFromStar() + 4;
		for (int i = num - 7; i < num; i++) {
			int ca = num - i;
			Row row = sheet.getRow((int) i);
			Cell cell = row.getCell(1);
			switch (ca) {
			case 7:
				workDto.setSunday(StringUtil.getCellValue(cell));
				break;
			case 6:
				workDto.setSaturday(StringUtil.getCellValue(cell));
				break;
			case 5:
				workDto.setFriday(StringUtil.getCellValue(cell));
				break;
			case 4:
				workDto.setThursday(StringUtil.getCellValue(cell));
				break;
			case 3:
				workDto.setWednesday(StringUtil.getCellValue(cell));
				break;
			case 2:
				workDto.setTuesday(StringUtil.getCellValue(cell));
				break;
			case 1:
				workDto.setMonday(StringUtil.getCellValue(cell));
				break;
			}
		}
		map.put(name, workDto);
		is.close();
	}
}
