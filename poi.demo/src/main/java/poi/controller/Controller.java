package poi.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ENUM.ENUM_CATEGORY;
import Util.CellUtil;
import Util.SaveUtil;
import Util.StringUtil;
import poi.dto.CellTypeDto;
import poi.dto.WorkTimeDto;
import poi.dto.XlsDto;

public class Controller {

	public void method_A(String inPath, String ouPath) throws Exception {
		LinkedHashSet<XlsDto> list_last = new LinkedHashSet<XlsDto>();
		LinkedHashSet<XlsDto> list_next = new LinkedHashSet<XlsDto>();
//		Map<String, WorkTimeDto> map = new HashMap<String, WorkTimeDto>();
		File file = new File(inPath);
		File[] listFiles = file.listFiles();
		for (File myfile : listFiles) {
			if (myfile.isFile()) {
				method_B(myfile.getPath(), list_last, ENUM_CATEGORY.LASTWEEK.code());
				method_B_1(myfile.getPath(), list_next, ENUM_CATEGORY.NEXTWEEK.code());
				// method_B_2(myfile.getPath(),map, ENUM_CATEGORY.COST.code());
			}
		}
		method_C(list_last, ouPath,ENUM_CATEGORY.LASTWEEK.code());
		method_C_1(list_next, ouPath,ENUM_CATEGORY.NEXTWEEK.code());
//		method_D(list_last, list_next, ouPath);
	}

	private void method_B_2(String path, Map<String, WorkTimeDto> map, Integer code) throws Exception {
		String name = path.substring(path.lastIndexOf("-"));
		InputStream is = new FileInputStream(path);
	}

	private LinkedHashSet<XlsDto> method_B(String path, LinkedHashSet<XlsDto> list, int num) throws Exception {
		InputStream is = new FileInputStream(path);
		XSSFWorkbook hw = new XSSFWorkbook(is);
		Sheet sheet = hw.getSheetAt(num);
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
						xls.setDays(new Double(cell.getNumericCellValue()).intValue());
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
		return list;
	}

	private LinkedHashSet<XlsDto> method_B_1(String path, LinkedHashSet<XlsDto> list, int num) throws Exception {
		InputStream is = new FileInputStream(path);
		XSSFWorkbook hw = new XSSFWorkbook(is);
		Sheet sheet = hw.getSheetAt(num);// 查詢第二頁
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
						xls.setDays(new Double(cell.getNumericCellValue()).intValue());
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
		return list;
	}

	private void method_C(HashSet<XlsDto> list, String ouPath, Integer integer) throws Exception {
		CellUtil cellUtil =new CellUtil();
		SaveUtil saveUtil =new SaveUtil();
		
		FileInputStream fs = new FileInputStream(ouPath);
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		XSSFSheet sheet = wb.getSheetAt(integer);
		CellTypeDto tpye = cellUtil.getTpye(sheet.getRow(3));
		XSSFRow row = sheet.getRow(0);
		int lastRowNum = sheet.getLastRowNum();
		FileOutputStream out = new FileOutputStream(ouPath);
		for(XlsDto xlsDto:list){
			row = sheet.createRow(++lastRowNum);
			saveUtil.SaveMethod("1", tpye.getXssf0(), row.createCell(0));
			saveUtil.SaveMethod(xlsDto.getType(), tpye.getXssf1(), row.createCell(1));
			saveUtil.SaveMethod(xlsDto.getName(), tpye.getXssf2(), row.createCell(2));
			saveUtil.SaveMethod(xlsDto.getStardate(), tpye.getXssf3(), row.createCell(3));
			saveUtil.SaveMethod(xlsDto.getEnddate(), tpye.getXssf4(), row.createCell(4));
			saveUtil.SaveMethod(xlsDto.getDays(), tpye.getXssf5(), row.createCell(5));
			saveUtil.SaveMethod(xlsDto.getDutyperson(), tpye.getXssf6(), row.createCell(6));
			saveUtil.SaveMethod(xlsDto.getLv(), tpye.getXssf7(), row.createCell(7));
			saveUtil.SaveMethod(xlsDto.getDesc(), tpye.getXssf8(), row.createCell(8));
		}
		out.flush();
		wb.write(out);
		out.close();
		fs.close();
		System.out.println("Last写入完毕");
	}
	
	private void method_C_1(HashSet<XlsDto> list, String ouPath, Integer integer) throws Exception {
		CellUtil cellUtil =new CellUtil();
		SaveUtil saveUtil =new SaveUtil();
		
		FileInputStream fs = new FileInputStream(ouPath);
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		XSSFSheet sheet = wb.getSheetAt(integer);
		CellTypeDto tpye = cellUtil.getTpye(sheet.getRow(3));
		XSSFRow row = sheet.getRow(0);
		int lastRowNum = sheet.getLastRowNum();
		FileOutputStream out = new FileOutputStream(ouPath);
		for(XlsDto xlsDto:list){
			row = sheet.createRow(++lastRowNum);
			saveUtil.SaveMethod("1", tpye.getXssf0(), row.createCell(0));
			saveUtil.SaveMethod(xlsDto.getType(), tpye.getXssf1(), row.createCell(1));
			saveUtil.SaveMethod(xlsDto.getName(), tpye.getXssf2(), row.createCell(2));
			saveUtil.SaveMethod(xlsDto.getStardate(), tpye.getXssf3(), row.createCell(3));
			saveUtil.SaveMethod(xlsDto.getEnddate(), tpye.getXssf4(), row.createCell(4));
			saveUtil.SaveMethod(xlsDto.getDays(), tpye.getXssf5(), row.createCell(5));
			saveUtil.SaveMethod(xlsDto.getDutyperson(), tpye.getXssf6(), row.createCell(6));
			saveUtil.SaveMethod("1", tpye.getXssf7(), row.createCell(7));
		}
		out.flush();
		wb.write(out);
		out.close();
		fs.close();
		System.out.println("Next写入完毕");
	}
	
}
