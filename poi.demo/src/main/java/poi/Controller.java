package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ENUM.ENUM_CATEGORY;

public class Controller {

	public void method_A(String inPath, String ouPath) throws Exception {
		HashSet<XlsDto> list_last = new HashSet<XlsDto>();
		HashSet<XlsDto> list_next = new HashSet<XlsDto>();
		Map<String,WorkTimeDto> map =new HashMap<String,WorkTimeDto>();
		File file = new File(inPath);
		File[] listFiles = file.listFiles();
		for (File myfile : listFiles) {
			if (myfile.isFile()) {
				method_B(myfile.getPath(), list_last, ENUM_CATEGORY.LASTWEEK.code());
				method_B_1(myfile.getPath(), list_next, ENUM_CATEGORY.NEXTWEEK.code());
//				method_B_2(myfile.getPath(),map, ENUM_CATEGORY.COST.code());
			}
		}
		// method_C(list, ouPath);
		method_D(list_last, list_next, ouPath);
	}

	private void method_B_2(String path, Map<String,WorkTimeDto> map, Integer code) throws Exception {
		String name =path.substring(path.lastIndexOf("-"));
		InputStream is = new FileInputStream(path);
	}

	private HashSet<XlsDto> method_B(String path, HashSet<XlsDto> list, int num) throws Exception {
		InputStream is = new FileInputStream(path);
		XSSFWorkbook hw = new XSSFWorkbook(is);
		Sheet sheet = hw.getSheetAt(num);// 查詢第二頁
		for (int i = 2; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			XlsDto xls = new XlsDto();
			for (int j = 1; j <= 8; j++) {
				Cell cell = row.getCell(j);
				if(cell==null)
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
		is.close();
		return list;
	}

	private  HashSet<XlsDto> method_B_1(String path, HashSet<XlsDto> list, int num) throws Exception {
		InputStream is = new FileInputStream(path);
		XSSFWorkbook hw = new XSSFWorkbook(is);
		Sheet sheet = hw.getSheetAt(num);// 查詢第二頁
		for (int i = 2; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
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
		is.close();
		return list;
	}

	/*private  void method_C(HashSet<XlsDto> list_last, HashSet<XlsDto> list_next, String path) throws Exception {
		FileInputStream fs = new FileInputStream(path);
		POIFSFileSystem ps = new POIFSFileSystem(fs);
		HSSFWorkbook wb = new HSSFWorkbook(ps);
		HSSFSheet sheet = wb.getSheetAt(0);
		HSSFRow row = sheet.getRow(0);
		int lastRowNum = sheet.getLastRowNum();
		FileOutputStream out = new FileOutputStream(path);

		for (XlsDto xlsDto : list_last) {
			row = sheet.createRow(lastRowNum++);
			row.createCell(1).setCellValue(xlsDto.getType());
			row.createCell(2).setCellValue(xlsDto.getName());
			row.createCell(3).setCellValue(xlsDto.getStardate());
			row.createCell(4).setCellValue(xlsDto.getEnddate());
			row.createCell(5).setCellValue(xlsDto.getDays());
			row.createCell(6).setCellValue(xlsDto.getDutyperson());
			row.createCell(7).setCellValue(xlsDto.getLv());
			row.createCell(8).setCellValue(xlsDto.getDesc());
		}
		out.flush();
		wb.write(out);
		out.close();
	}*/

	private void method_D(HashSet<XlsDto> list_last, HashSet<XlsDto> list_next, String path) throws Exception {
		FileOutputStream fos = new FileOutputStream(path);
		Workbook wb = new HSSFWorkbook();
		Sheet xssfsheet = wb.createSheet(ENUM_CATEGORY.LASTWEEK.desc());
		Sheet xssfsheet1 = wb.createSheet(ENUM_CATEGORY.NEXTWEEK.desc());
		method_E(list_last, wb, xssfsheet);
		method_E_1(list_next, wb, xssfsheet1);
		wb.write(fos);
		fos.close();
	}

	private void method_E(HashSet<XlsDto> list, Workbook wb, Sheet xssfsheet) {
		int i = 0;
		for (XlsDto xlsDto : list) {
			System.out.println(xlsDto);
			Row row = xssfsheet.createRow(i++);
			for (int n = 1; n < 9; n++) {
				Cell cell = row.createCell(n);
				Font font = wb.createFont();
				font.setFontName("宋体");// 設置字體
				font.setFontHeightInPoints((short) 10);// 字號大小
				CellStyle wrapStyle = wb.createCellStyle();
				wrapStyle.setFont(font);
				switch (n) {
				case 1:
					cell.setCellValue(xlsDto.getType());
					break;
				case 2:
					cell.setCellValue(xlsDto.getName());
					break;
				case 3:
					cell.setCellValue(xlsDto.getStardate() == null ? new Date() : xlsDto.getStardate());
					DataFormat format3 = wb.createDataFormat();
					wrapStyle.setDataFormat(format3.getFormat("yyyy-m-d"));
					cell.setCellStyle(wrapStyle);
					break;
				case 4:
					cell.setCellValue(xlsDto.getStardate() == null ? new Date() : xlsDto.getStardate());
					DataFormat format4 = wb.createDataFormat();
					wrapStyle.setDataFormat(format4.getFormat("yyyy-m-d"));
					cell.setCellStyle(wrapStyle);
					break;
				case 5:
					cell.setCellValue(xlsDto.getDays());
					break;
				case 6:
					cell.setCellValue(xlsDto.getDutyperson());
					break;
				case 7:
					cell.setCellValue(xlsDto.getLv());
					wrapStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00%"));
					cell.setCellStyle(wrapStyle);
					break;
				case 8:
					cell.setCellValue(xlsDto.getDesc());
					break;
				}
			}
		}
	}

	private void method_E_1(HashSet<XlsDto> list, Workbook wb, Sheet xssfsheet) {
		int i = 0;
		for (XlsDto xlsDto : list) {
			System.out.println(xlsDto);
			Row row = xssfsheet.createRow(i++);
			for (int n = 1; n < 9; n++) {
				Cell cell = row.createCell(n);
				Font font = wb.createFont();
				font.setFontName("宋体");// 設置字體
				font.setFontHeightInPoints((short) 10);// 字號大小
				CellStyle wrapStyle = wb.createCellStyle();
				wrapStyle.setFont(font);
				switch (n) {
				case 1:
					cell.setCellValue(xlsDto.getType());
					break;
				case 2:
					cell.setCellValue(xlsDto.getName());
					break;
				case 3:
					cell.setCellValue(xlsDto.getStardate() == null ? new Date() : xlsDto.getStardate());
					DataFormat format3 = wb.createDataFormat();
					wrapStyle.setDataFormat(format3.getFormat("yyyy-m-d"));
					cell.setCellStyle(wrapStyle);
					break;
				case 4:
					cell.setCellValue(xlsDto.getStardate() == null ? new Date() : xlsDto.getStardate());
					DataFormat format4 = wb.createDataFormat();
					wrapStyle.setDataFormat(format4.getFormat("yyyy-m-d"));
					cell.setCellStyle(wrapStyle);
					break;
				case 5:
					cell.setCellValue(xlsDto.getDays());
					break;
				case 6:
					cell.setCellValue(xlsDto.getDutyperson());
					break;
				}
			}
		}
	}

	public String getCellValue(Cell cell) {
		if(cell==null){
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
