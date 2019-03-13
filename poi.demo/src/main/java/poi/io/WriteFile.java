package poi.io;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Util.CellUtil;
import Util.SaveUtil;
import Util.StringUtil;
import poi.dto.CellTypeDto;
import poi.dto.WorkTimeDto;
import poi.dto.XlsDto;

public class WriteFile {

	public static void writeSecondPage(HashSet<XlsDto> list, String ouPath, Integer integer) throws Exception {
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
		System.out.println("Second写入完毕");
	}
	
	public static void writeThirdPage(HashSet<XlsDto> list, String ouPath, Integer integer) throws Exception {
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
		System.out.println("Third写入完毕");
	}
	
	public static void writeFourthPage(String ouPath,Map<String, WorkTimeDto> map, Integer integer)throws Exception{
		
		Map<String,Integer> nameMap=new HashMap<String,Integer>();
		
		FileInputStream fs = new FileInputStream(ouPath);
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		XSSFSheet sheet = wb.getSheetAt(integer);
		XSSFRow row = sheet.getRow(1);
		int lastCellNum = row.getLastCellNum();
		for (int i = 1; i < lastCellNum; i++) {
			if(row.getCell(i)!=null){
				nameMap.put(StringUtil.getCellValue(row.getCell(i)), i);
			}
		}
		System.out.println(nameMap);
		fs.close();
		System.out.println("Fourth写入完毕");
	}
	
	public static void main(String[] args) {
		String ouPath="D:\\Config\\LI-国宝人寿-PMC-周报-20190301.xls";
		int page=3;
		try {
			writeFourthPage(ouPath,null,page);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
