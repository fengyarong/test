package poi.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
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
import poi.io.ReadFile;
import poi.io.WriteFile;

public class Controller {

	public void doItBegin(String inPath, String ouPath) throws Exception {
		LinkedHashSet<XlsDto> list_second = new LinkedHashSet<XlsDto>();
		LinkedHashSet<XlsDto> list_third = new LinkedHashSet<XlsDto>();
		Map<String, WorkTimeDto> map = new HashMap<String, WorkTimeDto>();

		WorkTimeDto workTimeDto =new WorkTimeDto();
		workTimeDto.setMonday("8");
		workTimeDto.setThursday("8");
		workTimeDto.setWednesday("8");
		workTimeDto.setTuesday("8");
		workTimeDto.setFriday("8");
		map.put("李凯",workTimeDto);

		File file = new File(inPath);
		File[] listFiles = file.listFiles();
		for (File myfile : listFiles) {
			if (myfile.isFile()) {
				System.out.println("开始读取："+myfile.getName());
//				ReadFile.readSecondPage(myfile.getPath(), list_second, ENUM_CATEGORY.LASTWEEK.code());
				ReadFile.readThirdPage(myfile.getPath(), list_third, ENUM_CATEGORY.NEXTWEEK.code());
//				ReadFile.readFourthPage(myfile.getPath(),map, ENUM_CATEGORY.COST.code());
			}
		}
		System.out.println("读取完毕，执行写入操作！");
//		WriteFile.writeSecondPage(list_second, ouPath,ENUM_CATEGORY.LASTWEEK.code());
		WriteFile.writeThirdPage(list_third, ouPath,ENUM_CATEGORY.NEXTWEEK.code());
//		WriteFile.writeFourthPage(map, ouPath, ENUM_CATEGORY.COST.code());
	}
}
