package poi.io;

import ENUM.ENUM_CATEGORY;
import Util.DateUtil;
import Util.StringUtil;
import com.sun.org.apache.bcel.internal.generic.RETURN;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import poi.dto.MyObject;
import poi.dto.WorkTimeDto;
import poi.dto.XlsDto;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.CountDownLatch;

public class ThreadReadFile implements Runnable {

    private String path;
    private CountDownLatch countDownLatch;
    private MyObject myObject;

    public ThreadReadFile(String path, CountDownLatch countDownLatch, MyObject myObject) {
        this.path = path;
        this.countDownLatch = countDownLatch;
        this.myObject = myObject;
    }

    @Override
    public void run() {
        try {
            String name = path.substring(path.lastIndexOf("-") + 1, path.lastIndexOf("."));
            System.out.println(Thread.currentThread().getName()+"获取并执行了一个任务："+name);
            InputStream  is = new FileInputStream(path);
            XSSFWorkbook hw = new XSSFWorkbook(is);

            getSecond(hw.getSheetAt(ENUM_CATEGORY.LASTWEEK.code()),myObject.getSecond());
            getThired(hw.getSheetAt(ENUM_CATEGORY.NEXTWEEK.code()),myObject.getThired());
            getFourth(hw.getSheetAt(ENUM_CATEGORY.COST.code()),myObject.getFourth());

            is.close();
            countDownLatch.countDown();
            System.out.println(Thread.currentThread().getName()+"执行任务结束："+name);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private void getSecond(Sheet sheet,LinkedHashSet<XlsDto> list) {
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
    }

    private void getThired(Sheet sheet,LinkedHashSet<XlsDto> list){
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
    }

    private void getFourth(Sheet sheet,HashMap<String, WorkTimeDto> map){
        String name = path.substring(path.lastIndexOf("-") + 1, path.lastIndexOf("."));
        WorkTimeDto workDto = new WorkTimeDto();
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
        System.out.println("姓名："+name+",work:"+workDto);
        map.put(name, workDto);
    }
}
