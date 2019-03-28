package poi.io;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Util.CellUtil;
import Util.DateUtil;
import Util.SaveUtil;
import Util.StringUtil;
import poi.dto.CellTypeDto;
import poi.dto.WorkTimeDto;
import poi.dto.XlsDto;

public class WriteFile {

    public static void writeSecondPage(HashSet<XlsDto> list, String ouPath, Integer integer) throws Exception {
        CellUtil cellUtil = new CellUtil();
        SaveUtil saveUtil = new SaveUtil();

        FileInputStream fs = new FileInputStream(ouPath);
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(integer);
        CellTypeDto tpye = cellUtil.getTpye(sheet.getRow(2));
        XSSFRow row = sheet.getRow(0);
        int lastRowNum = sheet.getLastRowNum();
        int star = 2;
        FileOutputStream out = new FileOutputStream(ouPath);
        try {
            for (int i = star; i < lastRowNum; i++) {
                row = sheet.getRow(i);
                sheet.removeRow(row);
            }

            for (XlsDto xlsDto : list) {
                row = sheet.createRow(star++);
                saveUtil.SaveMethod("1", tpye.getXssf0(), row.createCell(0));
                saveUtil.SaveMethod(xlsDto.getType(), tpye.getXssf1(), row.createCell(1));
                saveUtil.SaveMethod(xlsDto.getName(), tpye.getXssf2(), row.createCell(2));
                saveUtil.SaveMethodBefore(xlsDto.getStardate(), tpye.getXssf3(), row.createCell(3));
                saveUtil.SaveMethodBefore(xlsDto.getEnddate(), tpye.getXssf4(), row.createCell(4));
                saveUtil.SaveMethod(xlsDto.getDays(), tpye.getXssf5(), row.createCell(5));
                saveUtil.SaveMethod(xlsDto.getDutyperson(), tpye.getXssf6(), row.createCell(6));
                saveUtil.SaveMethod(xlsDto.getLv(), tpye.getXssf7(), row.createCell(7));
                saveUtil.SaveMethod(xlsDto.getDesc(), tpye.getXssf8(), row.createCell(8));
            }
        } catch (Exception e) {
            System.out.println("Second有一个异常："+e.getMessage());
        } finally {
            wb.write(out);
            out.flush();
            out.close();
            fs.close();
            System.out.println("Second写入完毕");
        }
    }

    public static void writeThirdPage(HashSet<XlsDto> list, String ouPath, Integer integer) throws Exception {
        CellUtil cellUtil = new CellUtil();
        SaveUtil saveUtil = new SaveUtil();

        FileInputStream fs = new FileInputStream(ouPath);
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(integer);
        CellTypeDto tpye = cellUtil.getTpye(sheet.getRow(2));
        XSSFRow row = sheet.getRow(0);
        int lastRowNum = sheet.getLastRowNum();
        int star = 2;
        FileOutputStream out = new FileOutputStream(ouPath);
        try {
            for (int i = star; i < lastRowNum; i++) {
                row = sheet.getRow(i);
                sheet.removeRow(row);
            }
            for (XlsDto xlsDto : list) {
                row = sheet.createRow(star++);
                saveUtil.SaveMethod("1", tpye.getXssf0(), row.createCell(0));
                saveUtil.SaveMethod(xlsDto.getType(), tpye.getXssf1(), row.createCell(1));
                saveUtil.SaveMethod(xlsDto.getName(), tpye.getXssf2(), row.createCell(2));
                saveUtil.SaveMethodAfter(xlsDto.getStardate(), tpye.getXssf3(), row.createCell(3));
                saveUtil.SaveMethodAfter(xlsDto.getEnddate(), tpye.getXssf4(), row.createCell(4));
                saveUtil.SaveMethod(xlsDto.getDays(), tpye.getXssf5(), row.createCell(5));
                saveUtil.SaveMethod(xlsDto.getDutyperson(), tpye.getXssf6(), row.createCell(6));
                saveUtil.SaveMethod("1", tpye.getXssf7(), row.createCell(7));
            }
        } catch (Exception e) {
            System.out.println("Third有一个异常:"+e.getMessage());
        } finally {
            wb.write(out);
            out.flush();
            out.close();
            fs.close();
            System.out.println("Third写入完毕");
        }
    }

    public static void writeFourthPage(Map<String, WorkTimeDto> map, String ouPath, Integer integer) throws Exception {

        Map<String, Integer> nameMap = new HashMap<String, Integer>();

        FileInputStream fs = new FileInputStream(ouPath);
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(integer);
        XSSFRow row = sheet.getRow(1);
        int lastCellNum = row.getLastCellNum();
        for (int i = 1; i < lastCellNum; i++) {
            if (row.getCell(i) != null) {
                nameMap.put(StringUtil.getCellValue(row.getCell(i)), i);
            }
        }
        //遍历需要写入的姓名，找出其对应的位置
        FileOutputStream out = new FileOutputStream(ouPath);
        try {
            for (String name : map.keySet()) {
                Integer addnum = nameMap.get(name);
                WorkTimeDto workTimeDto = map.get(name);
                int num = DateUtil.getDaysFromStar() + 4;
                for (int i = num - 7; i < num; i++) {
                    int ca = num - i;
                    XSSFCell cell = sheet.getRow(i).getCell(addnum);
                    switch (ca) {
                        case 7:
                            if (StringUtil.isNumeric(workTimeDto.getSunday())) {
                                cell.setCellValue(Double.parseDouble(workTimeDto.getSunday()));
                            } else {
                                cell.setCellValue(workTimeDto.getSunday());
                            }
                            break;
                        case 6:
                            if (StringUtil.isNumeric(workTimeDto.getSaturday())) {
                                cell.setCellValue(Double.parseDouble(workTimeDto.getSaturday()));
                            } else {
                                cell.setCellValue(workTimeDto.getSaturday());
                            }
                            break;
                        case 5:
                            if (StringUtil.isNumeric(workTimeDto.getFriday())) {
                                cell.setCellValue(Double.parseDouble(workTimeDto.getFriday()));
                            } else {
                                cell.setCellValue(workTimeDto.getFriday());
                            }
                            break;
                        case 4:
                            if (StringUtil.isNumeric(workTimeDto.getThursday())) {
                                cell.setCellValue(Double.parseDouble(workTimeDto.getThursday()));
                            } else {
                                cell.setCellValue(workTimeDto.getThursday());
                            }
                            break;
                        case 3:
                            if (StringUtil.isNumeric(workTimeDto.getWednesday())) {
                                cell.setCellValue(Double.parseDouble(workTimeDto.getWednesday()));
                            } else {
                                cell.setCellValue(workTimeDto.getWednesday());
                            }
                            break;
                        case 2:
                            if (StringUtil.isNumeric(workTimeDto.getTuesday())) {
                                cell.setCellValue(Double.parseDouble(workTimeDto.getTuesday()));
                            } else {
                                cell.setCellValue(workTimeDto.getTuesday());
                            }
                            break;
                        case 1:
                            if (StringUtil.isNumeric(workTimeDto.getMonday())) {
                                cell.setCellValue(Double.parseDouble(workTimeDto.getMonday()));
                            } else {
                                cell.setCellValue(workTimeDto.getMonday());
                            }
                            break;
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Fourth有一个异常:"+e.getMessage());
        } finally {
            wb.write(out);
            out.flush();
            out.close();
            fs.close();
            System.out.println("Fourth写入完毕");
        }
    }

    public static void main(String[] args) {
        String ouPath = "D:\\Config\\LI-国宝人寿-PMC-周报-20190301.xls";
        int page = 3;
        WorkTimeDto w = new WorkTimeDto();
        w.setMonday("9.2");
        w.setTuesday("9.5");
        w.setWednesday("9");
        w.setThursday("9.8");
        w.setFriday("请假");
        Map<String, WorkTimeDto> map = new HashMap<String, WorkTimeDto>();
        map.put("赵政", w);
        try {
            writeFourthPage(map, ouPath, page);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
