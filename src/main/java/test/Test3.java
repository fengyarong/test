package test;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import poi.dto.CellTypeDto;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Test3 {
    public static void main(String[] args) throws Exception{
        String ouPath="D:\\project_doc\\project-docs\\3_项目跟踪与监控\\20_项目周报\\QA\\2019\\LI-国宝人寿-PMC-周报-20190315.xls";
        Integer integer=1;
        FileInputStream fs = new FileInputStream(ouPath);
        XSSFWorkbook wb = new XSSFWorkbook(fs);
        XSSFSheet sheet = wb.getSheetAt(integer);
        int lastRowNum = sheet.getLastRowNum();
        System.out.printf("last:"+lastRowNum);
        FileOutputStream out = new FileOutputStream(ouPath);
        while (lastRowNum>2){
            sheet.shiftRows(2,lastRowNum,-1);
        }
//        sheet.removeRowBreak(5);
        wb.write(out);
        out.flush();
        out.close();
        fs.close();
        System.out.printf("end");
    }

    public static void removeRow(HSSFSheet sheet, int rowIndex) {
        int lastRowNum=sheet.getLastRowNum();
        if(rowIndex>=0&&rowIndex<lastRowNum)
            sheet.shiftRows(rowIndex+1,lastRowNum,-1);//将行号为rowIndex+1一直到行号为lastRowNum的单元格全部上移一行，以便删除rowIndex行
        if(rowIndex==lastRowNum){
            HSSFRow removingRow=sheet.getRow(rowIndex);
            if(removingRow!=null)
                sheet.removeRow(removingRow);
        }
    }

}
