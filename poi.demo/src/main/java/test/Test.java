package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Util.DateUtil;
import poi.dto.WorkTimeDto;

public class Test extends TimerTask {

    public static void main(String[] args) {
        Timer t =new Timer();
        Test t1 =new Test();
        t.schedule(t1,new Date(),6000);
    }

    @Override
    public void run() {
        try {
            System.out.println(Thread.currentThread().getStackTrace());
            System.out.println("我先睡3秒");
            Thread.sleep(3000);
            System.out.println("睡完继续跑");
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }
}
