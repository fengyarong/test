package poi.run;

import poi.controller.Controller;

public class Start {
	
	public static void main(String[] args) {
		long l1 = System.currentTimeMillis();
		String inPath = "D:\\project_doc\\project-docs\\3_项目跟踪与监控\\50_个人周报\\2019\\201905\\20190510";// 文件地址
		String ouPath = "D:\\project_doc\\project-docs\\3_项目跟踪与监控\\20_项目周报\\QA\\2019\\LI-国宝人寿-PMC-周报-20190510.xlsx";// 輸出地址
//		String ouPath="C:\\Users\\April\\Desktop\\LI-国宝人寿-PMC-周报-20190412.xlsx";
		Controller c =new Controller();
		try {
			c.doItBegin(inPath, ouPath);
		} catch (Exception e) {
			e.printStackTrace();
		}
		long l2 = System.currentTimeMillis();
		System.out.println("一共花费了"+(l2-l1)/1000+"S");
	}
}
