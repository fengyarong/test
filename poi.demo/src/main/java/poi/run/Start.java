package poi.run;

import poi.controller.Controller;

public class Start {
	
	public static void main(String[] args) {
		String inPath = "D:\\project_doc\\project-docs\\3_项目跟踪与监控\\50_个人周报\\2019\\201903\\20190322";// 文件地址
		String ouPath = "D:\\project_doc\\project-docs\\3_项目跟踪与监控\\20_项目周报\\QA\\2019\\LI-国宝人寿-PMC-周报-20190315.xls";// 輸出地址
		Controller c =new Controller();
		try {
			c.doItBegin(inPath, ouPath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
