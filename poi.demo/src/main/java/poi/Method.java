package poi;

public class Method {
	
	public static void main(String[] args) {
		String inPath = "D:\\project_doc\\project-docs\\3_项目跟踪与监控\\50_个人周报\\2019\\201903\\20190308";// 文件地址
		String ouPath = "D:\\tmp\\LI-国宝人寿-PMC-周报-20190301.xls";// 輸出地址
		Controller c =new Controller();
		try {
			c.method_A(inPath, ouPath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
