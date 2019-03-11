package poi;

public class Method {
	
	public static void main(String[] args) {
		String inPath = "D:/project/文档库/project-docs/3_项目跟踪与监控/50_个人周报/2019/201903/20190308";// 文件地址
		String ouPath = "C:/Users/lizhe/Desktop/脚本执行.xls";// 輸出地址
		Controller c =new Controller();
		try {
			c.method_A(inPath, ouPath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
