package poi.run;

import poi.controller.Controller;

public class Start {
	
	public static void main(String[] args) {
		String inPath = "D:\\project_doc\\project-docs\\3_��Ŀ��������\\50_�����ܱ�\\2019\\201903\\20190308";// �ļ���ַ
		String ouPath = "D:\\tmp\\LI-��������-PMC-�ܱ�-20190301.xls";// ݔ����ַ
		Controller c =new Controller();
		try {
			c.method_A(inPath, ouPath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}