package test;

import java.util.Date;

public class Test1 {
	public static void main(String[] args) {
		double f =0.0d/0.0d;
//		System.out.println(Float.isNaN(f));
		System.out.println(f);
		System.out.println(f==f);
		float v=5;
		System.out.println(Float.isNaN(v));
		System.out.println(v==v);
		Date d =new Date();
		System.out.println(d==d);
	}
}
