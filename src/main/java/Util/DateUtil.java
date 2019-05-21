package Util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class DateUtil {

	public static boolean jdugeBeforeDate(Date date){
		Date now =new Date();
		Calendar calendar =Calendar.getInstance();
		calendar.setTime(now);
		calendar.add(Calendar.DATE, -7);//һ�������ж�
		if(date.before(calendar.getTime())||date.after(now)){
			//��ʾ���ڲ��ڱ�������
			return true;
		}
		return false;
	}
	
	/**
	 * ������_ʼ���^�˶�����
	 * @return
	 */
	public static int getDaysFromStar(){
		Date now =new Date();
		Date star=null;
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		try {
			star = sdf.parse("2019-01-01");
		} catch (ParseException e) {
			e.printStackTrace();
		}
		long nowtime = now.getTime();
		long startime = star.getTime();
		long i=(nowtime-startime)/3600000/24;
		return (int) i;
	}
	
	public static boolean jdugeAfterDate(Date date){
		Date now =new Date();
		Calendar calendar =Calendar.getInstance();
		calendar.setTime(now);
		calendar.add(Calendar.DATE, +7);//һ�������ж�
		if(date.after(calendar.getTime())||date.before(now)){
			//��ʾ���ڲ��ڱ�������
			return true;
		}
		return false;
	}
	
	public static void main(String[] args) {
		System.out.println(getDaysFromStar());
	}
}
