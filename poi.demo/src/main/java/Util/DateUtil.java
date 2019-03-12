package Util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class DateUtil {

	public static boolean jdugeDate(Date date){
		Date now =new Date();
		Calendar calendar =Calendar.getInstance();
		calendar.setTime(now);
		calendar.add(Calendar.DATE, -7);//一周以内判断
		if(date.before(calendar.getTime())||date.after(now)){
			//表示日期不在本周以内
			return true;
		}
		return false;
	}
	
	/**
	 * 求今年開始度過了多少天
	 * @return
	 */
	public static long getDaysFromStar(){
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
		return i;
	}
	
	public static void main(String[] args) {
		System.out.println(getDaysFromStar());
	}
}
