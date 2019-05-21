package poi.dto;

import java.util.Date;
import java.util.Objects;

public class XlsDto {
	//任务类型
	private String type;
	//任务名称
	private String name;
	//开始时间
	private Date stardate;
	//结束时间
	private Date enddate;
	//耗时
	private Double days;
	//责任人
	private String dutyperson;
	//完成比例
	private double lv;
	//客户满意度
	private String desc;
	
	@Override
	public String toString() {
		return "XlsDto [type=" + type + ", name=" + name + ", stardate=" + stardate + ", enddate=" + enddate + ", days="
				+ days + ", dutyperson=" + dutyperson + ", lv=" + lv + ", desc=" + desc + "]";
	}

	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		XlsDto xlsDto = (XlsDto) o;
		return Double.compare(xlsDto.lv, lv) == 0 &&
				Objects.equals(type, xlsDto.type) &&
				Objects.equals(name, xlsDto.name) &&
				Objects.equals(stardate, xlsDto.stardate) &&
				Objects.equals(enddate, xlsDto.enddate) &&
				Objects.equals(days, xlsDto.days) &&
				Objects.equals(dutyperson, xlsDto.dutyperson) &&
				Objects.equals(desc, xlsDto.desc);
	}

	@Override
	public int hashCode() {
		return Objects.hash(type, name, stardate, enddate, days, dutyperson, lv, desc);
	}

	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public Date getStardate() {
		return stardate;
	}
	public void setStardate(Date stardate) {
		this.stardate = stardate;
	}
	public Date getEnddate() {
		return enddate;
	}
	public void setEnddate(Date enddate) {
		this.enddate = enddate;
	}
	public Double getDays() {
		return days;
	}
	public void setDays(Double days) {
		this.days = days;
	}
	public String getDutyperson() {
		return dutyperson;
	}
	public void setDutyperson(String dutyperson) {
		this.dutyperson = dutyperson;
	}
	public double getLv() {
		return lv;
	}
	public void setLv(double lv) {
		this.lv = lv;
	}
	public String getDesc() {
		return desc;
	}

	public void setDesc(String desc) {
		this.desc = desc;
	}
	
	
	
	
}
