package ENUM;

public enum ENUM_CATEGORY {
	LASTWEEK(1,"上周完成任务"),
	NEXTWEEK(2,"本周计划任务"),
	COST(3,"成本管理");
	private final Integer code;
    private final String desc;

    ENUM_CATEGORY(Integer code ,String desc) {
        this.code = code;
        this.desc = desc;
    }

    public Integer code() {
        return code;
    }

    public String desc() {
        return desc;
    }
}
