package poi.dto;

import java.util.HashMap;
import java.util.LinkedHashSet;

public class MyObject {
    private LinkedHashSet<XlsDto> second;
    private LinkedHashSet<XlsDto> thired;
    private HashMap<String,WorkTimeDto> fourth;

    public LinkedHashSet<XlsDto> getSecond() {
        return second;
    }

    public void setSecond(LinkedHashSet<XlsDto> second) {
        this.second = second;
    }

    public LinkedHashSet<XlsDto> getThired() {
        return thired;
    }

    public void setThired(LinkedHashSet<XlsDto> thired) {
        this.thired = thired;
    }

    public HashMap<String, WorkTimeDto> getFourth() {
        return fourth;
    }

    public void setFourth(HashMap<String, WorkTimeDto> fourth) {
        this.fourth = fourth;
    }
}
