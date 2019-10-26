package mairo.dto;

import java.util.List;

public class NtcBasedAllocationXlsxTemplate {
  private String timeInterval;
  private String timezone;
  private String region;

  private List<XlsxMtu> mtuList;

  public String getTimeInterval() {
    return timeInterval;
  }

  public void setTimeInterval(String timeInterval) {
    this.timeInterval = timeInterval;
  }

  public String getTimezone() {
    return timezone;
  }

  public void setTimezone(String timezone) {
    this.timezone = timezone;
  }

  public String getRegion() {
    return region;
  }

  public void setRegion(String region) {
    this.region = region;
  }

  public List<XlsxMtu> getMtuList() {
    return mtuList;
  }

  public void setMtuList(List<XlsxMtu> mtuList) {
    this.mtuList = mtuList;
  }
}
