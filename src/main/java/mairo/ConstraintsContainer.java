package mairo;

import java.util.List;

public class ConstraintsContainer {
  private String area1Name;
  private String area2Name;
  private String area3Name;
  private String area4Name;
  private String area5Name;
  private List<XlsxConstraintRow> constraintRows;

  public String getArea5Name() {
    return area5Name;
  }

  public void setArea5Name(String area5Name) {
    this.area5Name = area5Name;
  }

  public String getArea1Name() {
    return area1Name;
  }

  public void setArea1Name(String area1Name) {
    this.area1Name = area1Name;
  }

  public String getArea2Name() {
    return area2Name;
  }

  public void setArea2Name(String area2Name) {
    this.area2Name = area2Name;
  }

  public String getArea3Name() {
    return area3Name;
  }

  public void setArea3Name(String area3Name) {
    this.area3Name = area3Name;
  }

  public String getArea4Name() {
    return area4Name;
  }

  public void setArea4Name(String area4Name) {
    this.area4Name = area4Name;
  }

  public List<XlsxConstraintRow> getConstraintRows() {
    return constraintRows;
  }

  public void setConstraintRows(List<XlsxConstraintRow> constraintRows) {
    this.constraintRows = constraintRows;
  }
}
