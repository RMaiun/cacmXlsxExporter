package mairo.dto;

import java.util.List;

public class XlsxMtu {
  private String interval;
  private int constraintsNum;
  private int resultsNum;
  private ConstraintsContainerXlsxTemplate constraintContainer;
  private List<ResultXlsxTemplate> results;

  public int getConstraintsNum() {
    return constraintsNum;
  }

  public void setConstraintsNum(int constraintsNum) {
    this.constraintsNum = constraintsNum;
  }

  public int getResultsNum() {
    return resultsNum;
  }

  public void setResultsNum(int resultsNum) {
    this.resultsNum = resultsNum;
  }

  public String getInterval() {
    return interval;
  }

  public void setInterval(String interval) {
    this.interval = interval;
  }


  public ConstraintsContainerXlsxTemplate getConstraintContainer() {
    return constraintContainer;
  }

  public void setConstraintContainer(ConstraintsContainerXlsxTemplate constraintContainer) {
    this.constraintContainer = constraintContainer;
  }

  public List<ResultXlsxTemplate> getResults() {
    return results;
  }

  public void setResults(List<ResultXlsxTemplate> results) {
    this.results = results;
  }


}
