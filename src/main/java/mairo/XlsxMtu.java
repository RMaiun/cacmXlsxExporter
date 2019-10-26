package mairo;

import java.util.List;

public class XlsxMtu {
  private String interval;
  private String summary;

  private ConstraintsContainer constraintContainer;
  private List<ResultXlsxTemplate> results;

  public String getInterval() {
    return interval;
  }

  public void setInterval(String interval) {
    this.interval = interval;
  }

  public String getSummary() {
    return summary;
  }

  public void setSummary(String summary) {
    this.summary = summary;
  }

  public ConstraintsContainer getConstraintContainer() {
    return constraintContainer;
  }

  public void setConstraintContainer(ConstraintsContainer constraintContainer) {
    this.constraintContainer = constraintContainer;
  }

  public List<ResultXlsxTemplate> getResults() {
    return results;
  }

  public void setResults(List<ResultXlsxTemplate> results) {
    this.results = results;
  }


}
