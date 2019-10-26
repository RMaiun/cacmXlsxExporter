package mairo;

public final class ResultXlsxTemplateBuilder {
  private String direction;
  private String name;
  private String eic;
  private String type;
  private String location;
  private String maxFlow;
  private String maxExchange;
  private String maxExchangeAfterRemedy;
  private String aac;
  private String trm;
  private String ttc;

  private ResultXlsxTemplateBuilder() {
  }

  public static ResultXlsxTemplateBuilder aResultXlsxTemplate() {
    return new ResultXlsxTemplateBuilder();
  }

  public ResultXlsxTemplateBuilder withDirection(String direction) {
    this.direction = direction;
    return this;
  }

  public ResultXlsxTemplateBuilder withName(String name) {
    this.name = name;
    return this;
  }

  public ResultXlsxTemplateBuilder withEic(String eic) {
    this.eic = eic;
    return this;
  }

  public ResultXlsxTemplateBuilder withType(String type) {
    this.type = type;
    return this;
  }

  public ResultXlsxTemplateBuilder withLocation(String location) {
    this.location = location;
    return this;
  }

  public ResultXlsxTemplateBuilder withMaxFlow(String maxFlow) {
    this.maxFlow = maxFlow;
    return this;
  }

  public ResultXlsxTemplateBuilder withMaxExchange(String maxExchange) {
    this.maxExchange = maxExchange;
    return this;
  }

  public ResultXlsxTemplateBuilder withMaxExchangeAfterRemedy(String maxExchangeAfterRemedy) {
    this.maxExchangeAfterRemedy = maxExchangeAfterRemedy;
    return this;
  }

  public ResultXlsxTemplateBuilder withAac(String aac) {
    this.aac = aac;
    return this;
  }

  public ResultXlsxTemplateBuilder withTrm(String trm) {
    this.trm = trm;
    return this;
  }

  public ResultXlsxTemplateBuilder withTtc(String ttc) {
    this.ttc = ttc;
    return this;
  }

  public ResultXlsxTemplate build() {
    ResultXlsxTemplate resultXlsxTemplate = new ResultXlsxTemplate();
    resultXlsxTemplate.setDirection(direction);
    resultXlsxTemplate.setName(name);
    resultXlsxTemplate.setEic(eic);
    resultXlsxTemplate.setType(type);
    resultXlsxTemplate.setLocation(location);
    resultXlsxTemplate.setMaxFlow(maxFlow);
    resultXlsxTemplate.setMaxExchange(maxExchange);
    resultXlsxTemplate.setMaxExchangeAfterRemedy(maxExchangeAfterRemedy);
    resultXlsxTemplate.setAac(aac);
    resultXlsxTemplate.setTrm(trm);
    resultXlsxTemplate.setTtc(ttc);
    return resultXlsxTemplate;
  }
}
