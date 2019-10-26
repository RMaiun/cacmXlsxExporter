package mairo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class Main {

  public static void main(String[] args) throws IOException {
    FileOutputStream fileOut = new FileOutputStream("C:\\Hobby\\cacmXlsxExporter\\src\\main\\resources\\templates\\contacts.xlsx");
    UberWriter writer = new UberWriter();

    NtcBasedAllocationXlsxTemplate data = new NtcBasedAllocationXlsxTemplate();
    data.setTimeInterval("16.05.2019 00:00 - 17.05.2019 00:00");
    data.setTimezone("CET");
    data.setRegion("REG|CWE");

    XlsxMtu mtu1 = new XlsxMtu();
    mtu1.setInterval("16.05.2019 00:00 - 01:00");
    mtu1.setSummary("There is/are 195 CB/CO and 2 results in MTU.");
    ConstraintsContainer constraints = new ConstraintsContainer();
    constraints.setArea1Name("BZN|FR");
    constraints.setArea2Name("BZN|DE-LU");
    constraints.setArea3Name("BZN|NL");
    constraints.setArea4Name("BZN|BE");
    constraints.setArea5Name("BZN|AT");
    constraints.setConstraintRows(constraintValuesDtoList());
    mtu1.setConstraintContainer(constraints);
    mtu1.setResults(resultXlsxTemplateList());
    data.setMtuList(Arrays.asList(mtu1, mtu1));
    writer.generateTemplate(data, fileOut);
    fileOut.close();
  }

  private static List<ResultXlsxTemplate> resultXlsxTemplateList() {
    ResultXlsxTemplate r1 = ResultXlsxTemplateBuilder.aResultXlsxTemplate()
        .withDirection("BZN 1 > BZN 2")
        .withName("IC1")
        .withEic("22T2021…")
        .withType("AC Link")
        .withLocation("Location 1")
        .withMaxFlow("200")
        .withMaxExchange("1342")
        .withMaxExchangeAfterRemedy("2182")
        .withAac("-0.124")
        .withTrm("1.1014")
        .withTtc("10324")
        .build();

    ResultXlsxTemplate r2 = ResultXlsxTemplateBuilder.aResultXlsxTemplate()
        .withDirection("BZN 1 > BZN 2")
        .withName("IC2")
        .withEic("22T4021…")
        .withType("AC Link")
        .withLocation("Location 2")
        .withMaxFlow("220")
        .withMaxExchange("1542")
        .withMaxExchangeAfterRemedy("5182")
        .withAac("-1.124")
        .withTrm("8.1014")
        .withTtc("90324")
        .build();
    return Arrays.asList(r1, r2);
  }

  private static List<XlsxConstraintRow> constraintValuesDtoList() {

    XlsxConstraintRow dto1 = new XlsxConstraintRow();
    dto1.setConstraintId("18498330000");
    dto1.setCbCoName("Constraint1");
    dto1.setCbCoEic("22T2021…");
    dto1.setCbCoType("Transformer");
    dto1.setCbCoLocation("Location 1");
    dto1.setCbCoAreaIn("BZN 1");
    dto1.setCbCoAreaOut("BZN1");
    dto1.setArea1("-0.00208");
    dto1.setArea2("-0.00199");
    dto1.setArea3("-0.01844");
    dto1.setArea4("-0.09726");
    dto1.setArea5("0.00196");
    dto1.setfRef("1342");
    dto1.setfMax("2182");
    dto1.setActualFlow("164");

    XlsxConstraintRow dto2 = new XlsxConstraintRow();
    dto2.setConstraintId("18498330001");
    dto2.setCbCoName("Constraint2");
    dto2.setCbCoEic("12T2021…");
    dto2.setCbCoType("Ac Link");
    dto2.setCbCoLocation("Location 2");
    dto2.setCbCoAreaIn("BZN 2");
    dto2.setCbCoAreaOut("BZN2");
    dto2.setArea1("1.00208");
    dto2.setArea2("-0.10199");
    dto2.setArea3("2.01844");
    dto2.setArea4("3.09726");
    dto2.setArea5("4.00196");
    dto2.setfRef("1542");
    dto2.setfMax("7182");
    dto2.setActualFlow("964");
    return Arrays.asList(dto1, dto2);
  }
}
