import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        double id=10000000000001L;
        // Open the Excel file
        FileInputStream fis = new FileInputStream(new File("/home/jenil/Downloads/data.xlsx"));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(3);
        int numberOfRows = sheet.getLastRowNum() + 1;
        int numberOfColumns = sheet.getRow(0).getLastCellNum();
        System.out.println(numberOfColumns);
        System.out.println(numberOfRows);

        JSONArray jsonArray = new JSONArray();

        for (int i = 1; i < numberOfRows; i++) {
            Row row = sheet.getRow(i);
            JSONObject jsonObject = new JSONObject();
                jsonObject.put("rule.name", row.getCell(3));
                jsonObject.put("rule.catogary", "Network");
                //Rule context obj
                JSONObject rule_context = new JSONObject();
                    rule_context.put("rule.check.catagory", row.getCell(11));
                    rule_context.put("rule.check.type", row.getCell(10));
                    JSONArray rule_conditions = new JSONArray(); //array for rule conditions
                        JSONObject condition = new JSONObject();
                        condition.put("condition", row.getCell(12));
                        condition.put("result.pattern", row.getCell(13));
                        if(row.getCell(14)!=null)
                        {
                            if(row.getCell(14).toString() == "any")
                                condition.put("occurence", -1);}
//                          else
//                              condition.put("occurence", row.getCell(14));
                        condition.put("operator", row.getCell(15));
                        rule_conditions.put(condition);
                    rule_context.put("rule.conditions", rule_conditions);
                jsonObject.put("rule_context", rule_context);
                jsonObject.put("rule.auto.remediation", "no");
                jsonObject.put("rule.description", row.getCell(5));
                jsonObject.put("rule.severity", row.getCell(17));
                jsonObject.put("rule.tags", "tags");
                jsonObject.put("rule.rationale", row.getCell(6));
                jsonObject.put("rule.impact", row.getCell(7));
                jsonObject.put("rule.default.value", row.getCell(35));
                jsonObject.put("rule.references", row.getCell(34));
                jsonObject.put("rule.additional.information", row.getCell(16));
                //Array for rule control
                JSONArray rule_controls = new JSONArray();
                JSONObject control1 = new JSONObject();//control 1 obj
                    control1.put("rule.control.version", "v"+row.getCell(22));
                        String input_title = row.getCell(18).toString();
                            String title;
                            Pattern pattern_title = Pattern.compile("TITLE:(.*)");
                            Matcher matcher_title = pattern_title.matcher(input_title);
                            if (matcher_title.find()) {
                                title = matcher_title.group(1).trim(); // Extract and trim the result
                            }
                            else
                                title = "";
                    control1.put("rule.control.name", title);

                        String input1 = row.getCell(18).toString();
                        String description1;
                        Pattern pattern1 = Pattern.compile("DESCRIPTION:(.*)");
                        Matcher matcher1 = pattern1.matcher(input1);
                        if (matcher1.find()) {
                                description1 = matcher1.group(1).trim(); // Extract and trim the result
                        }
                        else
                            description1 = "";
                        control1.put("rule.control.description", description1);

                        String[] rules_ig1 = {"", "", ""};
                        if (row.getCell(25).toString() != "" ) {
                            rules_ig1[0] = "ig1";
                        }
                        if (row.getCell(26).toString() != "" ) {
                            rules_ig1[1] = "ig2";
                        }
                        if (row.getCell(27).toString() != "" ) {
                            rules_ig1[2] = "ig3";
                        }
                    control1.put("rule.control.ig", rules_ig1);

                JSONObject control2 = new JSONObject();//control 2 obj
                    control2.put("rule.control.version", "v"+row.getCell(28));

                    String input_title2 = row.getCell(18).toString();
                    String title2;
                    Pattern pattern_title2 = Pattern.compile("TITLE:(.*)");
                    Matcher matcher_title2 = pattern_title2.matcher(input_title2);
                    if (matcher_title2.find()) {
                        title2 = matcher_title.group(1).trim(); // Extract and trim the result
                    }
                    else
                        title2 = "";

                    control2.put("rule.control.name", title2);
                    control2.put("rule.control.name", title2);
                    String input = row.getCell(19).toString();
                    //extract description
                    String description;
                    Pattern pattern = Pattern.compile("DESCRIPTION:(.*)");
                    Matcher matcher = pattern.matcher(input);
                    if (matcher.find()) {
                        description = matcher.group(1).trim(); // Extract and trim the result
                    }
                    else
                        description = "";
                    control2.put("rule.control.description", description);
                    String[] rules_ig2 = {"", "", ""};
                    if (row.getCell(31).toString() != ""){
                        rules_ig2[0] = "ig1";
                    }
                    if (row.getCell(32).toString() != "" ) {
                        rules_ig2[1] = "ig2";
                    }
                    if (row.getCell(33).toString() != "" ) {
                       rules_ig2[2] = "ig3";
                    }
                    control2.put("rule.control.ig", rules_ig2);
                rule_controls.put(control1);
                rule_controls.put(control2);
            jsonObject.put("rule.controls", rule_controls);
            jsonObject.put("id", id++);
            jsonArray.put(jsonObject);
        }
        System.out.println(jsonArray);
        FileWriter f=new FileWriter(new File("/home/jenil/Downloads/output.json"));
        f.write(jsonArray.toString());
        f.close();

//        String s = sheet.getRow(3).getCell(25).toString();
//        System.out.println(s=="");
        //System.out.println("output "+s);
        workbook.close();
        fis.close();
    }




}

