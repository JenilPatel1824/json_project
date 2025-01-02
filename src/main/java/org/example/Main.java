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
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        double id=10000000000001L;

        int flag_occurence=0;
        String flag_occurences="";
        // Open the Excel file

        FileInputStream fis = new FileInputStream(new File("/home/jenil/Downloads/data.xlsx"));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(3);
        int numberOfRows = sheet.getLastRowNum() + 1;
        int numberOfColumns = sheet.getRow(0).getLastCellNum();
        System.out.println(numberOfColumns);
        System.out.println(numberOfRows);

        JSONArray jsonArray = new JSONArray();

        for (int i = 3; i < numberOfRows; i++) {
            Row row = sheet.getRow(i);
            JSONObject jsonObject = new JSONObject();
            if(row.getCell(3).toString()!=""){
                jsonObject.put("rule.name", row.getCell(3));}
            jsonObject.put("rule.catogary", "Network");
                //Rule context obj
            JSONObject rule_context = new JSONObject();
            if(row.getCell(11).toString()!="")
                rule_context.put("rule.check.catagory", row.getCell(11));
            if(row.getCell(10).toString()!=""){
                rule_context.put("rule.check.type", row.getCell(10));}
                    JSONArray rule_conditions = new JSONArray(); //array for rule conditions
                        JSONObject condition = new JSONObject();
            if(row.getCell(12)!=null && row.getCell(12).toString()!="")
                condition.put("condition", row.getCell(12));
            if(row.getCell(12)!=null && row.getCell(13).toString()!=""){
                String cellContent = row.getCell(13).toString(); // Get the cell value
                cellContent = cellContent.replace("\n", "");    // Remove newlines
                condition.put("result.pattern", cellContent);
                condition.put("result.pattern", cellContent);}
                        if(row.getCell(14)!=null)
                        {
                            if(row.getCell(14).toString().equals("any"))
                                condition.put("occurence", -1);
                        }
//                            else
//                              condition.put("occurence", row.getCell(14));
            if(row.getCell(12)!=null && row.getCell(15).toString()!=""){
                condition.put("operator", row.getCell(15));}
                        rule_conditions.put(condition);
                    rule_context.put("rule.conditions", rule_conditions);
                jsonObject.put("rule_context", rule_context);
                jsonObject.put("rule.auto.remediation", "no");
            if(row.getCell(5).toString()!="")
                jsonObject.put("rule.description", row.getCell(5));
            if(row.getCell(17).toString()!=""){
                jsonObject.put("rule.severity", row.getCell(17));}
                jsonObject.put("rule.tags", "tags");
            if(row.getCell(6).toString()!="")
                jsonObject.put("rule.rationale", row.getCell(6));
            if(row.getCell(7).toString()!="")
                jsonObject.put("rule.impact", row.getCell(7));
            if(row.getCell(35).toString()!="")
                jsonObject.put("rule.default.value", row.getCell(35));
            if(row.getCell(34).toString()!="")
                jsonObject.put("rule.references", row.getCell(34));
            if(row.getCell(16).toString()!=""){
                jsonObject.put("rule.additional.information", row.getCell(16));}
                //Array for rule control
                JSONArray rule_controls = new JSONArray();
                JSONObject control1 = new JSONObject();//control 1 obj

            String rule_control_version = row.getCell(18).toString();
            Pattern versionPattern = Pattern.compile("CONTROL:(.*?)(?=DESCRIPTION:)");
            String version1;
            Matcher versionMatcher = versionPattern.matcher(rule_control_version);
            if (versionMatcher.find()) {
                version1 = versionMatcher.group(1).trim();
            }
            else
                version1="";

            if(version1!=""){
                control1.put("rule.control.version", version1);}

                        String input_title = row.getCell(18).toString();
                            String title;
                            Pattern pattern_title = Pattern.compile("^TITLE:(.*)CONTROL:");
                            Matcher matcher_title = pattern_title.matcher(input_title);
                            if (matcher_title.find()) {
                                title = matcher_title.group(1).trim();
                            }
                            else
                                title = "";
                            if(title!="")
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
                        if(description1!="")
                            control1.put("rule.control.description", description1);


                        List<String> rules_ig1 = new ArrayList<>();

                        if (row.getCell(25).toString() != "" ) {
                            //rules_ig1[0] = "ig1";
                            rules_ig1.add("ig1");

                        }
                        if (row.getCell(26).toString() != "" ) {
                            //rules_ig1[1] = "ig2";
                            rules_ig1.add("ig2");
                        }
                        if (row.getCell(27).toString() != "" ) {
                            //rules_ig1[2] = "ig3";
                            rules_ig1.add("ig3");
                        }
                        if(!rules_ig1.isEmpty())
                            control1.put("rule.control.ig", rules_ig1);

                JSONObject control2 = new JSONObject();//control 2 obj
            String rule_control_version2 = row.getCell(19).toString();
            Pattern versionPattern2 = Pattern.compile("CONTROL:(.*?)(?=DESCRIPTION:)");
            String version2;
            Matcher versionMatcher2 = versionPattern2.matcher(rule_control_version2);
            if (versionMatcher2.find()) {
                version2 = versionMatcher2.group(1).trim();
            }
            else
                version2="";

            if(version2!=""){
                control2.put("rule.control.version", version2);}

            String input_title2 = row.getCell(19).toString();
            String title2;

            Pattern pattern_title2 = Pattern.compile("^TITLE:(.*)CONTROL:");
            Matcher matcher_title2 = pattern_title2.matcher(input_title2);

            if (matcher_title2.find()) {
                title2 = matcher_title2.group(1).trim(); // Extract and trim the result
            } else {
                title2 = "";
            }


            if(title2!=""){
                control2.put("rule.control.name", title2);}
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
                    if(description!="")
                        control2.put("rule.control.description", description);
                       List<String> rules_ig2 = new ArrayList<>();
                    StringBuffer sb2 = new StringBuffer("");
                    if (row.getCell(31).toString() != ""){
                        rules_ig2.add("ig1");
                    }
                    if (row.getCell(32).toString() != "" ) {
                        //rules_ig2[1] = "ig2";
                        rules_ig2.add("ig2");
                    }
                    if (row.getCell(33).toString() != "" ) {
//                       rules_ig2[2] = "ig3";
                        rules_ig2.add("ig3");
                    }
                    if(!rules_ig2.isEmpty())
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
//        flag_occurences=sheet.getRow(5).getCell(14).toString();
//        if(flag_occurences.equals("any"))
//        {
//            System.out.println("any true if");
//        }
//        System.out.println(flag_occurence);
//        System.out.println("Last :"+flag_occurences);
        workbook.close();
        fis.close();
    }




}

