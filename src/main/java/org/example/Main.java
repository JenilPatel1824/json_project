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
            if(row.getCell(11).toString().equals("advance"))
            {
                continue;
            }
            JSONObject jsonObject = new JSONObject();
                jsonObject.put("rule.name", row.getCell(3));
            jsonObject.put("rule.category", "Network");
//            jsonObject.put("rule.type","Default");
                //Rule context obj
            JSONObject rule_context = new JSONObject();
                rule_context.put("rule.check.category", row.getCell(11));
                rule_context.put("rule.check.type", row.getCell(10));
                    JSONArray rule_conditions = new JSONArray(); //array for rule conditions
                        JSONObject condition = new JSONObject();
                        int condition_flag1=0;
                        int condition_flag2=0;
                        int condition_flag3=0;
                        int condition_flag4=0;

                        if(row.getCell(12)!=null && row.getCell(12).toString().equals(""))
                        {
                            condition_flag1=1;
                        }
                condition.put("condition", row.getCell(12));
            if(row.getCell(12)!=null ) {
                String cellContent = row.getCell(13).toString(); // Get the cell value
                cellContent = cellContent.replace("\n", "");
                // Remove newlines

                if (cellContent.equals("")) {
                    condition_flag2 = 1;
                }
                condition.put("result.pattern", cellContent);
                if (row.getCell(14) != null) {
                    if (row.getCell(14).toString().equals("any"))
                        condition.put("occurrence", -1);
                    else if(row.getCell(14).toString().equals(""))
                    {
                        condition.put("occurrence", "");
                    }
                    else
                    {
                        condition.put("occurrence", 1);
                    }
                }
                if (row.getCell(14).toString().equals("")) {
                    condition_flag3 = 1;
                }
//                            else
//                              condition.put("occurence", row.getCell(14));
                if (row.getCell(15) != null) {
                    String operator = row.getCell(15).toString();
                    String upper = operator.toUpperCase();
                    if (row.getCell(15).toString().equals("")) {
                        condition_flag4 = 1;
                    }
                    condition.put("operator", upper);
                }
                // if(condition_flag1==0 && condition_flag2==0 && condition_flag3==0 && condition_flag4==0){}
                rule_conditions.put(condition);
                //if(condition_flag1==0 && condition_flag2==0 && condition_flag3==0 && condition_flag4==0);{}


                if (condition_flag1 == 1 && condition_flag2 == 1 && condition_flag3 == 1 && condition_flag4 == 1) {
                    rule_context.put("rule_conditions", new JSONArray());
                } else {
                    rule_context.put("rule.conditions", rule_conditions);
                }


                jsonObject.put("rule.context", rule_context);
                if (row.getCell(8) != null) {
                    jsonObject.put("rule.auto.remediation", row.getCell(8));
                }
                jsonObject.put("rule.description", row.getCell(5).toString());
                if (row.getCell(17).toString() != "") {
                    String sev = row.getCell(17).toString();
                    String sevupper = sev.toUpperCase();
                    jsonObject.put("rule.severity", sevupper);
                }
                else {
                    jsonObject.put("rule.severity", "");
                }
                List<String> rule_tags_list = new ArrayList<>();
                jsonObject.put("rule.tags", rule_tags_list);
                jsonObject.put("rule.rationale", row.getCell(6).toString());
                jsonObject.put("rule.impact", row.getCell(7));
                jsonObject.put("rule.default.value", row.getCell(35));
                jsonObject.put("rule.references", row.getCell(34));
                jsonObject.put("rule.additional.information", row.getCell(16));
                //Array for rule control
                JSONArray rule_controls = new JSONArray();
                JSONObject control1 = new JSONObject();//control 1 obj
                int control_flag1 = 0;
                int control_flag2 = 0;
                int control_flag3 = 0;
                int control_flag4 = 0;

                String rule_control_version = row.getCell(18).toString();

                 //System.out.println("result:"+row.getCell(22).toString().equals("0.0"));
                if (!row.getCell(22).toString().equals("0.0") && !row.getCell(22).toString().equals("") ) {
                    control1.put("rule.control.version", row.getCell(22));
                } else  {
                    control_flag4 = 1;
                    control1.put("rule.control.version", "");

                }


                String input_title = row.getCell(18).toString();
                String title;
                Pattern pattern_title = Pattern.compile("^TITLE:(.*)CONTROL:");
                Matcher matcher_title = pattern_title.matcher(input_title);
                if (matcher_title.find()) {
                    title = matcher_title.group(1).trim();
                } else
                    title = "";

                if (title.equals("")) {
                    control_flag1 = 1;
                }

                control1.put("rule.control.name", title);

                String input1 = row.getCell(18).toString();
                String description1;
                Pattern pattern1 = Pattern.compile("DESCRIPTION:(.*)");
                Matcher matcher1 = pattern1.matcher(input1);
                if (matcher1.find()) {
                    description1 = matcher1.group(1).trim(); // Extract and trim the result
                } else
                    description1 = "";
                if (description1.equals("")) {
                    control_flag2 = 1;
                }

                control1.put("rule.control.description", description1);


                List<String> rules_ig1 = new ArrayList<>();

                if (row.getCell(25).toString().equals("X")) {
                    //rules_ig1[0] = "ig1";
                    rules_ig1.add("ig1");

                }
                if (row.getCell(26).toString().equals("X")) {
                    //rules_ig1[1] = "ig2";
                    rules_ig1.add("ig2");
                }
                if (row.getCell(27).toString().equals("X")) {
                    //rules_ig1[2] = "ig3";
                    rules_ig1.add("ig3");
                }
                if (!rules_ig1.isEmpty())
                    control1.put("rule.control.ig", rules_ig1);
                if (rules_ig1.isEmpty()) {
                    control_flag3 = 1;
                    control1.put("rule.control.ig", new ArrayList<>());

                }

                System.out.println(row.getCell(3));
                System.out.println(row.getCell(22));

                System.out.println( control_flag1+" "+control_flag2+" "+control_flag3+" "+control_flag4);
                System.out.println("Control 1: "+control1);



                JSONObject control2 = new JSONObject();//control 2 obj
                int control_flag5 = 0;
                int control_flag6 = 0;
                int control_flag7 = 0;
                int control_flag8 = 0;


                if (!row.getCell(28).toString().equals("0.0") && !row.getCell(28).toString().equals("")) {

                    control2.put("rule.control.version", row.getCell(28));
                } else {
                    control_flag8 = 1;
                    control2.put("rule.control.version", "");
                }


                String input_title2 = row.getCell(19).toString();
                String title2;

                Pattern pattern_title2 = Pattern.compile("^TITLE:(.*)CONTROL:");
                Matcher matcher_title2 = pattern_title2.matcher(input_title2);

                if (matcher_title2.find()) {
                    title2 = matcher_title2.group(1).trim(); // Extract and trim the result
                } else {
                    title2 = "";
                }
                if (title.equals("")) {
                    control_flag5 = 1;
                }


                control2.put("rule.control.name", title2);
                String input = row.getCell(19).toString();
                //extract description
                String description;
                Pattern pattern = Pattern.compile("DESCRIPTION:(.*)");
                Matcher matcher = pattern.matcher(input);
                if (matcher.find()) {
                    description = matcher.group(1).trim(); // Extract and trim the result
                } else
                    description = "";
                if (description.equals("")) {
                    control_flag6 = 1;
                }

                control2.put("rule.control.description", description);
                List<String> rules_ig2 = new ArrayList<>();
                StringBuffer sb2 = new StringBuffer("");
                if (row.getCell(31).toString().equals("X")) {
                    rules_ig2.add("ig1");
                }
                if (row.getCell(32).toString().equals("X")) {
                    //rules_ig2[1] = "ig2";
                    rules_ig2.add("ig2");
                }
                if (row.getCell(33).toString().equals("X")) {
//                       rules_ig2[2] = "ig3";
                    rules_ig2.add("ig3");
                }
                if (!rules_ig2.isEmpty())
                    control2.put("rule.control.ig", rules_ig2);
                else {
                    control2.put("rule.control.ig", new ArrayList<>());
                    control_flag7 = 1;
                }

//                if (!(control_flag1 == 1 && control_flag2 == 1 && control_flag3 == 1 && control_flag4 == 1)) {
//                    rule_controls.put(control1);
//                }
//
//
//                if (!(control_flag5 == 1 && control_flag6 == 1 && control_flag7 == 1 && control_flag8 == 1)) {
//                    rule_controls.put(control1);
//                }

                 if(!(control_flag1 == 1 && control_flag2 == 1 && control_flag3 == 1 && control_flag4 == 1 )){
                    //jsonObject.put("rule.controls",new ArrayList<>());
                     rule_controls.put(control1);

                 }
                if(!(control_flag5 == 1 && control_flag6 == 1 && control_flag7 == 1 && control_flag8 == 1 )){
                    //jsonObject.put("rule.controls",new ArrayList<>());
                    rule_controls.put(control2);

                }
//                 else{
//                     rule_controls.put(control1);
//                     rule_controls.put(control2);
//                     jsonObject.put("rule.controls", rule_controls);
//            }



            jsonObject.put("rule.controls", rule_controls);}
            jsonObject.put("id", id++);
            jsonArray.put(jsonObject);

        //System.out.println(jsonArray);
            System.out.println(id);
        System.out.println("30th rowwwwww : "+sheet.getRow(i).getCell(4));
            System.out.println("30th rowwwwww : "+sheet.getRow(i).getCell(25));

            System.out.println("30th rowwwwww : "+row.getCell(26));
            System.out.println("30th rowwwwww : "+row.getCell(27));
            System.out.println("30th rowwwwww : "+row.getCell(18));



            FileWriter f=new FileWriter(new File("/home/jenil/IdeaProjects/json_project/output.json"));
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




}}



