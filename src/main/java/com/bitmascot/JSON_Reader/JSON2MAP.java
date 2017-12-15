package com.bitmascot.JSON_Reader;

import com.bitmascot.Excel_Reader._Excel_Reader;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import java.io.FileReader;
import java.util.*;

/**
 * Created by torikul on 12/3/2017.
 */
public class JSON2MAP {


    private String excelLocation = "resources/excle/Staff_Tracking_Dec2017_Touhid.xlsx";

    public void saveToExcel(String pathname, String dirName) throws Throwable {
        JSONParser parser = new JSONParser();
        Object obj = parser.parse(new FileReader(pathname));
        JSONObject jsonObject = (JSONObject) obj;

        Map<String, Object> map = jsonToMap(jsonObject);

        //ALL ACTIVITIES
        Map<String, String> activities = (Map) map.get("activities");


        _Excel_Reader excel_reader = new _Excel_Reader(excelLocation);

        String sheetName = dirName;

        int mainStartRowIndex = excel_reader.getRowIndex(sheetName, "Resource(s)");
        int mainEndRowIndex = excel_reader.getRowIndex(sheetName, "Cumulative Report");

        int rowIndexofTeam = excel_reader.getRowIndexByFormulaValue(sheetName, mainStartRowIndex, mainEndRowIndex, (String) map.get("team"));

        int rowIndexOfTotal = excel_reader.getRowIndexByRagne(sheetName, rowIndexofTeam, mainEndRowIndex, "Total");

        int entryPresence = excel_reader.getRowIndexByFormulaValue(sheetName, rowIndexofTeam, rowIndexOfTotal, (String) map.get("name"));

        if (entryPresence == -1) {
            //name not found
            System.out.println("Name not found.");
        } else {
            //name found
            Set<String> keys = activities.keySet();
            for (String key : keys) {
                int indexofKey = excel_reader.findCloLoc(sheetName, getTaskname(key));
                excel_reader.setCellData(sheetName, indexofKey, entryPresence, Integer.parseInt(activities.get(key)), null);
            }
        }
        System.out.println("Data saved for " + (String) map.get("name") + " Team: " + (String) map.get("team"));

    }


    public Map<String, Object> jsonToMap(JSONObject json) {
        Map<String, Object> retMap = new HashMap<String, Object>();

        if (json != null) {
            retMap = toMap(json);
        }
        return retMap;
    }

    public Map<String, Object> toMap(JSONObject object) {
        Map<String, Object> map = new HashMap<String, Object>();

        Iterator<String> keysItr = object.keySet().iterator();
        while (keysItr.hasNext()) {
            String key = keysItr.next();
            Object value = object.get(key);

            if (value instanceof JSONArray) {
                value = toList((JSONArray) value);
            } else if (value instanceof JSONObject) {
                value = toMap((JSONObject) value);
            }
            map.put(key, value);
        }
        return map;
    }

    public List<Object> toList(JSONArray array) {
        List<Object> list = new ArrayList<Object>();
        for (int i = 0; i < array.size(); i++) {
            Object value = array.get(i);
            if (value instanceof JSONArray) {
                value = toList((JSONArray) value);
            } else if (value instanceof JSONObject) {
                value = toMap((JSONObject) value);
            }
            list.add(value);
        }
        return list;
    }

    public String getTaskname(String key) {
        String taskName = null;
        switch (key) {
            case "1":
                taskName = "1. Requirement Analysis";
                break;
            case "2":
                taskName = "2. Estimation";
                break;
            case "3":
                taskName = "3. Design";
                break;
            case "4":
                taskName = "4. Front-End Development";
                break;
            case "5":
                taskName = "5. Development";
                break;
            case "6":
                taskName = "6. Testing";
                break;
            case "7":
                taskName = "7. Bug Fixings";
                break;
            case "8":
                taskName = "8. Code Review";
                break;
            case "9":
                taskName = "9. Research";
                break;
            case "10":
                taskName = "10. Internal Communication";
                break;
            case "11":
                taskName = "11. Client Communication";
                break;
            case "12":
                taskName = "12. PRM";
                break;
            case "13":
                taskName = "13. DevOps";
                break;
            default:
                taskName = "not found";
        }
        return taskName;
    }

}
