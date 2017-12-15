package com.bitmascot.Perform;

import com.bitmascot.File_functions.DirectoryFunctions;
import com.bitmascot.JSON_Reader.JSON2MAP;

import java.util.List;

public class Perform {

    private String jsonDir;
    private DirectoryFunctions directoryFunctions;
    private JSON2MAP json2MAP;

    public Perform() {
        this.jsonDir = "resources/json/";
        this.directoryFunctions = new DirectoryFunctions(jsonDir);
        this.json2MAP = new JSON2MAP();
    }

    public void performTask() {
        List<String> files = directoryFunctions.getallFiles();

        files.stream().forEach(file -> {
            try {
                json2MAP.saveToExcel(file, directoryFunctions.getdir(file));
            } catch (Throwable throwable) {
                throwable.printStackTrace();
            }
        });

    }

}
