package com.bitmascot.File_functions;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DirectoryFunctions {

    private String directLocation;

    public DirectoryFunctions(String directLocation) {
        this.directLocation = directLocation;
    }

    public List<String> getAllDirecotry() {
        File directoryLocation = new File(directLocation);
        List<String> directories = Arrays.asList(directoryLocation.list((dir, name) -> {
            return new File(dir, name).isDirectory();
        }));
        return directories;
    }

    public List<String> getallFiles(List<String> directories) {

        List<String> allFiles = new ArrayList<>();
        for (String dir : directories) {
            File[] files = new File(directLocation + dir).listFiles();
            for (File file : files) {
                if (file.isFile()) {
                    String filename = file.getName();
                    String extension = filename.substring(filename.lastIndexOf(".") + 1, filename.length());
                    if (extension.equalsIgnoreCase("json")) {
                        allFiles.add(directLocation + dir + "/" + file.getName());
                    }
                }
            }
        }

        return allFiles;
    }

    public List<String> getallFiles() {
        return getallFiles(getAllDirecotry());
    }

    public String getdir(String text) {
        String regex = "\\d{2}\\-[A-Z][a-z]{2}";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(text);
        if (matcher.find())
            return (matcher.group());
        else
            return null;
    }
}
