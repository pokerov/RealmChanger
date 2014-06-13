package com.chris.RealmChanger;

import javax.swing.*;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class RealmList {

    private File realmList;

    public void setRealmList(File realmList) {
        if (realmList == null) return;
        this.realmList = realmList;
    }

    public File getRealmList() {
        return realmList;
    }

    public Map<String, String> loadServersFromFile(String fileName) {
        Map<String, String> sList = new HashMap<>();
        try {
            File f = new File(fileName);
            char[] data = new char[(int)f.length()];
            FileReader fileReader = new FileReader(f);
            fileReader.read(data);
            fileReader.close();

            String line = new String(data);
            String[] lineArray = line.split("\\n");

            for (int i = 0; i <= lineArray.length - 1; i += 3) {
                sList.put(lineArray[i], lineArray[i+1] + "\n" + lineArray[i+2]);
            }
        }
        catch (IOException e) {
            handleError(e.getMessage());
            System.exit(0);
        }
        return sList;
    }

    private void handleError(String message) {
        JOptionPane.showMessageDialog(null, message, "Error", JOptionPane.ERROR_MESSAGE);
    }
}
