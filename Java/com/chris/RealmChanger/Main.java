package com.chris.RealmChanger;


import javax.swing.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        List<String> lst = new ArrayList<>(Arrays.asList(args));
        if (lst.contains("servers.txt")) new Window("servers.txt");
        else JOptionPane.showMessageDialog(null, "No present server list.", "Error", JOptionPane.ERROR_MESSAGE);
    }
}
