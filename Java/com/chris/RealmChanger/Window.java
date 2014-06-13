package com.chris.RealmChanger;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Vector;

public class Window extends JFrame {

    private final RealmList realmList = new RealmList();
    private final JLabel realmLoc = new JLabel("Realm list not set.");
    private final JButton btnChange = new JButton("Change Realm");
    private final JList<String> lst = new JList<>();
    private Vector<String> serverNames = new Vector<>();
    private final Map<String, String> servers =  new HashMap<>();
    private final JFileChooser fileChooser = new JFileChooser();
    private String arg;

    public Window(String arg) {
        this.arg = arg;
        servers.putAll(realmList.loadServersFromFile(arg));
        init();
    }

    private void init() {
        setTitle("Realm Changer");
        setBounds(50, 50, 300, 250);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setAlwaysOnTop(true);
        setResizable(false);
        setLocationRelativeTo(null); // Center window

        BorderLayout borderLayout = new BorderLayout(0, 0);

        JPanel panel = new JPanel(borderLayout);
        panel.setVisible(true);
        add(panel);

        lst.setFont(new Font("Calibri", Font.PLAIN, 16));
        servers.forEach( (key, value) -> serverNames.add(key));
        lst.setListData(serverNames);
        JScrollPane lstScroll = new JScrollPane(lst);
        panel.add(lstScroll, BorderLayout.NORTH);

        JLabel lTop = new JLabel("Available Server Realms");
        panel.add(lTop, BorderLayout.CENTER);

        JButton btnSeek = new JButton("Seek Realm");
        btnSeek.addActionListener(ae -> {
            fileChooser.setDialogTitle("Seek Realm List");
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Realm Lists", "wtf");
            fileChooser.setFileFilter(filter);
            int result = fileChooser.showOpenDialog(Window.this);
            if (result == JFileChooser.APPROVE_OPTION) {
                realmLoc.setText(fileChooser.getSelectedFile().getAbsolutePath());
                realmList.setRealmList(fileChooser.getSelectedFile());
                lst.setEnabled(true);
                btnChange.setEnabled(true);
                btnSeek.setEnabled(false);
            }
        });

        panel.add(btnSeek, BorderLayout.WEST);

        btnChange.addActionListener(ae -> {
            try {
                FileWriter fileWriter = new FileWriter(realmList.getRealmList(), false);
                fileWriter.write(servers.get(lst.getSelectedValue()));
                fileWriter.flush();
                fileWriter.close();
            }
            catch (IOException e) {
                JOptionPane.showMessageDialog(this, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        });
        panel.add(btnChange, BorderLayout.EAST);

        realmLoc.setFont(new Font("Calibri", Font.BOLD, 14));
        panel.add(realmLoc, BorderLayout.SOUTH);

        lst.setEnabled(false);
        btnChange.setEnabled(false);

        pack();
        setVisible(true);
    }

}
