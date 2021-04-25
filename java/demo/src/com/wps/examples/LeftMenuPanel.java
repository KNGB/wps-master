package com.wps.examples;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.util.HashMap;
import java.util.Map;

/**
 * Project              :   Java API Examples
 * Comments             :   界面左侧第一列标签栏，栏中每一个标签按钮对应一类接口，切换标签时，切换第二列标签栏中对应的接口按钮。维护一个Map，创建时标签数量和内容从Map中动态获取
 * JDK version used     :   JDK1.8
 * Author               :   Kingsoft
 * Create Date          :   2019-12-09
 * Version              :   1.0
 */
public class LeftMenuPanel extends JPanel {

    private JTabbedPane tabbedPane;
    private Map<String, JPanel> panels = new HashMap<>();
    public LeftMenuPanel() {
        this.setLayout(new BorderLayout());
        tabbedPane = new JTabbedPane();
        tabbedPane.setPreferredSize(new Dimension(350, 300));
        tabbedPane.setTabPlacement(JTabbedPane.LEFT);
        this.add(tabbedPane);
    }

    public void addButton(String category, String name, ActionListener listener, boolean enable){
        JPanel panel = null;
        if(panels.containsKey(category)){
            panel = panels.get(category);
        }else{
            panel = new JPanel();
            panel.setLayout(new GridLayout(30,1, 0, 0));
            tabbedPane.add(category, panel);
            panels.put(category, panel);
        }
        JButton button = new JButton(name);
        button.setEnabled(enable);
        button.addActionListener(listener);
        panel.add(button);
    }

    public void addButton(String category, String name, ActionListener listener){
        this.addButton(category,name,listener,true);
    }

}
