package com.wps.examples;

import javax.swing.*;
import java.awt.*;

/**
 * Project              :   Java API Examples
 * Comments             :   承载Office的Panel，在窗体上创建一个画布装载WPS
 * JDK version used     :   JDK1.8
 * Author               :   Kingsoft
 * Create Date          :   2019-12-09
 * Version              :   1.0
 */
public class OfficePanel extends JPanel {

    private Canvas canvas = null;
    public OfficePanel() {
        this.setLayout(new BorderLayout());
        canvas = new Canvas() {
            public void paint(Graphics g) {
                super.paint(g);
            }
        };
        this.add(canvas, BorderLayout.CENTER);
    }

    public Canvas getCanvas(){
        return canvas;
    }
}
