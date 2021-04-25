package com.wps.examples;

import com.wps.api.tree.et.*;
import com.wps.api.tree.ex.IWindowEx;
import com.wps.runtime.utils.WpsArgs;
import com.wps.runtime.utils.Platform;
import com4j.Variant;
import sun.awt.WindowIDProvider;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.peer.ComponentPeer;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * Project              :   Java API Examples
 * Comments             :   WPS表格标签页面，包含一个左侧标签栏（LeftMenuPanel）和右侧的wps显示区域（OfficePanel）
 * JDK version used     :   JDK1.8
 * Author               :   Kingsoft
 * Create Date          :   2019-12-09
 * Version              :   1.0
 */

public class EtMainPanel extends JPanel {

    private LeftMenuPanel menuPanel;
    private OfficePanel officePanel;
    private Application app = null;
    private static final int DEFAULT_LCID = 2052; //默认的lcid,遇到有lcid的接口传这个就可以.

    public EtMainPanel() {
        this.setLayout(new BorderLayout());
        menuPanel = new LeftMenuPanel();
        officePanel = new OfficePanel();
        this.add(menuPanel, BorderLayout.WEST);
        this.add(officePanel, BorderLayout.CENTER);
        initMenu();
        initRibbon();
        initFrameListener();
    }

    private void initMenu(){
        menuPanel.addButton("常用", "初始化", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Canvas client = officePanel.getCanvas();
                if (app != null) {
                    JOptionPane.showMessageDialog(client, "已经初始化过，不需要重新初始化！");
                    return;
                }

                long nativeWinId = 0;

                try {
                    if (Platform.isWindows()) {
                        ComponentPeer peer = client.getPeer();
                        Class clsCanvasPeer = Class.forName("sun.awt.windows.WComponentPeer");
                        Method getHWnd = clsCanvasPeer.getDeclaredMethod("getHWnd");
                        getHWnd.setAccessible(true);

                        nativeWinId = (long)getHWnd.invoke(peer);
                    } else {
                        WindowIDProvider pid = (WindowIDProvider) client.getPeer();
                        nativeWinId = pid.getWindow();
                    }
                } catch (NoSuchMethodException | IllegalArgumentException | InvocationTargetException | ClassNotFoundException | IllegalAccessException ex) {
                    ex.printStackTrace();
                }

                WpsArgs args = WpsArgs.ARGS_MAP.get(WpsArgs.App.ET);
                args.setWinid(nativeWinId);
                args.setHeight(client.getHeight());
                args.setWidth(client.getWidth());
//                args.setCrypted(false); //wps2016需要关闭加密
                app = ClassFactory.createApplication();
                app.put_Visible(DEFAULT_LCID, true);
            }
        });


        menuPanel.addButton("常用", "创建新表格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Workbooks().Add(Variant.getMissing(), DEFAULT_LCID);
            }
        });

        menuPanel.addButton("常用", "关闭", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveWorkbook().Close(false, Variant.getMissing(), Variant.getMissing(), DEFAULT_LCID);
            }
        });

        menuPanel.addButton("常用", "获取版本", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String version = String.valueOf(app.get_Build(DEFAULT_LCID));
                JOptionPane.showMessageDialog(null, version);
            }
        });
    }


    private void initRibbon(){

        menuPanel.addButton("功能区", "隐藏/显示-功能区", new ActionListener() {
            private boolean visible = false;
            @Override
            public void actionPerformed(ActionEvent e) {
            	// windows企业版要2020年11月以后的版本才支持IWindowEx接口
                com.wps.api.tree.et.Window win = app.get_ActiveWindow();
                IWindowEx winEx = win.queryInterface(IWindowEx.class);
                winEx.put_ShowRibbon(visible);
                visible = !visible;
            }
        });

        menuPanel.addButton("功能区", "禁用/启用-剪切按钮", new ActionListener() {
            private boolean enabled = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                // windows企业版要2020年11月以后的版本才支持SetEnabledMso函数
                app.get_CommandBars().SetEnabledMso("Cut", enabled);
                enabled = !enabled;
            }
        });

        menuPanel.addButton("功能区", "隐藏/显示-剪切按钮", new ActionListener() {
            private boolean visible = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                // windows企业版要2020年11月以后的版本才支持SetVisibleMso函数
                app.get_CommandBars().SetVisibleMso("Cut", visible);
                visible = !visible;
            }
        });
    }

    private void initFrameListener() {
        if (!Platform.isWindows())
            return;

        final JPanel thiz = this;
        SwingUtilities.invokeLater(new Runnable() {

            @Override
            public void run() {
                JFrame p = (JFrame) SwingUtilities.getWindowAncestor(thiz);
                p.addWindowListener(new WindowAdapter() {

                    @Override
                    public void windowClosing(WindowEvent arg0) {
                        if (app != null) {
                            app.Quit();
                        }
                    }
                });
            }
        });
    }
}
