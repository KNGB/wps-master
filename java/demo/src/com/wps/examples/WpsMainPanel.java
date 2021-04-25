package com.wps.examples;

import com.wps.api.tree.ex.WindowEx;
import com.wps.api.tree.kso.*;
import com.wps.api.tree.kso.events._CommandBarButtonEvents;
import com.wps.api.tree.wps.*;
import com.wps.api.tree.wps.ClassFactory;
import com.wps.api.tree.wps.events.ApplicationEvents4;
import com.wps.api.tree.wps.KsoOfdServiceProviderType;
import com.wps.api.tree.ex._ApplicationEx;
import com.wps.api.tree.ex._DocumentEx;
import com.wps.runtime.utils.WpsArgs;
import com.wps.runtime.utils.Platform;
import com4j.EventCookie;
import com4j.Holder;
import com4j.Variant;

import org.apache.commons.codec.binary.Base64;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import sun.awt.WindowIDProvider;

import javax.swing.*;
import java.awt.*;
import java.awt.FileDialog;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.peer.ComponentPeer;
import java.io.File;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;


/**
 * Project              :   Java API Examples
 * Comments             :   WPS文字标签页面，包含一个左侧标签栏（LeftMenuPanel）和右侧的wps显示区域（OfficePanel）
 * JDK version used     :   JDK1.8
 * Author               :   Kingsoft
 * Create Date          :   2019-12-09
 * Version              :   1.0
 */
public class WpsMainPanel extends JPanel {

    private LeftMenuPanel menuPanel;
    private OfficePanel officePanel;
    private Application app = null;
    private EventCookie cookie = null;
    private EventCookie cookie2 = null;
    private static Logger logger = LoggerFactory.getLogger(WpsMainPanel.class);
    private CommandBarControl control = null;

    public WpsMainPanel() {
        this.setLayout(new BorderLayout());
        menuPanel = new LeftMenuPanel();
        officePanel = new OfficePanel();
        this.add(menuPanel, BorderLayout.WEST);
        this.add(officePanel, BorderLayout.CENTER);
        initNormalMenu();           //常用接口
        initLocalMenu();            //本地文档相关接口
        initOFDPDFMenu();           //ofd和pdf相关接口
        initMenu();                 //菜单栏操作接口
        initRibbon();               // Ribbon风格界面
        initRepairedMenu();         //修订接口
        initFieldsMenu();           //公文域操作接口
        initSignMenu();             //手写批注相关接口
        initEventMenu();            //事件监听回调接口
        initOthersMenu();           //其他
        initFrameListener();        //注册窗口关闭事件处理
    }

    public static String getPath(String title, int type){
        FileDialog dialog = new FileDialog((JFrame)null, title, type);
        dialog.setVisible(true);
        if(dialog.getDirectory() == null || dialog.getFile() == null)
            throw new RuntimeException("选择的文件不能为空!");
        return dialog.getDirectory() + dialog.getFile();
    }

    private void initNormalMenu(){
        menuPanel.addButton("常用", "初始化", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Canvas client = officePanel.getCanvas();
                if (app != null) {
                    if(null != app.get_ActiveWindow()){
                        JOptionPane.showMessageDialog(client, "已经初始化过，不需要重新初始化！");
                        return;
                    }
                }

                long nativeWinId = 0;
                try {
                    if (Platform.isWindows()) {
                    	ComponentPeer peer = client.getPeer();
                    	Class<?> clsCanvasPeer = Class.forName("sun.awt.windows.WComponentPeer");
                    	Method getHWnd = clsCanvasPeer.getDeclaredMethod("getHWnd");
                    	getHWnd.setAccessible(true);

                        nativeWinId = (long)getHWnd.invoke(peer);
                    } else {
                        WindowIDProvider pid = (WindowIDProvider) client.getPeer();

                        Class<?> clsCanvasPeer = Class.forName("sun.awt.X11.XEmbedCanvasPeer");
                        Method removeXEmbedDropTarget = clsCanvasPeer.getDeclaredMethod("removeXEmbedDropTarget");
                        removeXEmbedDropTarget.setAccessible(true);
                        removeXEmbedDropTarget.invoke(pid);

                        Method detachChiled = clsCanvasPeer.getDeclaredMethod("detachChild");
                        detachChiled.setAccessible(true);
                        Method isXEmbedActive = clsCanvasPeer.getDeclaredMethod("isXEmbedActive");
                        isXEmbedActive.setAccessible(true);
                        Boolean isActive = (Boolean) isXEmbedActive.invoke(pid);
                        if(isActive){
                            detachChiled.invoke(pid);
                        }

                        nativeWinId = pid.getWindow();
                    }
                } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException | ClassNotFoundException ex) {
                    ex.printStackTrace();
                }

                WpsArgs args = WpsArgs.ARGS_MAP.get(WpsArgs.App.WPS);
//              args.setPath("/home/wps/workspace/wps_2016/build_debug/WPSOffice/office6/wps"); //手动指定的wps程序路径（默认调用/usr/bin/wps）
                args.setWinid(nativeWinId);
                args.setHeight(client.getHeight());
                args.setWidth(client.getWidth());
//                args.setCrypted(false); //wps2016需要关闭加密
                app = ClassFactory.createApplication();
                app.put_Visible(true);
            }
        });


        menuPanel.addButton("常用", "创建新文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Documents().Add(Variant.getMissing(), Variant.getMissing(), Variant.getMissing(), Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "获取版本", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String version = app.get_Build();
                JOptionPane.showMessageDialog(null, version);
            }
        });

        menuPanel.addButton("常用", "关闭", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().Close(false, Variant.getMissing(), Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "打印文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                ((Document)app.get_ActiveDocument()).PrintOut(
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "在当前光标处插入图片", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().get_InlineShapes().AddPicture(getPath("请选择一张图片", FileDialog.LOAD),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "当前光标处插入图片_浮于文字上方", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Document pDoc = app.get_ActiveDocument();
                String imagePath =getPath("请选择一张图片", FileDialog.LOAD);
                ImageIcon image = new ImageIcon(imagePath);
                double width = image.getIconWidth();
                double height = image.getIconHeight();
                double left = Double.parseDouble(app.get_Selection().get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary).toString());
                double top = Double.parseDouble(app.get_Selection().get_Information(WdInformation.wdVerticalPositionRelativeToTextBoundary).toString());
                int page = Integer.parseInt(app.get_Selection().get_Information(WdInformation.wdActiveEndPageNumber).toString());

                Object range = pDoc.GoTo(WdGoToItem.wdGoToPage,WdGoToDirection.wdGoToAbsolute,page,Variant.getMissing());
                Variant anchor = new Variant();
                anchor.set(range);
                double pageTop = 0.0;
                View pView = app.get_ActiveWindow().get_View();
                pView.get_DisplayPageBoundaries();
                if(pView != null && !pView.get_DisplayPageBoundaries()){
                    PageSetup pPageSetup = pDoc.get_PageSetup();
                    pageTop = pPageSetup.get_TopMargin();
                }
                pDoc.get_Shapes().AddPicture(
                        imagePath,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        left,
                        top - pageTop,
                        width,
                        height,
                        anchor
                );
            }
        });

        menuPanel.addButton("常用", "插入图片_带尺寸坐标", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Document pDoc = app.get_ActiveDocument();
                double pageTop = 0.0;
                View pView = app.get_ActiveWindow().get_View();
                pView.get_DisplayPageBoundaries();
                if(pView != null && !pView.get_DisplayPageBoundaries()){
                    PageSetup pPageSetup = pDoc.get_PageSetup();
                    pageTop = pPageSetup.get_TopMargin();
                }
                pDoc.get_Shapes().AddPicture(getPath("请选择一张图片", FileDialog.LOAD), //10, 20, 80, 100
                        Variant.getMissing(),
                        Variant.getMissing(),
                        (double)10,
                        (double)20 - pageTop,
                        80,
                        100,
                        Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "获取文本内容", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String text = app.get_ActiveDocument().get_Content().get_Text();
                text.replace("/[\r]/g","\r\n");
                JOptionPane.showMessageDialog(null, "<html><body><p style='width: 200px;'>"+text+"</p></body></html>");//文本内容折行显示
            }
        });

        menuPanel.addButton("常用", "设置审批用户名", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.put_UserName("shenpiyonghu123");
            }
        });

        menuPanel.addButton("常用", "获取审批用户名", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JOptionPane.showMessageDialog(null, "<html><body><p style='width: 200px;'>"+app.get_UserName()+"</p></body></html>");
            }
        });

    }

    private  void initLocalMenu(){
        menuPanel.addButton("本地文档", "打开可编辑本地文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Documents().Open(getPath("请选择一个文字文件", FileDialog.LOAD),
                        Variant.getMissing(),
                        false,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("本地文档", "打开可编辑本地文档_通过文档对话框", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc=new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.showDialog(new JLabel(), "打开");
                String fileName = jfc.getSelectedFile().getPath();

                app.get_Documents().Open(fileName,
                        Variant.getMissing(),
                        false,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("本地文档", "只读模式打开本地文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Documents().Open(getPath("请选择一个文字文件", FileDialog.LOAD),
                        Variant.getMissing(),
                        true,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                ).Protect(WdProtectionType.wdAllowOnlyReading,Variant.getMissing(),
                        Variant.getMissing(),Variant.getMissing(),Variant.getMissing());
            }
        });


        menuPanel.addButton("本地文档", "保存本地弹窗", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                HashMap formatMap = new HashMap<String,Integer>();
                formatMap.put("doc",0);
                formatMap.put("wps",0);
                formatMap.put("wpt",1);
                formatMap.put("dot",1);
                formatMap.put("txt",7);
                formatMap.put("rtf",6);
                formatMap.put("html",8);
                formatMap.put("htm",8);
                formatMap.put("mhtml",9);
                formatMap.put("xml",11);
                formatMap.put("docx",12);
                formatMap.put("docm",13);
                formatMap.put("dotx",14);
                formatMap.put("dotm",15);
                formatMap.put("uof",100);
                formatMap.put("uot",101);
                formatMap.put("ofd",102);
                formatMap.put("pdf",17);

                //需要打开本地文件对话框
                JFileChooser jfc=new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.showDialog(new JLabel(), "另存为");
                String fileName = jfc.getSelectedFile().getPath();
                String[] formats = fileName.split("\\.");
                String format;
                if(formats.length == 0 || !formatMap.containsKey(formats[formats.length - 1])) { //路径不存在.或者后缀名无法识别，默认保存为.wps
                    format = "wps";
                    fileName += ".wps";
                }
                else {
                    format = formats[formats.length - 1];
                }
                // 获取到文件名保存
                Document pDoc = app.get_ActiveDocument();
                pDoc.SaveAs(fileName,
                        formatMap.get(format),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("本地文档", "打开隐藏文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Documents().Open(getPath("请选择一个文字文件", FileDialog.LOAD), false, false, false, "", "", false, "", "", 0, 0, false, false, false, 0, false);
            }
        });

    }

    private void initMenu(){

        menuPanel.addButton("菜单栏", "隐藏输出ofd和pdf按钮", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                CommandBar commandBar = app.get_CommandBars().get_Item("file");
                if(commandBar == null){
                    logger.error("CommandBar is null!");
                    return ;
                }
                CommandBarControl controlPdf = commandBar.get_Controls().get_Item(25);
                if(controlPdf == null){
                    logger.error("ControlPdf is null!");
                    return ;
                }
                controlPdf.put_Visible(false);

                CommandBarControl controlOfd = commandBar.get_Controls().get_Item(26);
                if(controlOfd == null){
                    logger.error("ControlOfd is null!");
                    return ;
                }
                controlOfd.put_Caption("输出为OFD格式(&G)...");
                controlOfd.put_Visible(false);
            }
        });

        menuPanel.addButton("菜单栏", "隐藏文件菜单", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                CommandBar commandBar = app.get_CommandBars().get_Item("Menu Bar");
                if(commandBar == null){
                    logger.error("CommandBar is null!");
                    return ;
                }
                CommandBarControl control = commandBar.get_Controls().get_Item(1);
                if(control == null){
                    logger.error("CommandBarControl is null!");
                    return ;
                }
                control.put_Visible(false);
            }
        });

        menuPanel.addButton("菜单栏", "禁用/启用-接受修订按钮", new ActionListener() {
            private boolean enable = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                enable = !enable;
                app.get_CommandBars().get_Item("Track Changes").get_Controls().get_Item(7).put_Enabled(enable);
                app.get_CommandBars().get_Item("Reviewing").get_Controls().get_Item(5).put_Enabled(enable);
            }
        });

        menuPanel.addButton("菜单栏", "禁用/启用-拒绝修订按钮", new ActionListener() {
            private boolean enable = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                enable = !enable;
                app.get_CommandBars().get_Item("Track Changes").get_Controls().get_Item(8).put_Enabled(enable);
                app.get_CommandBars().get_Item("Reviewing").get_Controls().get_Item(6).put_Enabled(enable);

            }
        });

        menuPanel.addButton("菜单栏", "显示工具条", new ActionListener() {
            private HashMap<String,Integer> m_BarCatchMap = new HashMap<String, Integer>();
            @Override
            public void actionPerformed(ActionEvent e) {
                m_BarCatchMap.put("Standard", 0);
                m_BarCatchMap.put("Formatting", 0);
                m_BarCatchMap.put("Tables and Borders", 0);
                m_BarCatchMap.put("Align", 0);
                m_BarCatchMap.put("Reviewing", 0);
                m_BarCatchMap.put("Extended Formatting", 0);
                m_BarCatchMap.put("Drawing", 0);
                m_BarCatchMap.put("Symbol", 0);
                m_BarCatchMap.put("3-D Settings", 0);
                m_BarCatchMap.put("Shadow Settings", 0);
                m_BarCatchMap.put("Mail Merge", 0);
                m_BarCatchMap.put("Control Toolbox", 0);
                m_BarCatchMap.put("Outlining", 0);
                m_BarCatchMap.put("Forms", 0);
                app.get_CommandBars().get_Item("menu bar").put_Enabled(true);

                for (String key : m_BarCatchMap.keySet()) {
                    app.get_CommandBars().get_Item(key).put_Enabled(true);
                }
            }
        });

        menuPanel.addButton("菜单栏", "隐藏工具条", new ActionListener() {
            private HashMap<String,Integer> m_BarCatchMap = new HashMap<String, Integer>();
            @Override
            public void actionPerformed(ActionEvent e) {
                m_BarCatchMap.put("Standard", 0);
                m_BarCatchMap.put("Formatting", 0);
                m_BarCatchMap.put("Tables and Borders", 0);
                m_BarCatchMap.put("Align", 0);
                m_BarCatchMap.put("Reviewing", 0);
                m_BarCatchMap.put("Extended Formatting", 0);
                m_BarCatchMap.put("Drawing", 0);
                m_BarCatchMap.put("Symbol", 0);
                m_BarCatchMap.put("3-D Settings", 0);
                m_BarCatchMap.put("Shadow Settings", 0);
                m_BarCatchMap.put("Mail Merge", 0);
                m_BarCatchMap.put("Control Toolbox", 0);
                m_BarCatchMap.put("Outlining", 0);
                m_BarCatchMap.put("Forms", 0);
                app.get_CommandBars().get_Item("menu bar").put_Enabled(false);

                for (String key : m_BarCatchMap.keySet()) {
                    app.get_CommandBars().get_Item(key).put_Enabled(false);
                }
            }
        });
    }

    private void initOFDPDFMenu(){
        menuPanel.addButton("ofd和pdf", "本地导出ofd", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                ((Document)app.get_ActiveDocument()).SaveAs(
                        getPath("保存", FileDialog.SAVE),
                        102,            //保存格式， 102-->>ofd
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("ofd和pdf", "ofd厂商设置为Suwell", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_OfdExportOptions().put_SelectServiceProvider(KsoOfdServiceProviderType.KsoOfdServiceProviderSuwell);
            }
        });
        menuPanel.addButton("ofd和pdf", "ofd厂商设置为Foxit", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_OfdExportOptions().put_SelectServiceProvider(KsoOfdServiceProviderType.KsoOfdServiceProviderFoxit);
            }
        });
     }

    private void initRepairedMenu(){
        menuPanel.addButton("修订", "开启/关闭修订", new ActionListener() {
            private boolean m_status = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                m_status = !m_status;
                app.get_ActiveDocument().put_TrackRevisions(m_status);
            }
        });

        menuPanel.addButton("修订", "显示标记的最终状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(true);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewFinal);

            }
        });

        menuPanel.addButton("修订", "显示原始状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(false);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewOriginal);

            }
        });

        menuPanel.addButton("修订", "显示最终状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(false);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewFinal);

            }
        });

        menuPanel.addButton("修订", "打印修订", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(true);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewFinal);
                ((Document)app.get_ActiveDocument()).PrintOut(
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("修订", "打印原始状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(false);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewOriginal);
                ((Document)app.get_ActiveDocument()).PrintOut(
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("修订", "打印最终状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(false);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewFinal);
                ((Document)app.get_ActiveDocument()).PrintOut(
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("修订", "保护文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().Protect(WdProtectionType.wdAllowOnlyReading,Variant.getMissing(),Variant.getMissing(),Variant.getMissing(),Variant.getMissing());
            }
        });

        menuPanel.addButton("修订", "停止保护", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().Unprotect(Variant.getMissing());
            }
        });

        menuPanel.addButton("修订", "禁止剪切", new ActionListener() {
            private HashMap<String,Integer> m_enableCutMap = new HashMap<String, Integer>();
            @Override
            public void actionPerformed(ActionEvent e) {
                //cut
                m_enableCutMap.put("Chart Area Popup Menu", 2);
                m_enableCutMap.put("Dml Shapes", 2);
                m_enableCutMap.put("WordArt Context Menu", 1);
                m_enableCutMap.put("Tables", 1);
                m_enableCutMap.put("Track Changes", 1);
                m_enableCutMap.put("Inline Picture", 1);
                m_enableCutMap.put("Lists", 1);
                m_enableCutMap.put("Frames", 1);
                m_enableCutMap.put("Field Display List Numbers", 1);
                m_enableCutMap.put("Fields", 1);
                m_enableCutMap.put("Shapes", 1);
                m_enableCutMap.put("Text", 1);
                m_enableCutMap.put("Spelling Suggestions Popup Menu", 5);
                m_enableCutMap.put("Edit", 3);
                m_enableCutMap.put("Table Text", 1);
                m_enableCutMap.put("Table Cells", 1);
                m_enableCutMap.put("Whole Table", 1);
                m_enableCutMap.put("Picture Context Menu", 2);
                m_enableCutMap.put("Numbered Popup Menu", 2);
                m_enableCutMap.put("Bulleted Popup Menu", 2);
                m_enableCutMap.put("Standard", 11);

                for (String key : m_enableCutMap.keySet()) {
                    app.get_CommandBars().get_Item(key).get_Controls().get_Item(m_enableCutMap.get(key)).put_Enabled(false);
                }
            }
        });

        menuPanel.addButton("修订", "禁止复制", new ActionListener() {
            private HashMap<String,Integer> m_enableCopyMap = new HashMap<String, Integer>();
            @Override
            public void actionPerformed(ActionEvent e) {
                //copy
                m_enableCopyMap.put("Chart Area Popup Menu", 1);
                m_enableCopyMap.put("Dml Shapes", 1);
                m_enableCopyMap.put("WordArt Context Menu", 2);
                m_enableCopyMap.put("Tables", 2);
                m_enableCopyMap.put("Track Changes", 2);
                m_enableCopyMap.put("Inline Picture", 2);
                m_enableCopyMap.put("Lists", 2);
                m_enableCopyMap.put("Frames", 2);
                m_enableCopyMap.put("Field Display List Numbers", 1);
                m_enableCopyMap.put("Fields", 2);
                m_enableCopyMap.put("Shapes", 2);
                m_enableCopyMap.put("Text", 2);
                m_enableCopyMap.put("Spelling Suggestions Popup Menu", 4);
                m_enableCopyMap.put("Edit", 4);
                m_enableCopyMap.put("Table Text", 2);
                m_enableCopyMap.put("Table Cells", 2);
                m_enableCopyMap.put("Whole Table", 2);
                m_enableCopyMap.put("Picture Context Menu", 1);
                m_enableCopyMap.put("Numbered Popup Menu", 1);
                m_enableCopyMap.put("Bulleted Popup Menu", 1);
                m_enableCopyMap.put("Standard", 12);

                for (String key : m_enableCopyMap.keySet()) {
                    app.get_CommandBars().get_Item(key).get_Controls().get_Item(m_enableCopyMap.get(key)).put_Enabled(false);
                }
            }
        });

    }

    private void initFieldsMenu(){

        menuPanel.addButton("公文域", "插入一个公文域", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // WordBasic接口未支持QueryInterface，只能用Com4jObject表示，然后通过invoke调用函数。
                // 但java上WordBasic接口定义未列出函数的DispID，尚不能调用
                // Com4jObject wordBasic = app.get_WordBasic();
                WordBasic wordBasic = (WordBasic) app.get_WordBasic();
                if(wordBasic == null){
                    logger.error("WordBasic is null");
                    return ;
                }
                wordBasic.insertDocumentField("wps1");
            }
        });

        menuPanel.addButton("公文域", "插入多个公文域", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                WordBasic wordBasic = (WordBasic) app.get_WordBasic();
                if(wordBasic == null){
                    logger.error("WordBasic is null");
                    return ;
                }
                wordBasic.insertMultiDocumentField("份号,密级,保密期限,紧急程度,后缀标志,发文机关名称,发文机关代字,年份,发文顺序号,签发人_1,签发人_2,标题,主送机关,正文,发文机关署名,成文日期,抄送机关_1,抄送机关_2,抄送机关_3,抄送机关_4,抄送机关_5,印发机关,印发日期", "，",",");
            }
        });

        menuPanel.addButton("公文域", "设置公文域底纹显示_函数调用", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).put_ShowDocumentFieldTarget(1);
            }
        });

        menuPanel.addButton("公文域", "设置公文域底纹隐藏_函数调用", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).put_ShowDocumentFieldTarget(0);
            }
        });

        menuPanel.addButton("公文域", "判断公文域底纹状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Document pDoc = app.get_ActiveDocument();
                if(pDoc == null){
                    logger.error("Document is null");
                    return ;
                }
                _DocumentEx pDocEx = pDoc.queryInterface(_DocumentEx.class);
                if(pDocEx == null){
                    logger.error("_DocumentEx is null");
                    return ;
                }
                JOptionPane.showMessageDialog(null,pDocEx.get_ShowDocumentFieldTarget());
            }
        });

        menuPanel.addButton("公文域", "获取公文域列表", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                WordBasic wordBasic = (WordBasic) app.get_WordBasic();
                if(wordBasic == null){
                    logger.error("WordBasic is null");
                    return ;
                }
                JOptionPane.showMessageDialog(null, wordBasic.getAllDocumentField());
            }
        });

        menuPanel.addButton("公文域", "公文域删除", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_DocumentFields().Item("wps1").Delete(true);
            }
        });

        menuPanel.addButton("公文域", "公文域显示", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_DocumentFields().Item("wps1").put_Hidden(false);
                officePanel.grabFocus();
            }
        });

        menuPanel.addButton("公文域", "公文域不显示", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_DocumentFields().Item("wps1").put_Hidden(true);
            }
        });

        menuPanel.addButton("公文域", "查询公文域内容", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String value = app.get_ActiveDocument().get_DocumentFields().Item("wps1").get_Value();
                JOptionPane.showMessageDialog(null, "<html><body><p style='width: 200px;'>"+value+"</p></body></html>");
            }
        });

        menuPanel.addButton("公文域", "公文域可编辑", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_DocumentFields().Item("wps1").put_ReadOnly(false);
            }
        });

        menuPanel.addButton("公文域", "公文域不可编辑", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_DocumentFields().Item("wps1").put_ReadOnly(true);
            }
        });

        menuPanel.addButton("公文域", "公文域插入文档内容", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_DocumentFields().Item("wps1").InsertDocument(getPath("请选择一个文字文件", FileDialog.LOAD));
            }
        });

        menuPanel.addButton("公文域", "判断公文域是否存在", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if(app.get_ActiveDocument().get_DocumentFields().Exists("wps1")){;
                    JOptionPane.showMessageDialog(null,true);
                }else{
                    JOptionPane.showMessageDialog(null,false);
                }
            }
        });

        menuPanel.addButton("公文域", "删除(backspace)", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().TypeBackspace();
            }
        });

        menuPanel.addButton("公文域", "光标移动到公文域的指定位置", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                WordBasic wordBasic = (WordBasic) app.get_WordBasic();
                if(wordBasic == null){
                    logger.error("WordBasic is null");
                    return ;
                }
                wordBasic.cursorToDocumentField("wps1",4);
            }
        });

        menuPanel.addButton("公文域", "光标选中指定公文域", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                WordBasic wordBasic = (WordBasic) app.get_WordBasic();
                if(wordBasic == null){
                    logger.error("WordBasic is null");
                    return ;
                }
                wordBasic.cursorToDocumentField("wps1",5);
            }
        });

        menuPanel.addButton("公文域", "光标位置插入文本", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().TypeText("Kingsoft");
            }
        });

        menuPanel.addButton("公文域", "设置公文标识", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                DocumentProperties properties = (DocumentProperties)app.get_ActiveDocument().get_CustomDocumentProperties();
                DocumentProperty property = properties.get_Item("公文标识",0);
                if (null == property){
                    properties.Add("公文标识",false,Variant.getMissing(),"kingsoft",Variant.getMissing(),0);
                }else{
                    property.put_Value(0,"kingsoft");
                }
            }
        });

        menuPanel.addButton("公文域", "获取公文标识", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                DocumentProperties properties = (DocumentProperties)app.get_ActiveDocument().get_CustomDocumentProperties();
                if(properties == null){
                    logger.error("DocumentProperties is null");
                    return ;
                }
                JOptionPane.showMessageDialog(null,properties.get_Item("公文标识",0).get_Value(0).toString());
            }
        });

        menuPanel.addButton("公文域", "设置文种", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                DocumentProperties properties = (DocumentProperties)app.get_ActiveDocument().get_CustomDocumentProperties();
                DocumentProperty property = properties.get_Item("文种",0);
                if (null == property){
                    properties.Add("文种",false,Variant.getMissing(),"纪要",Variant.getMissing(),0);
                }else{
                    property.put_Value(0,"纪要");
                }
            }
        });

        menuPanel.addButton("公文域", "获取文种", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                DocumentProperties properties = (DocumentProperties)app.get_ActiveDocument().get_CustomDocumentProperties();
                JOptionPane.showMessageDialog(null,properties.get_Item("文种",0).get_Value(0).toString());
            }
        });

        menuPanel.addButton("公文域", "隐藏多个公文域", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                DocumentFields fields = app.get_ActiveDocument().get_DocumentFields();
                if(fields == null){
                    logger.error("DocumentFields is null");
                    return ;
                }
                String split = ",";
                String id = "份号,密级,保密期限,紧急程度,后缀标志,发文机关名称,发文机关代字,年份,发文顺序号,签发人_1,签发人_2,标题,主送机关,正文,发文机关署名,成文日期,抄送机关_1,抄送机关_2,抄送机关_3,抄送机关_4,抄送机关_5,印发机关,印发日期";
                String [] ids = id.split(split);
                for(int i = 0; i < ids.length; i++){
                    String name = ids[i];
                    fields.Item(name).put_Hidden(true);
                }
            }
        });

        menuPanel.addButton("公文域", "显示多个公文域", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                DocumentFields fields = app.get_ActiveDocument().get_DocumentFields();
                if(fields == null){
                    logger.error("DocumentFields is null");
                    return ;
                }
                String split = ",";
                String id = "份号,密级,保密期限,紧急程度,后缀标志,发文机关名称,发文机关代字,年份,发文顺序号,签发人_1,签发人_2,标题,主送机关,正文,发文机关署名,成文日期,抄送机关_1,抄送机关_2,抄送机关_3,抄送机关_4,抄送机关_5,印发机关,印发日期";
                String [] ids = id.split(split);
                for(int i = 0; i < ids.length; i++){
                    String name = ids[i];
                    fields.Item(name).put_Hidden(false);
                }
            }
        });

        menuPanel.addButton("公文域", "删除多个公文域", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                DocumentFields fields = app.get_ActiveDocument().get_DocumentFields();
                if(fields == null){
                    logger.error("DocumentFields is null");
                    return ;
                }
                String split = ",";
                String id = "份号,密级,保密期限,紧急程度,后缀标志,发文机关名称,发文机关代字,年份,发文顺序号,签发人_1,签发人_2,标题,主送机关,正文,发文机关署名,成文日期,抄送机关_1,抄送机关_2,抄送机关_3,抄送机关_4,抄送机关_5,印发机关,印发日期";
                String [] ids = id.split(split);
                for(int i = 0; i < ids.length; i++){
                    String name = ids[i];
                    fields.Item(name).Delete(true);
                }
            }
        });

        menuPanel.addButton("公文域", "修改公文域内容", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                WordBasic wordBasic = (WordBasic) app.get_WordBasic();
                if(wordBasic == null){
                    logger.error("WordBasic is null");
                    return ;
                }
                String[] fields = wordBasic.getAllDocumentField().split(",");
                String names = "";
                String values = "";
                for(int i = 0; i < fields.length; i++){
                    names += fields[i];
                    if(fields[i].equals("正文")){
                        values += "正文";
                    }else{
                        values += String.valueOf(i+1);
                    }
                    if(i != fields.length -1){
                        names += "@#_*@";
                        values += "@#_*@";
                    }
                }
                try {
                    wordBasic.setMultiDocumentField(names, new String(Base64.encodeBase64(values.getBytes("UTF-8")), "UTF-8"), true, "@#_*@");
                } catch (UnsupportedEncodingException ex) {
                    ex.printStackTrace();
                }
            }
        });

        menuPanel.addButton("公文域", "修改公文域内容_除正文", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                WordBasic wordBasic = (WordBasic) app.get_WordBasic();
                if(wordBasic == null){
                    logger.error("WordBasic is null");
                    return ;
                }
                String[] fields = wordBasic.getAllDocumentField().split(",");
                String names = "";
                String values = "";
                for(int i = 0; i < fields.length; i++){
                    if(fields[i].equals("正文")){
                        continue;
                    }else{
                        values += String.valueOf(i+100);
                    }
                    names += fields[i];
                    if(i != fields.length -1){
                        names += "@#_*@";
                        values += "@#_*@";
                    }
                }
                try {
                    wordBasic.setMultiDocumentField(names, new String(Base64.encodeBase64(values.getBytes("UTF-8")), "UTF-8"), true, "@#_*@");
                } catch (UnsupportedEncodingException ex) {
                    ex.printStackTrace();
                }
            }
        });


    }

    private void initSignMenu(){
        //手写批签
        menuPanel.addButton("手写批注", "隐藏手写签批", new ActionListener() {

            @Override
            public void actionPerformed(ActionEvent e) {
                CommandBar commandBar = app.get_CommandBars().get_Item("Reviewing");
                if(commandBar == null){
                    logger.error("CommandBar is null");
                    return ;
                }
                commandBar.get_Controls().get_Item(15).put_Visible(false);
            }
        });

        menuPanel.addButton("手写批注", "进入手写签批状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // windows版本仅2020年11月之后版本才支持EnterInkDraw
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).EnterInkDraw();
            }
        });

        menuPanel.addButton("手写批注", "退出手写签批状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // windows版本仅2020年11月之后版本才支持ExitInkDraw
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).ExitInkDraw();
            }
        });

        menuPanel.addButton("手写批注", "设置签批颜色", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // windows版本调用需标注方法 DISPID或VTID
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawColor("#BBFFFF");
            }
        });

        menuPanel.addButton("手写批注", "设置使用笔", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawStyle(0);
            }
        });

        menuPanel.addButton("手写批注", "设置批注笔类型_圆珠笔", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawPenStyle(0);
            }
        });

        menuPanel.addButton("手写批注", "设置批注笔类型_水彩笔", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawPenStyle(1);
            }
        });

        menuPanel.addButton("手写批注", "设置批注笔类型_荧光笔", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawPenStyle(2);
            }
        });

        menuPanel.addButton("手写批注", "设置使用形状", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawStyle(1);
            }
        });

        menuPanel.addButton("手写批注", "设置批注形状类型_直线", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawShapeStyle(0);
            }
        });

        menuPanel.addButton("手写批注", "设置批注形状类型_波浪线", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawShapeStyle(1);
            }
        });

        menuPanel.addButton("手写批注", "设置批注形状类型_矩形", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawShapeStyle(2);
            }
        });

        menuPanel.addButton("手写批注", "设置批注磅值_0.25", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawLineStyle(0);
            }
        });

        menuPanel.addButton("手写批注", "设置批注磅值_4", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawLineStyle(4);
            }
        });

        menuPanel.addButton("手写批注", "设置批注磅值_8", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawLineStyle(8);
            }
        });

        menuPanel.addButton("手写批注", "进入橡皮擦", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawEraser(true);
            }
        });

        menuPanel.addButton("手写批注", "退出橡皮擦", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().queryInterface(_DocumentEx.class).SetInkDrawEraser(false);
            }
        });

    }

    private void initEventMenu(){
        //事件监听及回调
        menuPanel.addButton("事件监听及回调", "注册关闭事件", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            if (cookie != null)
                return;
            cookie = app.advise(ApplicationEvents4.class, new ApplicationEvents4() {
                @Override
                public void DocumentBeforeClose(Document doc, Holder<Boolean> cancel) {
                    logger.info("调用注册的关闭事件");
                    //! 由于事件可能是在ui线程触发并阻塞等待，在此显示弹框可能会把ui卡住。
                    if (!Platform.isWindows()) {
                        JOptionPane.showMessageDialog(null,"关闭事件被拦截");
                        cancel.value = true;
                    }
                }
            });
            }
        });

        menuPanel.addButton("事件监听及回调", "取消注册关闭事件", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (cookie == null)
                    return;
                cookie.close();
                cookie = null;
            }
        });

        menuPanel.addButton("事件监听及回调", "添加一个按钮", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if(control != null){
                    JOptionPane.showMessageDialog(null,"按钮已存在");
                    return;
                }
                CommandBar bar = app.get_CommandBars().Add("test",1, Variant.getMissing(), Variant.getMissing());
                control = bar.get_Controls().Add(1,1,"test",1,"test");
                bar.put_Visible(true);
                control.put_Visible(true);
                control.put_Caption("test");
                JOptionPane.showMessageDialog(null,"按钮添加成功");
            }
        });

        menuPanel.addButton("事件监听及回调", "注册按钮事件", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if(control == null){
                    JOptionPane.showMessageDialog(null,"请先添加按钮");
                    return;
                }
                if(cookie2 != null){
                    JOptionPane.showMessageDialog(null,"事件已注册");
                    return;
                }
                cookie2 = control.advise(_CommandBarButtonEvents.class, new _CommandBarButtonEvents() {
                    @Override
                    public void Click(CommandBarButton ctrl, Holder<Boolean> cancelDefault) {
                        JOptionPane.showMessageDialog(null,"你点击了按钮 test");
                        cancelDefault.value = true;
                    }
                });
                JOptionPane.showMessageDialog(null,"注册事件成功");
            }
        });

        menuPanel.addButton("事件监听及回调", "取消注册按钮事件", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if(control == null){
                    JOptionPane.showMessageDialog(null,"请先添加按钮");
                    return;
                }
                cookie2.close();
                cookie2 = null;
                JOptionPane.showMessageDialog(null,"取消注册事件成功");
            }
        });

    }

    private void initOthersMenu(){

        menuPanel.addButton("其他", "设置打印图像对象", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Options().put_PrintDrawingObjects(true);
            }
        });

        menuPanel.addButton("其他", "设置打印隐藏文字", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Options().put_PrintHiddenText(true);
            }
        });

        menuPanel.addButton("其他", "获取文件大小", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                File file = new File(getPath("请选择一个文字文件", FileDialog.LOAD));
                JOptionPane.showMessageDialog(null,file.length());
            }
        });

        menuPanel.addButton("其他", "开启自动备份", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                // windows 版本需2020年11月之后版本才支持
                app.get_Application().queryInterface(_ApplicationEx.class).put_ForceBackupEnable(true);
            }
        });

        menuPanel.addButton("其他", "关闭自动备份", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Application().queryInterface(_ApplicationEx.class).put_ForceBackupEnable(false);
            }
        });

        menuPanel.addButton("其他", "设置标题行重复", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().get_Rows().put_HeadingFormat(-1);
            }
        });
        menuPanel.addButton("其他", "取消标题行重复", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().get_Rows().put_HeadingFormat(0);
            }
        });

        menuPanel.addButton("其他", "批量插入书签", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_Bookmarks().Add("book1",Variant.getMissing());
                app.get_Selection().TypeText("，");
                app.get_ActiveDocument().get_Bookmarks().Add("book2",Variant.getMissing());
                app.get_Selection().TypeText("，");
                app.get_ActiveDocument().get_Bookmarks().Add("book3",Variant.getMissing());
                app.get_Selection().TypeText("，");
                app.get_ActiveDocument().get_Bookmarks().Add("book4",Variant.getMissing());
                app.get_Selection().TypeText("，");
            }
        });

        menuPanel.addButton("其他", "批量删除书签", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_Bookmarks().Item("book1").Delete();
                app.get_ActiveDocument().get_Bookmarks().Item("book2").Delete();
                app.get_ActiveDocument().get_Bookmarks().Item("book3").Delete();
                app.get_ActiveDocument().get_Bookmarks().Item("book4").Delete();
            }
        });

        menuPanel.addButton("其他", "获取全部书签", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Bookmarks bookmarks = app.get_ActiveDocument().get_Bookmarks();
                String res = new String();
                res += bookmarks.Item("book1").get_Name()+"，";
                res += bookmarks.Item("book2").get_Name()+"，";
                res += bookmarks.Item("book3").get_Name()+"，";
                res += bookmarks.Item("book4").get_Name()+"，";
                JOptionPane.showMessageDialog(null,"<html><body><p style='width: 200px;'>"+res+"</p></body></html>");
            }
        });

        menuPanel.addButton("其他", "替换书签内容", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_Bookmarks().Item("book1").get_Range().put_Text("1111");
                app.get_ActiveDocument().get_Bookmarks().Item("book2").get_Range().put_Text("2222");
                app.get_ActiveDocument().get_Bookmarks().Item("book3").get_Range().put_Text("3333");
                app.get_ActiveDocument().get_Bookmarks().Item("book4").get_Range().put_Text("4444");
            }
        });

        menuPanel.addButton("其他", "从光标位置删除", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().TypeBackspace();
            }
        });

        menuPanel.addButton("其他", "插入表格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Range range = app.get_ActiveDocument().get_Content();
                app.get_ActiveDocument().get_Tables().Add(range,3,5,Variant.getMissing(),Variant.getMissing());
            }
        });

        menuPanel.addButton("其他", "适应文字1_选定区域的单元格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().get_Cells().Item(1).put_FitText(true);
            }
        });

        menuPanel.addButton("其他", "按行拆分", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Table table = app.get_ActiveDocument().get_Tables().Item(1);
                Row row = table.get_Rows().Item(1);
                row.get_Cells().Split(3,2,Variant.getMissing());
            }
        });

        menuPanel.addButton("其他", "按列拆分", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Table table = app.get_ActiveDocument().get_Tables().Item(1);
                Column column = table.get_Columns().Item(2);
                column.get_Cells().Split(3,3,Variant.getMissing());
            }
        });

        menuPanel.addButton("其他", "拆分单元格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Table table = app.get_ActiveDocument().get_Tables().Item(1);
                Row row = table.get_Rows().Item(2);
                row.get_Cells().Item(1).Split(2,3);
            }
        });

    }

    private void initRibbon(){

        menuPanel.addButton("功能区", "隐藏/显示-功能区", new ActionListener() {
            private boolean visible = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                // windows企业版要2020年11月以后的版本才支持IWindowEx接口
                com.wps.api.tree.wps.Window win = app.get_ActiveWindow();
                WindowEx winEx = win.queryInterface(WindowEx.class);
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
                            logger.info("退出WPS");
                            app.Quit(Variant.getMissing(), Variant.getMissing(), Variant.getMissing());
                        }
                    }
                });
            }
        });
    }
}
