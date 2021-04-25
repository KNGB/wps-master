package com.wps.examples;

import javax.swing.*;

/**
 * Project      : Wps JavaAPI Demo
 * Author       : 金山办公
 * Date         : 2019.12.9
 * Description  : 此项目为WPS二次开发Java版本Demo程序，二次开发接口的使用方式，均可以参照此项目中的代码样例实现。
 * Note         :
 *          1. Java中不提供函数默认参数，因而调用接口时必须给定义的所有参数都传参，如果某个对应位置的参数需要使用默认值，请传递 Variant.getMissing()
 *          2. 获取对象或数值的的接口可以直接判断返回值是否为null进行容错校验，操作类型的接口，容错校验需捕获Exception异常，以此判断操作是否执行成功。
 *          3. 所有对象的属性在设置和获取时，在java中均需使用加上get_或put_前缀的对应方法，无法通过 对象名.属性直接调用。
 */

public class Main {
    public static void main(String []arg) throws InterruptedException, ClassNotFoundException, UnsupportedLookAndFeelException, InstantiationException, IllegalAccessException {
        java.lang.System.setProperty("sun.awt.xembedserver", "true");           //Linux下必须加这一句才能调用
        //放开下面这段代码可以更换按钮主题风格
//        for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
//            if ("com.sun.java.swing.plaf.gtk.GTKLookAndFeel".equals(info.getClassName())) {
//                javax.swing.UIManager.setLookAndFeel(info.getClassName());
//                break;
//            }
//        }

        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                JFrame mainFrame = new JFrame();
                mainFrame.setTitle("WPS JAVA接口调用演示");                       //设置显示窗口标题
                mainFrame.setSize(1524, 768);                           //设置窗口显示尺寸
                mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);       //置窗口是否可以关闭
                mainFrame.setResizable(false);//禁止缩放
                JTabbedPane tabbedPane = new JTabbedPane();
                tabbedPane.add(new WpsMainPanel(), "WPS文字");
                tabbedPane.add(new EtMainPanel(), "WPS表格");
                tabbedPane.add(new WppMainPanel(), "WPS演示");
                mainFrame.add(tabbedPane);
                mainFrame.setVisible(true);                                     //设置窗口是否可见
            }
        });
    }

}
