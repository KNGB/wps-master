Notice: Licenses for 'dependencies\api-1.1.jar' can be found in the ‘dependencies\license_for_api-1.1.txt’ file.

# WPS JavaAPI 概述 
## 基本介绍
WPS JavaAPI 属于WPS的跨平台二次开发解决方案的一部分，使用 Java 作为开发语言，具备快速开发、跨平台等优良特性。使用Java原生swing和awt作为窗体容器，大大提升了对平台和操作系统的兼容性。在窗体中通过Java API直接与WPS进行交互，特性如下：

- 具有完整的功能。可通过多种不同的方法对文档、表单和演示文稿进行创作、格式设置和操控；通过鼠标、键盘执行的操作几乎都能通过 Java API 完成；结合Java程序语言的编程逻辑，可以轻松实现自动化文档编辑；

- 交互方式灵活。支持通过的接口调用的方式实现功能，还支持事件注册回调的方式实现自由交互。

- 标准化集成。API 文档完整，接口设计符合 Java 语法规范，上手难度低，缩短开发周期。

## 使用场景
- 在第三方开发的系统中想要实现一些文档编辑功能，然而自行开发一套文档编辑系统成本非常高也不现实，所以可以直接将WPS集成到第三方的系统中，通过调用WPS二次开发接口实现想要的文档编辑功能。

- 此场景仅适用于将WPS客户端（Windows/Linux）与Java客户端集成，在桌面计算机面向用户使用；**不建议将WPS客户端安装在服务端与业务系统集成，此类场景可集成WPS在线预览/在线编辑服务**

## Java整合wps二次开发流程介绍
### 一、通用说明
1. Java中不提供函数默认参数，因而调用接口时必须给定义的所有参数都传参，如果某个对应位置的参数需要使用默认值，请传递 `Variant.getMissing()`，例如：

    ```java
    app.get_ActiveDocument().Protect(WdProtectionType.wdAllowOnlyReading,Variant.getMissing(),Variant.getMissing(),Variant.getMissing(),Variant.getMissing());
    ```

2. 获取对象或数值的接口可以直接判断返回值是否为null进行容错校验，操作类型的接口，容错校验需捕获Exception异常，以此判断操作是否执行成功。
3. 所有对象的属性在设置和获取时，在java中均需使用加上get_或put_前缀的对应方法，无法通过“对象名.属性”直接调用，例如：

    ```java
    CommandBarControl controlOfd = commandBar.get_Controls().get_Item(26);
    controlOfd.put_Visible(false);
    // 不能使用 controlOfd.visible = false;
    ```

4. 遇到接口参数名为`lcid`的参数，默认传递2052即可。
   
### 二、Java嵌入WPS流程
1. 环境搭建
    - 安装jdk1.7,eclipse或idea
    - 新建空的Java项目
    - 导入依赖jar包，以idea为例：`File>>Project Structure>>Libraries>>Add(+)>>Java>>选择所有jar包>>apply>>ok`

2. 编写代码测试主类创建空白swing窗体

    ```java
    import javax.swing.*;
    public class TestMain {
        public static void main(String[] args) {
            java.lang.System.setProperty("sun.awt.xembedserver", "true");           //Linux下必须加这一句
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    JFrame mainFrame = new JFrame();
                    mainFrame.setTitle("WPS JAVA接口调用演示");                       //设置显示窗口标题
                    mainFrame.setSize(1524, 768);                                   //设置窗口显示尺寸
                    mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);       //置窗口是否可以关闭
                    mainFrame.setResizable(false);//禁止缩放
                    mainFrame.setVisible(true);                                     //设置窗口是否可见
                }
            });
        }
    }
    ```

3. 创建用于装载WPS的Panel：

    ```java
    //紧接着上一步main函数中的:mainFrame.setVisible(true)之后
    //创建一个JPanel用于装载画布
    JPanel wpsPanel = new JPanel();
    wpsPanel.setLayout(new BorderLayout());
    //初始化一个空白的画布用于装载WPS
    Canvas client = new Canvas() {
        public void paint(Graphics g) {
            super.paint(g);
        }
    };
    wpsPanel.add(client, BorderLayout.CENTER);  //添加画布
    mainFrame.add(wpsPanel);                    //添加JPane
    ```

4. 将WPS嵌入到窗体内

    ```java
    // 获取用于装载WPS的Canves的句柄
    WindowIDProvider pid = (WindowIDProvider) client.getPeer();
    XEmbedCanvasPeer peer = (XEmbedCanvasPeer) pid;
    peer.removeXEmbedDropTarget();

    //将画布句柄和长宽等信息作为参数初始化WPS窗体
    WpsArgs args = WpsArgs.ARGS_MAP.get(WpsArgs.App.WPS);
    args.setWinid(pid.getWindow());
    args.setHeight(client.getHeight());
    args.setWidth(client.getWidth());
    ClassFactory.createApplication();// 创建WPS Application对象
    ```

> 注意：
> 1. 获取用于装载WPS的Canves的句柄(pid和peer)的语句，必须在`mainFrame.setVisible(true)`之后调用，否则拿到空指针。
> 2. 获取用于装载WPS的Canves的句柄(pid和peer)的语句，不能写在JPanel的构造函数里，必须在JPanel对象创建完成后再获取，否则拿到空指针。

### 三、Java调用WPSAPI流程
1. 创建WPS App对象

    ```java
    // ......
    // ClassFactory.createApplication();
    Application app = ClassFactory.createApplication()// 创建一个WPS Application对象
    ```

2. 通过app调用二次开发接口操作文档。

    ```java
    //获取WPS app对象
    Application app = ClassFactory.createApplication();
    //打开文档
    app.get_Documents().Add(Variant.getMissing(), Variant.getMissing(), Variant.getMissing(), Variant.getMissing());
    //获取版本号
    String version = app.get_Build();
    //光标位置插入文本，此处写入版本号
    app.get_Selection().TypeText(version);
    //保存文档
    app.get_ActiveDocument().SaveAs(
            "~/桌面/test.wps",         //保存文件位置
            0,                        //文件类型 0-->wps|doc
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
    //关闭文档
    app.get_ActiveDocument().Close(false, Variant.getMissing(), Variant.getMissing());
    ```

### 四、注册事件回调流程
- 注册关闭事件

    ```java
    EventCookie cookie = app.advise(ApplicationEvents4.class, new ApplicationEvents4() {
        @Override
        public void DocumentBeforeClose(Document doc, Holder<Boolean> cancel) {
            JOptionPane.showMessageDialog(null,"关闭事件被拦截");
            cancel.value = true;        //cancel为true，代表取消关闭
        }
    });
    ```

- 取消注册关闭事件

    ```java
    cookie.close();
    cookie = null;
    ```

- 给按钮注册事件

    ```java
    // 添加一个按钮
    CommandBar bar = app.get_CommandBars().Add("test",1, Variant.getMissing(), Variant.getMissing());
    CommandBarControl control = bar.get_Controls().Add(1,1,"test",1,"test");
    bar.put_Visible(true);
    control.put_Visible(true);
    control.put_Caption("test");
    //给按钮注册事件
    EventCookie cookie = control.advise(_CommandBarButtonEvents.class, new _CommandBarButtonEvents() {
        @Override
        public void Click(CommandBarButton ctrl, Holder<Boolean> cancelDefault) {
            JOptionPane.showMessageDialog(null,"你点击了按钮 test");
            cancelDefault.value = true;
        }
    });
    ```

- 取消注册按钮事件

    ```java
    cookie.close();
    cookie = null;
    ```

## 完整API使用Demo
参见本项目`demo`。

### 目录结构

* demo 示例工程源代码
* dependencies 示例工程依赖的jar包, 其中以下两个是wps提供的开发包。
	- api-1.0.jar wps对象树的接口
	- runtime-1.0.jar 运行时包，负责接口调用和进程通信
  
### 使用eclipse导入
安装支持二次开发的2019版wps，安装jdk1.7和eclipse, 打开eclipse导入demo工程，运行main函数即可启动界面。

### Windows环境运行时需注意
- 文字组件（wps）对应的GUID是`000209FF-0000-0000-C000-000000000046`
- 表格组件（et）对应的GUID是`00024500-0000-0000-C000-000000000046`
- 演示组件（wpp）对应的GUID是`91493441-5A91-11CF-8700-00AA0060263B`
- **文字、表格、演示分别与微软的Word、Excel和PowerPoint所用的GUID相同，请确认注册表中相对应的程序值是WPS各组件**