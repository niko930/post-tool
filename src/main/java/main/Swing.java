package main;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import jxl.Sheet;
import jxl.Workbook;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.message.BasicHeader;
import org.apache.http.protocol.HTTP;

import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Map;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class Swing extends JFrame {    //继承JFrame顶层框架
    private static final int WIDTH = 900;        //程序的宽
    private static final int HEIGHT = 600;       //程序的高
    private static final String URL = "localhost:8080/xxxupdate";     //请求路径
    private static Integer count = 0;                   //总记录数

    //定义组件
    //上部组件
    private JPanel jp1;             //定义面板
    private JSplitPane jsp;         //定义拆分窗格
    private JTextArea jta1;         //定义文本域
    private JScrollPane jspane1;    //定义滚动窗格
    private JTextArea jta2;
    private JScrollPane jspane2;
    //下部组件
    private JPanel jp2;
    private JButton jb1, jb2;       //定义按钮
    private JComboBox jcb1;         //定义下拉框
    private JTextArea text;  //自定义输入框

    public Swing() {//构造函数
        //创建组件
        //上部组件
        jp1 = new JPanel();                 //创建面板
        jta1 = new JTextArea();             //创建多行文本框
        jta1.setLineWrap(true);             //设置多行文本框自动换行
        jta1.setEditable(false);            //不可编辑
        jspane1 = new JScrollPane(jta1);    //创建滚动窗格
        jta2 = new JTextArea();
        text = new JTextArea("请输入Authorization", 1, 20);
        jta2.setEditable(false);
        jta2.setLineWrap(true);
        jspane2 = new JScrollPane(jta2);
        jsp = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, jspane1, jspane2); //创建拆分窗格
        jsp.setDividerLocation(450);        //设置拆分窗格分频器初始位置
        jsp.setDividerSize(1);              //设置分频器大小
        //下部组件
        jp2 = new JPanel();
        jb1 = new JButton("上传");      //创建按钮
        jb2 = new JButton("开始");
        String[] name = {"需求打码倍数是否包含余额", "是", "否"};
        jcb1 = new JComboBox(name);         //创建下拉框

        jb1.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent event) {
                eventOnImport(jta1);
            }
        });
        jb2.addMouseListener(new MouseAdapter() {
            public void mouseClicked(MouseEvent e) {
                doPost(jta1, jta2);
            }
        });

        //设置布局管理
        jp1.setLayout(new BorderLayout());    //设置面板布局
        jp2.setLayout(new FlowLayout(FlowLayout.CENTER));

        //添加组件
        jp1.add(jsp);

        jp2.add(text);
        jp2.add(jcb1);
        jp2.add(jb1);
        jp2.add(jb1);
        jp2.add(jb2);

        this.add(jp1, BorderLayout.CENTER);
        this.add(jp2, BorderLayout.SOUTH);

        //设置窗体实行
        this.setTitle("小工具");                              //设置界面标题
        this.setIconImage(new ImageIcon("image/panda.ico").getImage());//设置标题图片
        this.setSize(WIDTH, HEIGHT);                         //设置界面像素
        this.setLocation(300, 300);                     //设置界面初始位置
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); //设置虚拟机和界面一同关闭
        this.setVisible(true);                               //设置界面可视化
    }

    /**
     * @param jta1 左侧显示框
     * @param jta2 右侧显示框
     */
    private void doPost(JTextArea jta1, JTextArea jta2) {
        System.out.println(jta1.getText());
        String[] s = jta1.getText().split("\n");
        System.out.println("size=" + s.length);
        int success = 0;
        int error = 0;
        String isCapitalString = (String) jcb1.getSelectedItem();
        String isCapital;
        if (isCapitalString.equals("是")) {
            isCapital = "1";
        } else {
            isCapital = "0";
        }
        if (jcb1.getSelectedIndex() == 0) {
            jta2.append("参数无效，请选择参数！\n");
            return;
        }
        if ("请输入Authorization".equals(text.getText()) || "".equals(text.getText())) {
            jta2.append("请输入Authorization!!!\n");
            return;
        }
        for (int i = 1; i < s.length; i++) {
            String[] split = s[i].split("\t");
            String name = split[0];
            String money = split[1];
            String damaMultiple = split[2];
            String note = split[3];
//            Map<String, String> params = new HashMap<String, String>();
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("userId", name);                 //ID
            jsonObject.put("balance", money);               //彩金
            jsonObject.put("damaMultiple", damaMultiple);   //需求打码倍数
            jsonObject.put("note", note);                   //备注
            jsonObject.put("isCapital", isCapital);         //需求打码倍数是否包含余额

//            params.put("userId", name);             //ID
//            params.put("balance", money);           //彩金
//            params.put("damaMultiple", damaMultiple);  //需求打码倍数
//            params.put("note", note);               //备注
//            params.put("isCapital", isCapital);     //需求打码倍数是否包含余额

            System.out.println(jsonObject);
            //1.定义请求类型
            HttpClient client = new DefaultHttpClient();
            HttpPost post = new HttpPost(URL);
            post.setHeader("Content-Type", "application/json;charset=UTF-8");
            post.setHeader("Authorization", text.getText());
            post.setHeader("Host", "***");
            //2.字符集
            String charset = "UTF-8";
            //3.判断用户是否传递参数
            if (jsonObject != null) {
                StringEntity stringEntity = new StringEntity(jsonObject.toString(), "utf-8");
                stringEntity.setContentEncoding(new BasicHeader(HTTP.CONTENT_TYPE, "application/json"));
                post.setEntity(stringEntity);
            }
            String result = null;
            try {
                // 发送请求
                HttpResponse httpResponse = client.execute(post);
                // 获取响应输入流
                InputStream inStream = httpResponse.getEntity().getContent();
                BufferedReader reader = new BufferedReader(new InputStreamReader(inStream, "utf-8"));
                StringBuilder strber = new StringBuilder();
                String line = null;
                while ((line = reader.readLine()) != null)
                    strber.append(line + "\n");
                inStream.close();
                result = strber.toString();
            } catch (Exception e) {
                jta2.append("连接错误\n");
                e.printStackTrace();
            }
//            String str = httpClient.doPost(URL, params);
            JSONObject parse = (JSONObject) JSON.parse(result);
            Integer code = (Integer) parse.get("status");
            String msg = (String) parse.get("msg");
            if (code == 1 && "SUCCESS".equals(msg)) {
                jta2.append(name + "\t" + money + "\t---成功\n");
                success++;
            } else {
                jta2.append(code + "\t" + msg + "\t---失败\n");
                error++;
            }
        }
        jta2.append("共计:" + count + " ,成功:" + success + " ,失败:" + error + "\n");
    }

    /**
     * excel上传
     *
     * @param jta1 左侧显示框
     */
    private static void eventOnImport(JTextArea jta1) {
        JFileChooser fileChooser = new JFileChooser();
        File chooseFile = fileChooser.getSelectedFile();
        //过滤Excel文件，只寻找以xls结尾的Excel文件，如果想过滤word文档也可以写上doc
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Text Files", "xls");
        fileChooser.setFileFilter(filter);
        int returnValue = fileChooser.showOpenDialog(null);
        //弹出一个文件选择提示框
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            //当用户选择文件后获取文件路径
            chooseFile = fileChooser.getSelectedFile();

            //根据文件路径初始化Excel工作簿
            Workbook workBook = null;
            try {
                workBook = Workbook.getWorkbook(chooseFile);
            } catch (Exception e) {
                e.printStackTrace();
            }
            //获取该工作表中的第一个工作表
            Sheet sheet = workBook.getSheet(0);
            //获取该工作表的行数，以供下面循环使用
            int rowSize = sheet.getRows();
            count = rowSize - 1;
            jta1.append("ID\t金额\t打码倍数\t备注\n");
            for (int i = 1; i < rowSize; i++) {
                Map<String, String> map = new HashMap();
                //获取姓名字段数据，第A列第i行，注意Excel中的行和列都是从0开始获取的，A列为0列
                String name = sheet.getCell(0, i).getContents();
                //去掉该所有空格
                name = name.replaceAll(" ", "");
                //获取第1列第i行
                String money = sheet.getCell(1, i).getContents();
                money = money.replaceAll(" ", "");
                //添加显示到jta1中
                String beishu = sheet.getCell(2, i).getContents();
                beishu = beishu.replaceAll(" ", "");
                String note = sheet.getCell(3, i).getContents();
                note = note.replaceAll(" ", "");
                jta1.append(name + "\t" + money + "\t" + beishu + "\t" + note + "\n");
                map.put(name, money);
            }
        }
    }

    public static void main(String[] args) {
        new Swing();    //显示界面
    }
}
