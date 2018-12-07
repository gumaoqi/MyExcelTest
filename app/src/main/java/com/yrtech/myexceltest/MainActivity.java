package com.yrtech.myexceltest;

import android.os.Environment;
import android.os.Handler;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Random;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {
    final String TAG = "MainActivity";
    Button addButton;
    Button deleteButton;
    Button updateButton;
    Button queryButton;

    File file;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        addButton = findViewById(R.id.add_button);
        deleteButton = findViewById(R.id.delete_button);
        updateButton = findViewById(R.id.update_button);
        queryButton = findViewById(R.id.query_button);
        addButton.setOnClickListener(this);
        deleteButton.setOnClickListener(this);
        updateButton.setOnClickListener(this);
        queryButton.setOnClickListener(this);

        file = new File(Environment.getExternalStorageDirectory().getPath() + "/aaaExcelTest");//创建文件夹
        if (!file.exists()) {
            file.mkdir();
        }
        file = new File(Environment.getExternalStorageDirectory().getPath() + "/aaaExcelTest/user.xls");//指明存放数据的excel表示
        if (!file.exists()) {//如果文件不存在，创建文件
            createFile(file);
        }
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.add_button:
                User user = new User("用户" + System.currentTimeMillis(), "男", new Random().nextInt(20) + 20 + "");
                addUser(user, file);
                break;
            case R.id.delete_button:
                deleteUser(file);
                break;
            case R.id.update_button:
                updateUser(file);
                break;
            case R.id.query_button:
                queryUser(file);
                break;
        }
    }

    private void createFile(File file) {//创建一个file，并将第一行前三列分别设为：姓名，性别，年龄
        OutputStream os = null;
        WritableWorkbook wwb = null;
        try {
            file.createNewFile();
            os = new FileOutputStream(file);
            //创建一个可写的Workbook
            wwb = Workbook.createWorkbook(os);

            //创建一个可写的sheet,第一个参数是名字,第二个参数是第几个sheet,写入第一行
            WritableSheet sheet = wwb.createSheet("第一个sheet", 0);
            //创建一个Label,第一个参数是x轴,第二个参数是y轴,第三个参数是内容,第四个参数可选,指定类型
            Label label1 = new Label(0, 0, "姓名");
            Label label2 = new Label(1, 0, "性别");
            Label label3 = new Label(2, 0, "年龄");

            //把label加入sheet对象中
            sheet.addCell(label1);
            sheet.addCell(label2);
            sheet.addCell(label3);
            wwb.write();
            //只有执行close时才会写入到文件中,可能在close方法中执行了io操作
            wwb.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {//关闭流
            try {
                if (os != null) {
                    os.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private void addUser(User user, File file) {
        Workbook originWwb = null;
        WritableWorkbook newWwb = null;
        try {
            //插入数据需要拿到原来的表格originWwb，然后通过其创建一个新的表格newWwb，在newWwb上完成插入操作
            originWwb = Workbook.getWorkbook(file);
            newWwb = Workbook.createWorkbook(file, originWwb);
            //获取指定索引的表格
            WritableSheet ws = newWwb.getSheet(0);
            // 获取该表格现有的行数，将数据插入到底部
            int row = ws.getRows();
            Label lab1 = new Label(0, row, user.getName());//参数分别代表：列数，行数，插入的内容
            Label lab2 = new Label(1, row, user.getSex());
            Label lab3 = new Label(2, row, user.getAge());
            ws.addCell(lab1);
            ws.addCell(lab2);
            ws.addCell(lab3);
            // 从内存中的数据写入到sd卡excel文件中。
            newWwb.write();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {//释放资源
            if (originWwb != null) {
                originWwb.close();
            }
            if (newWwb != null) {
                try {
                    newWwb.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private void deleteUser(File file) {
        Workbook originWwb = null;
        WritableWorkbook newWwb = null;
        try {
            //插入数据需要拿到原来的表格originWwb，然后通过其创建一个新的表格newWwb，在newWwb上完成插入操作
            originWwb = Workbook.getWorkbook(file);
            newWwb = Workbook.createWorkbook(file, originWwb);
            //获取指定索引的表格
            WritableSheet ws = newWwb.getSheet(0);
            // 获取该表格现有的行数，将数据插入到底部
            int row = ws.getRows();
            ws.removeRow(row - 1);
            // 从内存中的数据写入到sd卡excel文件中。
            newWwb.write();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {//释放资源
            if (originWwb != null) {
                originWwb.close();
            }
            if (newWwb != null) {
                try {
                    newWwb.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private void updateUser(File file) {
        Workbook originWwb = null;
        WritableWorkbook newWwb = null;
        try {
            //插入数据需要拿到原来的表格originWwb，然后通过其创建一个新的表格newWwb，在newWwb上完成插入操作
            originWwb = Workbook.getWorkbook(file);
            newWwb = Workbook.createWorkbook(file, originWwb);
            //获取指定索引的表格
            WritableSheet ws = newWwb.getSheet(0);
            // 获取该表格现有的行数，将数据插入到底部
            int row = ws.getRows();
            ws.removeRow(row - 1);
            Label lab1 = new Label(0, row - 1, "张三");//参数分别代表：列数，行数，插入的内容
            Label lab2 = new Label(1, row - 1, "女");
            Label lab3 = new Label(2, row - 1, "60");
            ws.addCell(lab1);
            ws.addCell(lab2);
            ws.addCell(lab3);
            // 从内存中的数据写入到sd卡excel文件中。
            newWwb.write();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {//释放资源
            if (originWwb != null) {
                originWwb.close();
            }
            if (newWwb != null) {
                try {
                    newWwb.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private void queryUser(File file) {
        InputStream is = null;
        Workbook workbook = null;
        try {
            is = new FileInputStream(file.getPath());//获取流
            workbook = Workbook.getWorkbook(is);
            Sheet sheet = workbook.getSheet(0);
            for (int i = 0; i < sheet.getRows(); i++) {
                Log.i(TAG, sheet.getCell(0, i).getContents() + "-" + sheet.getCell(1, i).getContents() +
                        "-" + sheet.getCell(2, i).getContents());
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                workbook.close();
            }
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
