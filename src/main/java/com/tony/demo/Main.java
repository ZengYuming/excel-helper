package com.tony.demo;

import com.tony.demo.excel_helper.ExcelWriter;
import com.tony.demo.excel_helper.consts.ExcelType;
import com.tony.demo.my_entity.read.MyEntity;
import com.tony.demo.my_entity.read.MyFieldType;
import com.tony.demo.excel_helper.ExcelReader;
import com.tony.demo.excel_helper.exception.ExcelContentInvalidException;
import com.tony.demo.excel_helper.exception.ExcelParseException;
import com.tony.demo.excel_helper.exception.ExcelRegexpValidFailedException;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by TonyZeng on 2017/9/4.
 */
public class Main {
    public static void main(String[] args) {
//        write("F:\\192.168.130.87\\trunk\\excel-helper\\src\\main\\writ.xlsx");
//
        read("F:\\192.168.130.87\\trunk\\excel-helper\\src\\main\\read.xlsx");
    }

    public static void read(String path) {
        try {
            ExcelReader eh = ExcelReader.readExcel(path);
            //注册自定义类型
            ExcelReader.registerNewType(MyFieldType.class);
            List<MyEntity> entitys = null;
            try {
                entitys = eh.toEntitys(MyEntity.class);
                for (MyEntity d : entitys) {
                    System.out.println(d.toString());
                }
            } catch (ExcelParseException e) {
                System.out.println(e.getMessage());
            } catch (ExcelContentInvalidException e) {
                System.out.println(e.getMessage());
            } catch (ExcelRegexpValidFailedException e) {
                System.out.println(e.getMessage());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void write(String path) {
        try {
            ExcelWriter excelWriter = new ExcelWriter();
            List<MyEntity> data = new ArrayList<>();
            MyEntity entity = new MyEntity();
            entity.setName("曾芋茗");
            entity.setSex("男");
            entity.setAge(56);
            entity.setSomato(new MyFieldType("消瘦"));
            entity.setTel(13544276658L);
            entity.setCreateDate(new Timestamp(1504747452000L));
            data.add(entity);
            OutputStream out = new FileOutputStream(path);
            excelWriter.write(ExcelType.xlsx, "sss", data, MyEntity.class, out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
