package com.excel2vcf.demo;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * author: HPH
 * data: 2017/7/10 10:31
 * 读取excel里面的个人信息。返回一个map.
 */
public class excel2list {
    public static Map<String, Object> getexcelcontent(){
        File files = new File("e:\\workbook.xlsx");
        List<String> namelist = new ArrayList<>();//存名字列
        List<String> numlist = new ArrayList<>();//存手机号列
        Map<String, Object> retmap = new HashMap<>();//返回的map
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(files));//新建工作簿
            XSSFSheet sheet = null;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {//遍历sheet页
                sheet = workbook.getSheetAt(i);
                for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {//遍历每一行
                    XSSFRow row = sheet.getRow(j);
                    namelist.add(String.valueOf(row.getCell(1)));//获取名字列
                    numlist.add(String.valueOf(row.getCell(4)));//获取手机号列
                }
            }
            namelist.remove(1);
            namelist.remove(0);
            numlist.remove(1);//去掉空行
            List<StringBuffer> sb = new ArrayList<>(); //存转换后的手机号
            for (String num : numlist) {
                sb.add(new StringBuffer(num));
            }
            sb.remove(0);//去掉空行 下面的for给手机号字符串格式化
            for (StringBuffer snumlist : sb) {
                snumlist.insert(1, "-");
                snumlist.insert(5, "-");
                snumlist.insert(9, "-");
            }
            retmap.put("namelist", namelist);
            retmap.put("numlist", sb);
            retmap.put("oldnamelist", numlist);
        }catch (Exception e){
            e.printStackTrace();
        }
        return retmap;
    }

    //一个遍历excel中所有格子的方法
    public void getAllCells(String path) {
        File files = new File(path);
        List<Object> namelist = new ArrayList<>();
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(files));//新建工作簿
            XSSFSheet sheet = null;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {//遍历sheet页
                sheet = workbook.getSheetAt(i);
                for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {//遍历每一行
                    XSSFRow row = sheet.getRow(j);
                    for (int k = 0; k < row.getPhysicalNumberOfCells(); k++) {//遍历每一行的格子
                        namelist.add(row.getCell(k));
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testfun()  {
        getAllCells("e:\\workbook.xlsx");
    }
}
