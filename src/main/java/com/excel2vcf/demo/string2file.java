package com.excel2vcf.demo;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * author: HPH
 * data: 2017/7/10 22:07
 * 凑vcf的内容
 */
public class string2file {
    //将content写入文件
    public static void string2files(String content) {
        try {
            File file = new File("e:\\user.vcf");
            if (!file.exists()) {
                file.createNewFile();
            }
            FileWriter fw = new FileWriter(file, true);
            BufferedWriter bw = new BufferedWriter(fw);
            bw.write(content);
            bw.flush();
            bw.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    //生成content内容
    public static String getcontent(Map<String, Object> map) {
        StringBuffer content = new StringBuffer();
        List<String> namelist = (List<String>) map.get("namelist");
        List<StringBuffer> numlist = (List<StringBuffer>) map.get("numlist");
        String vcfbegin = "BEGIN:VCARD";
        String aenter = "\n";
        String version = "VERSION:2.1";
        String qbencodehead = "N;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:;";
        String qb2head = "FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:";
        String tel = "TEL;CELL:";
        String end = "END:VCARD";
        for (int i = 0; i < namelist.size(); i++) {
            content.append(vcfbegin);
            content.append(aenter);
            content.append(version);
            content.append(aenter);
            content.append(qbencodehead);
            content.append(Quoted_printable.qpEncodeing(namelist.get(i)));
            content.append(";;;");
            content.append(aenter);
            content.append(qb2head);
            content.append("");
            content.append(aenter);
            content.append(tel);
            content.append(numlist.get(i));
            content.append(aenter);
            content.append(end);
            content.append(aenter);
        }
        return content.toString();
    }

}
