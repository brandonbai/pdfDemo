package com.github.brandonbai.pdfDemo;

import com.github.brandonbai.pdfDemo.util.FreemarkerUtil;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * main
 * @author brandon
 * @since 2018-08-01
 */
public class Main {

    public static void main(String[] args) throws Exception{

        Map<String, Object> map = new HashMap<String, Object>(3);
        map.put("name", "小明");
        map.put("address", "北京市朝阳区");
        map.put("email", "xiaoming@abc.com");
        String xmlStr = FreemarkerUtil.generate("resume.ftl", map);

        ByteArrayInputStream in = new ByteArrayInputStream(xmlStr.getBytes());
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);

        String outputFilePath = "/Users/xiaoming/简历.pdf";
        Docx4J.toPDF(wordMLPackage, new FileOutputStream(new File(outputFilePath)));

    }
}
