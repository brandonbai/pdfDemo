package com.github.brandonbai.pdfDemo;

import com.github.brandonbai.pdfDemo.util.Docx4JUtil;
import com.github.brandonbai.pdfDemo.util.FreemarkerUtil;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
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

        Map<String, Object> map = new HashMap<>(3);
        map.put("name", "小明");
        map.put("address", "北京市朝阳区");
        map.put("email", "xiaoming@abc.com");
        String ftlName = "resume.ftl";
        String outputFilePath = "/Users/jifeihu/Desktop/简历.pdf";
        FileOutputStream os = new FileOutputStream(outputFilePath);
        Docx4JUtil.process(ftlName, map, os);

    }
}
