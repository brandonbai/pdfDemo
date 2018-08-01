package com.github.brandonbai.pdfDemo.util;

import freemarker.template.Configuration;
import freemarker.template.Template;

import java.io.BufferedWriter;
import java.io.StringWriter;

/**
 * freemarker 工具类
 * @author brandon
 * @since 2017-08-01
 */
public class FreemarkerUtil {

    private static Configuration config = null;

    /**
     * Static initialization.
     *
     * Initialize the configuration of Freemarker.
     */
    static{
        config = new Configuration();
        config.setClassForTemplateLoading(FreemarkerUtil.class, "/ftl/");
        config.setTemplateUpdateDelay(0);
    }

    public static Configuration getConfiguation(){
        return config;
    }

    /**
     * @param template
     * @param variables
     * @return
     * @throws Exception
     */
    public static String generate(String template, Object obj) throws Exception{
        Configuration config = getConfiguation();
        config.setDefaultEncoding("UTF-8");
        Template tp = config.getTemplate(template);
        StringWriter stringWriter = new StringWriter();
        BufferedWriter writer = new BufferedWriter(stringWriter);
        tp.setEncoding("UTF-8");
        tp.process(obj, writer);
        String htmlStr = stringWriter.toString();
        writer.flush();
        writer.close();
        return htmlStr;
    }


}
