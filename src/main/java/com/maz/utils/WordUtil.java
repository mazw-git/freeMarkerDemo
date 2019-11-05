package com.maz.utils;


import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import freemarker.template.Configuration;
import freemarker.template.Template;


public class WordUtil {

    /**
     * 生成word文件
     * @param dataMap word中需要展示的动态数据，用map集合来保存
     * @param templateName word模板名称，例如：test.ftl
     * @param filePath 文件生成的目标路径，例如：D:/wordFile/
     * @param fileName 生成的文件名称，例如：test.doc
     */
    @SuppressWarnings("unchecked")
    public static void createWord(Map dataMap,String templateName,String filePath,String fileName){
        try {
            //创建配置实例
            Configuration configuration = new Configuration();

            //设置编码
            configuration.setDefaultEncoding("UTF-8");

            //ftl模板文件
            configuration.setClassForTemplateLoading(WordUtil.class,"/");

            //获取模板
            Template template = configuration.getTemplate(templateName);

            //输出文件
            File outFile = new File(filePath+File.separator+fileName);

            //如果输出目标文件夹不存在，则创建
            if (!outFile.getParentFile().exists()){
                outFile.getParentFile().mkdirs();
            }

            //将模板和数据模型合并生成文件
            Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));


            //生成文件
            template.process(dataMap, out);

            //关闭流
            out.flush();
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public static void main(String[] args) {
        /** 用于组装word页面需要的数据 */
       Map<String, Object> dataMap = new HashMap<String, Object>();
       dataMap.put("username", "小马");
       dataMap.put("date", "2019-11-01");
       dataMap.put("content", "零售价为V网好人");
       List<Map<String, Object>> list1 = new ArrayList<Map<String, Object>>();//题目  
       for(int i=0;i<10;i++){
           Map<String, Object> map = new HashMap<String, Object>();  
           map.put("title", "title"+i );  
           map.put("content", "content"+i);  
           map.put("author", "author" + i);  
           list1.add(map);  
       }
       dataMap.put("listInfo",list1);
       String filePath = "D:/doc_f/";
       String fileOnlyName = "生成Word文档.doc";
       /** 生成word  数据包装，模板名，文件生成路径，生成的文件名*/
       WordUtil.createWord(dataMap, "freeMarker.ftl", filePath, fileOnlyName);
       
       //wordToPdf("","");
   }
    
    /**
     * word转pdf
     * @param wordPath word的路径
     * @param pdfPath pdf的路径
     */
    public static boolean wordToPdf(String wordPath, String pdfPath){
        boolean result = false;
        try {
            XWPFDocument document=new XWPFDocument(new FileInputStream(new File("D:/doc_f/freeMarker.docx")));
            File outFile=new File("D:/doc_f/test.pdf");
            outFile.getParentFile().mkdirs();
            OutputStream out=new FileOutputStream(outFile);
            PdfOptions options= PdfOptions.create();
            PdfConverter.getInstance().convert(document,out,options);
            result = true;
        }
        catch (  Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}