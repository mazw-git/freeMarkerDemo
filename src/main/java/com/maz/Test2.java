package com.maz;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.lowagie.text.pdf.BaseFont;

import java.awt.Color;

import com.lowagie.text.Font;

import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;

public class Test2 {
    
    public static void main(String[] args) {
        try {
            
            
            /*Map<String, Object> dataMap = new HashMap<String, Object>();
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
            
            
            
            Configuration configuration = new Configuration();
            configuration.setClassForTemplateLoading(Test2.class, "/");
            Template template = configuration.getTemplate("document.xml");
            String outFilePath = "D:\\data.xml";
            File docFile = new File(outFilePath);
            FileOutputStream fos = new FileOutputStream(docFile);
            Writer out = new BufferedWriter(new OutputStreamWriter(fos), 10240);
            template.process(dataMap, out);
            out.close();
            ZipInputStream zipInputStream = ZipUtils.wrapZipInputStream(new FileInputStream(new File(
                    "D:\\freemarkFile\\freeMarker.zip")));
            ZipOutputStream zipOutputStream = ZipUtils.wrapZipOutputStream(new FileOutputStream(new File(
                    "D:\\freemarkFile\\test.docx")));
            String itemname = "word/document.xml";
            ZipUtils.replaceItem(zipInputStream, zipOutputStream, itemname, new FileInputStream(new File(
                    "D:\\data.xml")));*/
            
            wordToPdf("","");
            
        } catch (Exception e) {
            // TODO: handle exception
            e.printStackTrace();
        }
    }

    public static boolean wordToPdf(String wordPath, String pdfPath){
        boolean result = false;
        try {
            XWPFDocument document=new XWPFDocument(new FileInputStream(new File("D:/freemarkFile/test.docx")));
            File outFile=new File("D:/freemarkFile/mytest.pdf");
            outFile.getParentFile().mkdirs();
            OutputStream out=new FileOutputStream(outFile);
            PdfOptions options= PdfOptions.create();
            
          //中文字体处理
            options.fontProvider(new IFontProvider() {
                public Font getFont(String familyName, String encoding, float size, int style, Color color) {
                    try {
                        BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
                        Font fontChinese = new Font(bfChinese, size, style, color);
                        if (familyName != null)
                        fontChinese.setFamily(familyName);
                        return fontChinese;
                    } catch (Exception e) {
                        e.printStackTrace();
                        return null;
                    }
                }
            });

            
            PdfConverter.getInstance().convert(document,out,options);
            result = true;
        }
        catch (  Exception e) {
            e.printStackTrace();
        }
        return result;
    }
    
}
