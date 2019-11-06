package com.maz.utils;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class WordAndPdfUtil {
    private static Logger logger = LoggerFactory.getLogger(WordAndPdfUtil.class);


    public static void main(String[] args) throws Exception {

        String basePath = "C:\\Users\\suzan\\Desktop\\";
        Map<String, Object> dataMap = new HashMap<String, Object>();
        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>(16);
        for (int i = 0; i < 10; i++) {
            Map<String, Object> itemMap = new HashMap<String, Object>(16);
            itemMap.put("item", i+1);
            itemMap.put("name", "name"+(i+1));
            itemMap.put("age", "age"+ (i+1));
            itemMap.put("adress", "adress" + (i+1));
            list.add(itemMap);
        }
        dataMap.put("itemList", list);
        dataMap.put("title", "我是docx的标题");
        dataMap.put("text", "我是被替换后的docx的对象");

        makeWord(basePath, "wordToPdf.zip", "wordToPdf.docx", dataMap);
        wordToPdf(basePath + "wordToPdf.docx",
                basePath + "wordToPdf.pdf");
    }


    /**
     * 生成word
     * @param basePath word所在的文件夹
     * @param wordInZipName word改为zip后word的名字
     * @param wordOutName word输出的名称
     * @param dataMap word里面需要替换的参数
     */
    public static  boolean makeWord(String basePath, String wordInZipName, String wordOutName, Map<String,Object> dataMap) {
        logger.info("开始创建word");
        boolean result = false;
        /** 指定输出word文件的路径 **/
        String outFilePath = basePath + "data.xml";
        File docXmlFile = new File(outFilePath);

        try {
            /** 初始化配置文件 **/
            Configuration configuration = new Configuration();
            configuration.setDefaultEncoding("utf-8");
            /** 加载文件 **/
            configuration.setDirectoryForTemplateLoading(new File(basePath));
            /** 加载模板 **/
            Template template = configuration.getTemplate("document.xml");
            template.setEncoding("utf-8");
            logger.info("初始化配置文件，成功， 开始渲染数据");

            /**数据渲染到word**/
            FileOutputStream fos = new FileOutputStream(docXmlFile);
            OutputStreamWriter oWriter = new OutputStreamWriter(fos, "UTF-8");
            Writer out = new BufferedWriter(oWriter, 10240);
            logger.info("数据写入xml中");
            template.process(dataMap, out);
            out.close();
            fos.close();
            logger.info("数据写入xml完毕， 开始读取zip文件");

            /**读取压缩文件**/
            ZipInputStream zipInputStream = wrapZipInputStream(new FileInputStream(new File(basePath + wordInZipName)));
            /**压缩文件写入到目标路径**/
            File wordOutFile =  new File(basePath + wordOutName);
            logger.info("创建新的目录：{}", basePath + wordOutName);
            wordOutFile.createNewFile();
            ZipOutputStream zipOutputStream = wrapZipOutputStream(new FileOutputStream(wordOutFile));
            String itemName = "word/document.xml";
            /**替换参数**/
            logger.info("替换相关参数");
            replaceItem(zipInputStream, zipOutputStream, itemName, new FileInputStream(docXmlFile));
            logger.info("word生成成功");
            result = true;
        } catch (IOException e) {
            logger.error("IO异常：{}", e.getMessage());
            e.printStackTrace();
        } catch (TemplateException e) {
            logger.error("模板异常：{}", e.getMessage());
            e.printStackTrace();
        } finally {
            docXmlFile.delete();
        }
        return result;
    }

    /**
     * word转pdf
     * @param wordPath word的路径
     * @param pdfPath pdf的路径
     */
    public static boolean wordToPdf(String wordPath, String pdfPath){
        logger.info("wordPath:{}, pdfPath:{}", wordPath, pdfPath);
        boolean result = false;
        try {
            logger.info("开始word转pdf");
            XWPFDocument document=new XWPFDocument(new FileInputStream(new File(wordPath)));
            File outFile=new File(pdfPath);
            outFile.getParentFile().mkdirs();
            OutputStream out=new FileOutputStream(outFile);
            PdfOptions options= PdfOptions.create();
            PdfConverter.getInstance().convert(document,out,options);
            logger.info("word转pdf成功");
            result = true;
        }
        catch (  Exception e) {
            e.printStackTrace();
            logger.error("word转pdf失败");
        }
        return result;
    }




    /**
     * 替换某个 item,
     * @param zipInputStream zip文件的zip输入流
     * @param zipOutputStream 输出的zip输出流
     * @param itemName 要替换的 item 名称
     * @param itemInputStream 要替换的 item 的内容输入流
     */
    public static void replaceItem(
            ZipInputStream zipInputStream,
            ZipOutputStream zipOutputStream,
            String itemName,
            InputStream itemInputStream){
        if(null == zipInputStream){return;}
        if(null == zipOutputStream){return;}
        if(null == itemName){return;}
        if(null == itemInputStream){return;}
        ZipEntry entryIn;
        try {
            while((entryIn = zipInputStream.getNextEntry())!=null)
            {
                String entryName =  entryIn.getName();
                ZipEntry entryOut = new ZipEntry(entryName);
                // 只使用 name
                zipOutputStream.putNextEntry(entryOut);
                // 缓冲区
                byte [] buf = new byte[8*1024];
                int len;

                if(entryName.equals(itemName)){
                    // 使用替换流
                    while((len = (itemInputStream.read(buf))) > 0) {
                        zipOutputStream.write(buf, 0, len);
                    }
                } else {
                    // 输出普通Zip流
                    while((len = (zipInputStream.read(buf))) > 0) {
                        zipOutputStream.write(buf, 0, len);
                    }
                }
                // 关闭此 entry
                zipOutputStream.closeEntry();

            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            //e.printStackTrace();
            close(itemInputStream);
            close(zipInputStream);
            close(zipOutputStream);
        }
    }

    /**
     * 包装输入流
     */
    public static ZipInputStream wrapZipInputStream(InputStream inputStream){
        ZipInputStream zipInputStream = new ZipInputStream(inputStream);
        return zipInputStream;
    }

    /**
     * 包装输出流
     */
    public static ZipOutputStream wrapZipOutputStream(OutputStream outputStream){
        ZipOutputStream zipOutputStream = new ZipOutputStream(outputStream);
        return zipOutputStream;
    }
    private static void close(InputStream inputStream){
        if (null != inputStream){
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void close(OutputStream outputStream){
        if (null != outputStream){
            try {
                outputStream.flush();
                outputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
