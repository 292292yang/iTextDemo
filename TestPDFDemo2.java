package com.bgy.iText.test;

import java.io.FileOutputStream;
import java.io.IOException;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chapter;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfWriter;

public class TestPDFDemo2 {

    public static void main(String[] args) {
        try {
            modifyPdf();
        } catch (DocumentException | IOException ex) {
            ex.printStackTrace();
        }
    }
    
    /**
     * 添加样式，生成多页文档
     * @throws DocumentException
     * @throws IOException
     */
    public static void cssTest() throws DocumentException, IOException {
        Document  doc = new Document();
        
        PdfWriter writer = PdfWriter.getInstance(doc, new FileOutputStream("E:/itext/test5.pdf"));
        
        doc.open();
        
        BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
        
        //蓝色字体
        Font blueFont = new Font(bfChinese);
        blueFont.setColor(BaseColor.BLUE);
        //段落文本
        Paragraph paragraphBlue = new Paragraph("paragraphOne blue front", blueFont);
        
        doc.add(paragraphBlue);
        
        //绿色字体
        Font greenFont = new Font(bfChinese);
        greenFont.setColor(BaseColor.GREEN);
        //创建章节(pdf中，一页就是一个章节, 章节页码从0开始)
        Paragraph chapterTitle = new Paragraph("段落标题xxxx", greenFont);
        Chapter chapter1 = new Chapter(chapterTitle, 1);
        chapter1.setNumberDepth(0);
        
        Paragraph sectionTitle = new Paragraph("部分标题", greenFont);
        chapter1.addSection(sectionTitle);
        
        Paragraph sectionContent = new Paragraph("部分内容", blueFont);
        chapter1.add(sectionContent);
        
        //将章节添加到文章中
        doc.add(chapter1);
        
        Paragraph chapter2Title = new Paragraph("段落标题2xxxx", blueFont);
        Chapter chapter2 = new Chapter(chapter2Title, 2);
        doc.add(chapter2);
        
        doc.close();
        writer.close();
    }
    
    
    public static void modifyPdf() throws DocumentException, IOException {
        PdfReader pdfReader = new PdfReader("E:/itext/test5.pdf");
        PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileOutputStream("E:/itext/test6.pdf"));
        
        Image img = Image.getInstance("F:/aliyunServerReceive/IMG_2207.JPG");
        img.scaleAbsolute(50, 50);
        img.setAbsolutePosition(0, 700);
        
        for (int i = 0; i < pdfReader.getNumberOfPages(); i++) {
            PdfContentByte content = pdfStamper.getUnderContent(i);
            content.addImage(img);
        }
        
        pdfStamper.close();
        pdfReader.close();
    }

}
