package com.bgy.iText.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.List;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.ListItem;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPRow;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class TestPDFDemo1 {

    public static void main(String[] args) {
        try {
            test1();
        } catch (DocumentException | IOException ex) {
            ex.printStackTrace();
        }
    }
    
    public static void test1() throws FileNotFoundException, DocumentException {
        // 1.新建document对象
        Document doc = new Document();
        
        // 2.建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中。
        PdfWriter writer = PdfWriter.getInstance(doc, new FileOutputStream("E:/itext/test1.pdf"));
        
        //用户密码(打开文件需要输入此密码)
        String userPassword = "123456";
        //拥有者密码
        String ownerPassword = "hd";
//        writer.setEncryption(userPassword.getBytes(), ownerPassword.getBytes(), 
//                PdfWriter.ALLOW_PRINTING, PdfWriter.ENCRYPTION_AES_128);
        
        writer.setEncryption("".getBytes(), "".getBytes(), PdfWriter.ALLOW_PRINTING, PdfWriter.ENCRYPTION_AES_128);
        // 3.打开文档
        doc.open();

        // 4.添加一个内容段落
        doc.add(new Paragraph("some content here !"));
        
        //添加文档属性
        doc.addTitle("this is a title");
        doc.addAuthor("hsx");
        doc.addSubject("this is a subject");
        doc.addKeywords("keywords");
        doc.addCreationDate();
        doc.addCreator("hp.com");
        
        // 5.关闭文档
        doc.close();
        //关闭书写器
        writer.close();
    }
    
    /**
     * 添加图片到文档中
     * @throws DocumentException
     * @throws IOException
     */
    public static void addImgTest() throws DocumentException, IOException {
        Document  doc = new Document();
        
        PdfWriter writer = PdfWriter.getInstance(doc, new FileOutputStream("E:/itext/test2.pdf"));
        
        doc.open();
        
        doc.add(new Paragraph("HD content here"));
        
        //图片
        Image img1 = Image.getInstance("F:/aliyunServerReceive/IMG_2208.JPG");
        //设置图片位置的x轴和y轴
        img1.setAbsolutePosition(100f, 550f);
        //设置图片的宽度和高度
        img1.scaleAbsolute(200, 200);
        
        doc.add(img1);
        
        Image img2 = Image.getInstance(new URL("https://static.cnblogs.com/images/adminlogo.gif"));
        
        doc.add(img2);
        
        doc.close();
        writer.close();
    }

    /**
     * PDF中创建表格
     * @throws DocumentException
     * @throws FileNotFoundException
     */
    public static void tableTest() throws DocumentException, FileNotFoundException {
        Document  doc = new Document();
        
        PdfWriter writer = PdfWriter.getInstance(doc, new FileOutputStream("E:/itext/test3.pdf"));
        
        doc.open();
        
        doc.add(new Paragraph("HD content here"));
        
        //3列的表
        PdfPTable table = new PdfPTable(3);
        table.setWidthPercentage(100);          //宽度100%填充
        table.setSpacingBefore(10f);                //前间距
        table.setSpacingAfter(10f);                  //后间距
        
        List<PdfPRow> listRow = table.getRows();
        //设置列宽
        float[] columnWidths = {1f, 2f, 3f};
        table.setWidths(columnWidths);
        
        //行1
        PdfPCell[] cells1 = new PdfPCell[3];
        PdfPRow row1 = new PdfPRow(cells1);
        
        //单元格
        cells1[0] = new PdfPCell(new Paragraph("111"));
        cells1[0].setBorderColor(BaseColor.BLUE);
        cells1[0].setPaddingLeft(20f);
        cells1[0].setHorizontalAlignment(Element.ALIGN_CENTER);
        cells1[0].setVerticalAlignment(Element.ALIGN_MIDDLE);
        
        cells1[1] = new PdfPCell(new Paragraph("222"));
        cells1[2] = new PdfPCell(new Paragraph("333"));
        
        //行2
        PdfPCell[] cells2 = new PdfPCell[3];
        PdfPRow row2 = new PdfPRow(cells2);
        cells2[0] = new PdfPCell(new Paragraph("444"));
        
        listRow.add(row1);
        listRow.add(row2);
        //把表格添加到文件中
        doc.add(table);
        
        doc.close();
        writer.close();        
    }
    
    /**
     * PDF中创建列表
     * @throws DocumentException
     * @throws FileNotFoundException
     */
    public static void listTest() throws DocumentException, FileNotFoundException {
        Document  doc = new Document();
        
        PdfWriter writer = PdfWriter.getInstance(doc, new FileOutputStream("E:/itext/test4.pdf"));
        
        doc.open();
        
        doc.add(new Paragraph("HD content here"));
        
        //添加有序列表
        com.itextpdf.text.List orderedList = new com.itextpdf.text.List(com.itextpdf.text.List.ORDERED);
        orderedList.add(new ListItem("Item one"));
        orderedList.add(new ListItem("Item two"));
        orderedList.add(new ListItem("Item three"));
        
        doc.add(orderedList);
        
        doc.close();
        writer.close();
    }
}
