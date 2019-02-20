
package com.pagoda.nerp.uniseq.cgdd.common;

import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;
import com.pagoda.nerp.uniseq.api.erpgw.dto.FeePeriodDetail;
import com.pagoda.nerp.uniseq.api.erpgw.dto.StatementPeriod;
import com.pagoda.nerp.uniseq.api.erpgw.dto.StatementPeriodDetail;
import com.pagoda.nerp.uniseq.cgdd.dto.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;


public class PdfUtil {

    private static Logger logger = LoggerFactory.getLogger(PdfUtil.class);

    private static final int MAX_PAGE = 10;
    private static final int MAX_PAGE_5 = 5;
    private static final int MAX_PAGE_15 = 15;
    private static final String FOOTER = "第 %d 页，共 %d 页";

    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");
    private static final SimpleDateFormat TIME_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm");
    private static final SimpleDateFormat SIGN_FORMAT = new SimpleDateFormat("yyyy年MM月dd日");
    private static final DecimalFormat NUMBER_FORMAT = new DecimalFormat("###,###,###.####");


    private static final List<String> INVOICE_HEADERS = Arrays.asList("序号", "商品名称", "订单号", "单价", "数量", "单位", "金额");
    private static final List<String> SIGN_INVOICE_HEADERS = Arrays.asList("序号", "商品名称", "订单号", "单价", "数量", "单位", "金额", "实际入库品名", "实际入库单价", "实际入库数量", "实际入库金额", "收货人");
    private static final List<String> SIGN_CONSUMABLE_INVOICE_HEADERS = Arrays.asList("序号", "商品名称", "产品规格", "包装规格", "订单号", "单价", "数量", "单位", "金额", "实际入库数量", "实际入库金额", "收货人");

    private static final List<String> ORDER_HEADERS = Arrays.asList("序号", "商品名称", "单价", "数量", "单位", "金额", "到货机构", "到货日期", "验收标准");
    private static final List<String> ORDER_CONSUMABLE_HEADERS = Arrays.asList("序号", "商品名称", "产品规格", "包装规格", "单价", "数量", "单位", "金额", "到货机构", "到货日期", "验收标准");

    private static final List<String> ORDER_EXPORT_HEADERS = Arrays.asList("序号", "商品名称", "单价", "数量", "件单价", "件数", "件/规格", "单位", "金额", "要货机构", "要货时间", "产地");
    private static final List<String> ORDER_CONSUMABLE_EXPORT_HEADERS = Arrays.asList("序号", "商品名称", "产品规格", "单价", "数量", "件单价", "件数", "件/规格", "单位", "金额", "要货机构", "要货时间");

    private static final List<String> GOODS_HEADERS = Arrays.asList("序号", "商品名称", "核心标准", "主要标准");
    private static final List<String> CONSUMABLE_HEADERS = Arrays.asList("序号", "商品名称", "验收标准");

    private static final List<String> BILL_DETAIL_HEADERS = Arrays.asList("票据单号", "票据日期", "票据金额", "提前支取", "未付款", "商业折扣", "到货机构");
    private static final List<String> FEE_DETAIL_HEADERS = Arrays.asList("票据单号", "票据日期", "费用名称", "费用金额", "到货机构", "备注");
    /**
     * 月度账单状态（A-未出账单、B-账单已出待确认、C-已提出异议沟通中、D-账单有更新待确认、E-已确认并签章、F-已确认）
     */
    private static final HashMap<String, String> BILL_STATUS = new HashMap<String, String>(){{
        put("A","未出账单");
        put("B","账单已出待确认");
        put("C","已提出异议沟通中");
        put("D","账单有更新待确认");
        put("E","已确认并签章");
        put("F","已确认");
    }};

    private static final List<String> ORDER_REMARK = Arrays.asList("1、供应商必须在采购订单下达后4小时内完成订单确认。"
            ,"2、供应商必须按照订单规定的内容发货，若对订单有疑义，在确认订单之前沟通采购方采购员进行修改订单。"
            ,"3、供应商发货数量不能超过订单数量的10%。"
            ,"4、采购方货款结算周期为45天，供应商可以向采购方申请提前支取货款，提前支取货款的结算周期为以入采购方库房48小时后，按采购方综合评估供应方提前支付比例，支取相应批次的货款。"
            ,"5、若供应方违反了上述条款，无法完成订单内容，采购方有权对供应商追责。");

    private static final List<String> ORDER_REMARK_CONSUMABLE = Arrays.asList("1、供应商须在采购订单下达后4小时内完成订单确认，否则订单会自动失效退回\n（22:00-6:00时段不计入4小时时效，自动顺延）。"
            ,"2、供应商必须按照订单规定的内容发货，若对订单有疑义，在确认订单之前沟通采购方采购员。"
            ,"3、供应商发货数量不能超过订单数量。"
            ,"4、本订单是采购合同的一部分，订单违约视同合同违约。");

    public final static String ORDER_BUYER_KEY = "采购方签章";
    public final static String ORDER_VENDOR_KEY = "供应商签章";
    public final static String INVOICE_BUYER_KEY = "收货方签章";
    public final static String INVOICE_VENDOR_KEY = "发货方签章";
    public final static String STATEMENT_BUYER_KEY = "对账单制单人";
    public final static String STATEMENT_VENDOR_KEY = "供应商对账人";

    private static String join(String str1, String str2){
        return str1.concat("：").concat(str2 == null ? "" : str2);
    }

    /**
     * 创建表格
     * @param data  数据 List 或者 List<List>
     * @param font  字体
     * @param widths    设置每一列所占的长度 为空时均宽
     * @param spacing   每列间距
     * @param background
     * @param border
     * @param alignment 对齐方式
     * @return
     * @throws Exception
     */
    private static PdfPTable createTable(List data, Font font, float[] widths, int spacing, BaseColor background, int border, Object alignment)throws Exception {
        //创建一个N列的表格控件
        PdfPTable pdfTable = new PdfPTable(data.get(0) instanceof List ? ((List)data.get(0)).size() : data.size());
        pdfTable.setSpacingBefore(spacing);
        //设置表格占PDF文档100%宽度
        pdfTable.setWidthPercentage(100);
        //水平方向表格控件左对齐
        pdfTable.setHorizontalAlignment(PdfPTable.ALIGN_LEFT);
        //设置每一列所占的长度 默认为均宽
        if(widths != null) {
            pdfTable.setWidths(widths);
        }
        PdfPCell cell = new PdfPCell();
        cell.setBackgroundColor(background);
        int[] alignments = null;
        if(!alignment.getClass().isArray()) {
            cell.setHorizontalAlignment((int)alignment);
        }else{
            alignments = (int[]) alignment;
        }
        cell.setMinimumHeight(25.0f);
        cell.setUseAscender(true);
        cell.setUseDescender(true);
        cell.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
        cell.setBorder(border);

        PdfPCell cellMerge = new PdfPCell();
        cellMerge.setBackgroundColor(cell.getBackgroundColor());
        cellMerge.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
        cellMerge.setMinimumHeight(25.0f);
        cellMerge.setUseAscender(true);
        cellMerge.setUseDescender(true);
        cellMerge.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
        cellMerge.setBorder(border);
        cellMerge.setPaddingLeft(10);

        if(data.get(0) instanceof List) {
            for (int i = 0; i < data.size(); i++) {
                setTableCell(pdfTable, cell, cellMerge, (List) data.get(i), alignments, font);
            }
        }else{
            setTableCell(pdfTable, cell, cellMerge, data, alignments, font);
        }
        return pdfTable;
    }

    /**
     * 设置单元格
     * @param pdfTable
     * @param cell
     * @param cellMerge
     * @param data
     * @param alignments
     * @param font
     */
    private static void setTableCell(PdfPTable pdfTable, PdfPCell cell, PdfPCell cellMerge, List data, int[] alignments, Font font){
        for (int i = 0; i < data.size(); i++)  {
            String val = data.get(i) == null ? "" : data.get(i).toString();
            int col = 0;
            boolean isCenter = val.indexOf("colCenter") != -1;
            boolean isLeft = val.indexOf("colLeft") != -1;
            boolean isRight = val.indexOf("colRight") != -1;
            if(isCenter || isLeft || isRight){
                if(isCenter){
                    cellMerge.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
                    val = val.replace("colCenter", "");
                }else if(isLeft){
                    cellMerge.setHorizontalAlignment(PdfPCell.ALIGN_LEFT);
                    val = val.replace("colLeft", "");
                }else if(isLeft){
                    cellMerge.setHorizontalAlignment(PdfPCell.ALIGN_RIGHT);
                    val = val.replace("colRight", "");
                }
                col = 1;
                for(int j = i + 1; j < data.size(); j++, i++){
                    if(data.get(j) != null){
                        i = j - 1;
                        break;
                    }
                    col++;
                }
                cellMerge.setPhrase(new Paragraph(val, font));
                cellMerge.setColspan(col);
                pdfTable.addCell(cellMerge);
            }else {
                if(alignments != null){
                    cell.setHorizontalAlignment(alignments[i]);
                }
                cell.setPhrase(new Paragraph(val, font));
                pdfTable.addCell(cell);
            }
        }
    }


    private static List setMerger(int colNum, int[] indexs, Object[] values){
        return setMerger(colNum, indexs, values, 0);
    }

    /**
     * 设置合并列
     * @param colNum
     * @param indexs
     * @param locate    0-居中，1-靠左，2-靠右
     * @return
     */
    private static List setMerger(int colNum, int[] indexs, Object[] values, int locate){
        String flag = "colCenter";
        if(locate == 1){
            flag = "colLeft";
        }else if(locate == 2){
            flag = "colRight";
        }
        List row = new ArrayList();
        for(int i = 0; i < colNum; i++){
            int index = -1;
            for (int j = 0; j < indexs.length; j++){
                if(i == indexs[j]){
                    index = j;
                    break;
                }
            }
            row.add(index == -1 ?  null : flag + values[index]);
        }
        return row;
    }

    /**
     * 创建中文字体
     * @param fontSize
     * @return
     */
    private static Map<Integer, Font> createFont(List<Integer> fontSize){
        try {
            Map<Integer, Font> fontMap = new HashMap<>();
            BaseFont bfChinese = BaseFont.createFont("/stsong.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            fontSize.forEach(i -> fontMap.put(i, new Font(bfChinese, i, Font.NORMAL)));
            return fontMap;
        } catch (Exception e) {
            logger.error("创建中文字体失败 msg = {}", e);
        }
        return null;
    }

    private static void createPdf(String folder, String fileName, List elements) throws Exception{
        createPdf(folder, fileName, elements, null);
    }

    /**
     * 创建PDF
     * @param folder
     * @param fileName
     * @param elements
     * @throws Exception
     */
    private static void createPdf(String folder, String fileName, List elements, PageEventHelper pageEventHelper) throws Exception{
        Map<Integer, Font> font = createFont(Stream.of(10).collect(Collectors.toList()));
        //创建文件
        File file = new File(folder);
        if(!file.exists()){
            file.mkdirs();
        }
        final Document document =  new Document(PageSize.A4,20,20,20,20);
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(folder + fileName));
        if(pageEventHelper != null) {
            writer.setPageEvent(pageEventHelper);
        }
        //打开文件
        document.open();
//        document.newPage();
        for (Object ele: elements) {
            if(ele instanceof int[]){
                int[] pageCode = (int[]) ele;
                ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_RIGHT, new Phrase(String.format(FOOTER, pageCode[0], pageCode[1]), font.get(10)), document.right() - 10, document.bottom() - 10, 0);
                document.newPage();
            }

            if(ele instanceof List) {
                List<Element> pages = (List<Element>) ele;
                for (Element page : pages) {
                    document.add(page);
                }
            }
        }

        //设置属性
        //标题
        document.addTitle("电子单据");
        //作者
        document.addAuthor("pagoda");
        //主题
        document.addSubject("电子单据");
        //关键字
        document.addKeywords("优果联");
        //创建时间
        document.addCreationDate();
        //应用程序
        document.addCreator("pagoda.com.cn");

        //关闭文档
        document.close();
        //关闭书写器
        writer.close();
    }

    /**
     * 获取主题
     * @param val
     * @param font
     * @return
     */
    private static Paragraph getTitle(String val, Font font){
        Paragraph title = new Paragraph(val, font);
        title.setAlignment(Element.ALIGN_CENTER);
        return title;
    }

    /**
     * 获取单头
     * @param val
     * @param font
     * @return
     */
    private static Paragraph getBillHeader(String val, Font font){
        Paragraph singleHead = new Paragraph(val, font);
        singleHead.setAlignment(Element.ALIGN_RIGHT);
        singleHead.setIndentationRight(20);
        return singleHead;
    }

    /**
     * 获取信息头
     * @param data
     * @param font
     * @return
     * @throws Exception
     */
    public static PdfPTable getInfo(List data, Font font)throws Exception {
        return createTable(data, font, null, 15, BaseColor.WHITE, 0, PdfPCell.ALIGN_LEFT);
    }

    /**
     * 获取表格头
     * @param data
     * @param font
     * @param widths
     * @return
     * @throws Exception
     */
    public static PdfPTable getHeader(List data, Font font, float[] widths)throws Exception {
        return createTable(data, font, widths, 0, new BaseColor(0xF0, 0xF0, 0xF0), PdfPCell.BOX, PdfPCell.ALIGN_CENTER);
    }

    /**
     * 获取明细
     * @param data
     * @param font
     * @param widths
     * @return
     * @throws Exception
     */
    public static PdfPTable getBody(List data, Font font, float[] widths)throws Exception {
        return createTable(data, font, widths, 0, BaseColor.WHITE, PdfPCell.BOX, PdfPCell.ALIGN_CENTER);
    }

    /**
     * 获取签署人
     * @param type
     * @param first
     * @param second
     * @param font
     * @return
     * @throws Exception
     */
    private static PdfPTable getSigner(int type, String first, String second, Font font) throws Exception{
        String firstTitle = "";
        String secondTitle = "";
        switch(type){
            case 1 :
                firstTitle = INVOICE_VENDOR_KEY;
                secondTitle = INVOICE_BUYER_KEY;
                break;
            case 2:
                firstTitle = ORDER_BUYER_KEY;
                secondTitle = ORDER_VENDOR_KEY;
                break;
            case 3:
                firstTitle = STATEMENT_BUYER_KEY;
                secondTitle = STATEMENT_VENDOR_KEY;
        }
        return createTable(Arrays.asList(Arrays.asList(join(firstTitle, first), join(secondTitle, second))),
                font, null, 50, BaseColor.WHITE, PdfPCell.NO_BORDER, PdfPCell.ALIGN_LEFT);
    }

    /**
     * 获取签署日期
     * @param date
     * @param font
     * @return
     * @throws Exception
     */
    private static PdfPTable getSignDate(Date date, Font font) throws Exception{
        String signDate = join("签署日期", SIGN_FORMAT.format(date));
        return createTable(Arrays.asList(Arrays.asList(signDate, signDate)),
                font, null, 30, BaseColor.WHITE, PdfPCell.NO_BORDER, PdfPCell.ALIGN_LEFT);
    }

    /**
     * 数值格式化
     * @param num
     * @return
     */
    private static String num(Object num){
        return StringUtils.isEmpty(num) || num instanceof String ? "0"
                : NUMBER_FORMAT.format(num);
    }

    /**
     * 日期格式化
     * @param date
     * @return
     */
    private static String date(Date date){
        return date == null ? "" : DATE_FORMAT.format(date);
    }

    private static String time(Date date){
        return date == null ? "" : TIME_FORMAT.format(date);
    }


    public static void printStatement(StatementPeriod statementPeriod, String pdfPath) throws Exception {
        try {
            //字体设置
            Map<Integer, Font> font = createFont(Stream.of(10, 12, 14).collect(Collectors.toList()));
            //公共部分
            float[] widths = { 25, 20, 20, 20, 20, 20, 25 };
            float[] feeWidths = { 25, 20, 20, 20, 20, 25 };
            float[] infoWidths = { 25, 30, 20, 30 };

            //1、title 2、billHeader 3、info 4、header 5、body 6、remark 7、signer 8、signDate
            List element = new ArrayList();
            //加入logo "E://LLC/statement/pagoda_logo.png"
            Image image = Image.getInstance(PdfUtil.class.getClassLoader().getResource("pagoda_logo.png"));
            image.setAbsolutePosition(18, 810);
            image.scaleToFit(60f, 19f);

            Image title = Image.getInstance(PdfUtil.class.getClassLoader().getResource("pagoda_title.png"));
            title.setAbsolutePosition(180, 770);
            title.scaleToFit(193f,52f);

            Paragraph space =  getTitle("", font.get(14));
            space.setSpacingBefore(50);

            Paragraph billParagraph =  getTitle("对账单据明细：", font.get(12));
            billParagraph.setSpacingBefore(20);
            billParagraph.setSpacingAfter(10);
            billParagraph.setAlignment(Element.ALIGN_LEFT);


            Paragraph feeeParagraph =  getTitle("供应商扣项明细：", font.get(12));
            feeeParagraph.setSpacingBefore(20);
            feeeParagraph.setSpacingAfter(10);
            feeeParagraph.setAlignment(Element.ALIGN_LEFT);

            PdfPTable billHeader = getHeader(BILL_DETAIL_HEADERS, font.get(10), widths);
            PdfPTable feeHeader = getHeader(FEE_DETAIL_HEADERS, font.get(10), feeWidths);
            PdfPTable signDate = getSignDate(new Date(), font.get(12));
            PdfPTable signer = getSigner(3, statementPeriod.getUserName(), statementPeriod.getSignName(), font.get(12));

            BigDecimal defaultAmt = new BigDecimal(0.0);
            if(statementPeriod.getKpAmt() == null) statementPeriod.setKpAmt(defaultAmt);
            if(statementPeriod.getPkAmt() == null) statementPeriod.setPkAmt(defaultAmt);
            if(statementPeriod.getPayAmt() == null) statementPeriod.setPayAmt(defaultAmt);
            if(statementPeriod.getYuAmt() == null) statementPeriod.setYuAmt(defaultAmt);
            if(statementPeriod.getFpkAmt() == null) statementPeriod.setFpkAmt(defaultAmt);
            if(statementPeriod.getShipAmt() == null) statementPeriod.setShipAmt(defaultAmt);
            if(statementPeriod.getScfAmt() == null) statementPeriod.setScfAmt(defaultAmt);

            StringBuffer kpInfo = new StringBuffer();
            kpInfo.append(join("公司名称", statementPeriod.getCompanyName())).append("\n")
                    .append(join("纳税识别号", statementPeriod.getTaxpayerNo())).append("\n")
                    .append(join("地址电话", statementPeriod.getAddress())).append("\n")
                    .append(join("开户行及账号", statementPeriod.getBankInfo())).append("\n");
            List infoData = Arrays.asList(
                    Arrays.asList("对账单编号", statementPeriod.getBillNo(), "对账单状态", BILL_STATUS.get(statementPeriod.getStatus())),
                    Arrays.asList("供应商", statementPeriod.getProviderName(), "结算公司", statementPeriod.getCompanyName()),
                    Arrays.asList("开户行", statementPeriod.getProviderBankName(), "开户账号", statementPeriod.getProviderBankAccount()),
                    Arrays.asList("供货总额", num(statementPeriod.getShipAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN)), "提前支取", num(statementPeriod.getScfAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN))),
                    Arrays.asList("开票金额", num(statementPeriod.getKpAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN)), "预付款冲抵", num(statementPeriod.getYuAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN))),
                    Arrays.asList("票扣合计", num(statementPeriod.getPkAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN)), "非票扣合计", num(statementPeriod.getFpkAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN))),
                    Arrays.asList("本期实付金额(大写)", NumberToCN.number2CNMontrayUnit(statementPeriod.getPayAmt()), "本期实付金额(小写)", num(statementPeriod.getPayAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN))),
                    setMerger(4, new int[]{ 0, 1 }, new Object[]{ "开票信息", kpInfo.toString() }, 1));
            element.add(Arrays.asList(createTable(infoData, font.get(10), infoWidths, 0, BaseColor.WHITE, PdfPCell.BOX, PdfPCell.ALIGN_LEFT)));

            if(statementPeriod.getStatementPeriodDetails() != null && statementPeriod.getStatementPeriodDetails().size() > 0){
                List billDetailData =  new ArrayList<>();
                BigDecimal totalShipAmt =  new BigDecimal(0.0); // 票据金额
                BigDecimal totalScfAmt =  new BigDecimal(0.0); // 提前支取
                BigDecimal totalBalanceAmtAmt =  new BigDecimal(0.0); // 未付款
                BigDecimal totalDiscountAmt =  new BigDecimal(0.0); // 商业折扣
                int[] merge = new int[]{ 0, 2, 3, 4, 5, 6 };
                for (StatementPeriodDetail detail: statementPeriod.getStatementPeriodDetails()) {
                    if(detail.getShipAmt() == null) detail.setShipAmt(defaultAmt);
                    if(detail.getScfAmt() == null) detail.setScfAmt(defaultAmt);
                    if(detail.getBalanceAmt() == null) detail.setBalanceAmt(defaultAmt);
                    if(detail.getDiscountAmt() == null) detail.setDiscountAmt(defaultAmt);
                    totalShipAmt = totalShipAmt.add(detail.getShipAmt());
                    totalScfAmt = totalScfAmt.add(detail.getScfAmt());
                    totalBalanceAmtAmt = totalBalanceAmtAmt.add(detail.getBalanceAmt());
                    totalDiscountAmt = totalDiscountAmt.add(detail.getDiscountAmt());
                    billDetailData.add(Arrays.asList(detail.getShipNo(), detail.getShipDate(), num(detail.getShipAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN)), num(detail.getScfAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN)), num(detail.getBalanceAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN)), num(detail.getDiscountAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN)), detail.getBalanceDesc()));
                }
                billDetailData.add(setMerger(widths.length, merge, new String[]{ "合计", "￥"+ num(totalShipAmt.setScale(2, BigDecimal.ROUND_HALF_DOWN)), "￥"+ num(totalScfAmt.setScale(2, BigDecimal.ROUND_HALF_DOWN)), "￥"+ num(totalBalanceAmtAmt.setScale(2, BigDecimal.ROUND_HALF_DOWN)), "￥"+ num(totalDiscountAmt.setScale(2, BigDecimal.ROUND_HALF_DOWN)), ""  }));
                element.add(Arrays.asList(billParagraph, billHeader, getBody(billDetailData, font.get(10), widths)));
            }
            if(statementPeriod.getFeePeriodDetails() != null && statementPeriod.getFeePeriodDetails().size() > 0) {
                List feeDetailData = new ArrayList<>();
                BigDecimal totalFeeAmt =  new BigDecimal(0.0); // 费用金额
                int[] merge = new int[]{ 0, 3, 4 };
                for (FeePeriodDetail detail : statementPeriod.getFeePeriodDetails()) {
                    if(detail.getFeedAmt() == null) detail.setFeedAmt(defaultAmt);
                    totalFeeAmt = totalFeeAmt.add(detail.getFeedAmt());
                    feeDetailData.add(Arrays.asList(detail.getFeeBillNo(), detail.getFeeDate(), detail.getFeeName(), num(detail.getFeedAmt().setScale(2, BigDecimal.ROUND_HALF_DOWN)), detail.getOrgDesc(), detail.getRemark()));
                }
                feeDetailData.add(setMerger(feeWidths.length, merge, new String[]{ "合计", "￥"+ num(totalFeeAmt), "" }));
                element.add(Arrays.asList(feeeParagraph, feeHeader, getBody(feeDetailData, font.get(10), feeWidths)));
            }
            element.add(Arrays.asList(space, signer, signDate));
            PageEventHelper pageEventHelper = new PageEventHelper();
            pageEventHelper.setHeaders(Arrays.asList(image, title, space));
            createPdf(pdfPath, ConstDefine.EMAIL_ATTACHMENT_STATEMENT_STRANDER, element, pageEventHelper);
        } catch (Exception e) {
            logger.error("生成月度账单文件失败 msg = {}", e);
        }

    }


    public static void exportInvoice(List<InvoiceInfoPrintDto> data, String pdfPath) {
        try {
            //字体设置
            Map<Integer, Font> font = createFont(Stream.of(10, 12, 14).collect(Collectors.toList()));
            //公共部分
            int totalPage;
            int currentPage;
            float[] widths = { 10, 20, 15, 15, 15, 10, 15 };

            //合并行
            int[] mergeCount = {0, 6};
            int[] mergeRemark = {0};

            //1、title 2、billHeader 3、info 4、header 5、body 6、signer 7、signDate
            List element = new ArrayList();
            Paragraph title =  getTitle("发货单", font.get(14));
            PdfPTable header = getHeader(INVOICE_HEADERS, font.get(10), widths);

            for (InvoiceInfoPrintDto invoiceInfo : data) {
                List<InvoiceDetailsPrintDto> details = invoiceInfo.getDetails();
                //按配送中心分页
                totalPage = 0;
                currentPage = 1;
                String orgName = details.stream().findFirst().get().getOrgName();
                Map<String, Long> pageMap = details.stream().collect(Collectors.groupingBy(InvoiceDetailsPrintDto::getOrgName , Collectors.counting()));
                for (long i : pageMap.values()) {
                    totalPage = totalPage + ((int)i + MAX_PAGE_15 - 1 ) / MAX_PAGE_15;
                }

                String transportWay;
                String titleNo;
                if("A".equals(invoiceInfo.getTransportWay())){
                    transportWay = "汽运";
                    titleNo = "车牌号";
                }else  if("B".equals(invoiceInfo.getTransportWay())){
                    transportWay = "航运";
                    titleNo = "航班号";
                }else{
                    transportWay = "物流";
                    titleNo = "物流单号";
                }
                List<List> infoData =Arrays.asList(
                        Arrays.asList(join("发货方", invoiceInfo.getVendorName()), join("收货方", invoiceInfo.getCompanyName())),
                        Arrays.asList(join("运输方式", transportWay), join("到货机构", orgName)),
                        Arrays.asList(join(titleNo, invoiceInfo.getLicensePlate()), ""),
                        Arrays.asList(join("发货人", invoiceInfo.getDepartName()), ""),
                        Arrays.asList(join("发货日期", date(invoiceInfo.getDepartTime())), ""));
                Paragraph billHeader = getBillHeader(join("发货单号", invoiceInfo.getInvoiceNum()), font.get(10));
                PdfPTable infoTable = getInfo(infoData, font.get(10));

                BigDecimal totalAmt =  new BigDecimal(0.0); //总计
                List bodyData =  new ArrayList<>();
                int count = 0;
                int totalCount = 0;
                for (InvoiceDetailsPrintDto invoiceDetail : details) {
                    count++;
                    totalCount++;
                    BigDecimal totalPrice = invoiceDetail.getPurCount().multiply(invoiceDetail.getPurPrice());
                    totalAmt = totalAmt.add(totalPrice);
                    bodyData.add(Arrays.asList(count, invoiceDetail.getGoodsName(), invoiceDetail.getOrderNum(), num(invoiceDetail.getPurPrice()), num(invoiceDetail.getPurCount()), invoiceDetail.getPurUnitCode(), num(totalPrice)));
                    if(count >= MAX_PAGE_15 || !orgName.equals(invoiceDetail.getOrgName()) || totalCount == details.size()){
                        bodyData.add(setMerger(widths.length, mergeCount, new String[]{"合计", "￥"+ num(totalAmt) }));
                        bodyData.add(setMerger(widths.length, mergeRemark, new String[]{join("备注", invoiceInfo.getRemark())}, 1));
                        element.add(Arrays.asList(title, billHeader, infoTable, header, getBody(bodyData, font.get(10), widths)));
                        element.add(new int[]{currentPage, totalPage});
                        if(totalCount != details.size()) {
                            totalAmt = new BigDecimal(0.0); //总计
                            bodyData = new ArrayList<>();
                            count = 0;
                            currentPage++;
                            if (!orgName.equals(invoiceDetail.getOrgName())) {
                                orgName = invoiceDetail.getOrgName();
                                infoData.get(1).set(1, join("到货机构", orgName));
                            }
                        }
                    }
                }
            }
            createPdf(pdfPath, ConstDefine.EMAIL_ATTACHMENT_INVOICE_FILENAME, element);
        } catch (Exception e) {
            logger.error("生成发货单文件失败 msg = {}", e);
        }
    }

    public static void printInvoice(InvoiceInfoPrintDto invoiceInfo, String pdfPath) {
        try {
            //字体设置
            Map<Integer, Font> font = createFont(Stream.of(10, 12, 14).collect(Collectors.toList()));
            //公共部分
            int totalPage = 0;
            int currentPage = 1;
            float[] widths = { 10, 20, 15, 15, 15, 10, 15, 20, 15, 15, 15, 15 };

            //合并行
            int[] mergeCount = ConstDefine.CATEGORY_CONSUME_ORDER.equals(invoiceInfo.getCategory()) ? new int[]{0,  8, 9, 10, 11 } : new int[]{0, 6, 7, 10, 11 };
            int[] mergeRemark = {0};

            //1、title 2、billHeader 3、info 4、header 5、body 6、signer 7、signDate
            List element = new ArrayList();
            Paragraph title =  getTitle("收货单", font.get(14));
            PdfPTable header = getHeader(ConstDefine.CATEGORY_CONSUME_ORDER.equals(invoiceInfo.getCategory()) ? SIGN_CONSUMABLE_INVOICE_HEADERS : SIGN_INVOICE_HEADERS, font.get(10), widths);
            PdfPTable signDate = getSignDate(new Date(), font.get(12));
            PdfPTable signer = getSigner(1, invoiceInfo.getDepartName(), invoiceInfo.getSignReceiver(), font.get(12));
            Paragraph billHeader = getBillHeader(join("发货单号", invoiceInfo.getInvoiceNum()), font.get(10));
            List<InvoiceDetailsPrintDto> details = invoiceInfo.getDetails();
            //按配送中心分页
            String orgName = details.stream().findFirst().get().getOrgName();
            Map<String, Long> pageMap = details.stream().collect(Collectors.groupingBy(InvoiceDetailsPrintDto::getOrgName , Collectors.counting()));
            for (long i : pageMap.values()) {
                totalPage = totalPage + ((int)i + MAX_PAGE - 1 ) / MAX_PAGE;
            }
            String transportWay;
            String titleNo;
            if("A".equals(invoiceInfo.getTransportWay())){
                transportWay = "汽运";
                titleNo = "车牌号";
            }else  if("B".equals(invoiceInfo.getTransportWay())){
                transportWay = "航运";
                titleNo = "航班号";
            }else{
                transportWay = "物流";
                titleNo = "物流单号";
            }
            List<List> infoData =Arrays.asList(
                    Arrays.asList(join("发货方", invoiceInfo.getVendorName()), join("收货方", invoiceInfo.getCompanyName())),
                    Arrays.asList(join("运输方式", transportWay), join("到货机构", orgName)),
                    Arrays.asList(join(titleNo, invoiceInfo.getLicensePlate()), ""),
                    Arrays.asList(join("发货人", invoiceInfo.getDepartName()), ""),
                    Arrays.asList(join("发货日期", date(invoiceInfo.getDepartTime())), ""));
            PdfPTable infoTable = getInfo(infoData, font.get(10));

            BigDecimal totalAmt =  new BigDecimal(0.0); //总计
            BigDecimal storageTotalAmt =  new BigDecimal(0.0); //总计
            List bodyData =  new ArrayList<>();
            int count = 0;
            int totalCount = 0;
            for (InvoiceDetailsPrintDto invoiceDetail : details) {
                count++;
                totalCount++;
                BigDecimal totalPrice = invoiceDetail.getPurCount().multiply(invoiceDetail.getPurPrice());
                BigDecimal storageTotalPrice = invoiceDetail.getStorageCount() == null || invoiceDetail.getStoragePrice() == null ?  new BigDecimal(0)
                        : invoiceDetail.getStorageCount().multiply(invoiceDetail.getStoragePrice());
                totalAmt = totalAmt.add(totalPrice);
                storageTotalAmt = storageTotalAmt.add(storageTotalPrice);
                if(ConstDefine.CATEGORY_CONSUME_ORDER.equals(invoiceInfo.getCategory())) {
                    bodyData.add(Arrays.asList(count, invoiceDetail.getGoodsName(), invoiceDetail.getItemSpec(), invoiceDetail.getPackSpec(), invoiceDetail.getOrderNum(), num(invoiceDetail.getPurPrice()), num(invoiceDetail.getPurCount()), invoiceDetail.getPurUnitCode(), num(totalPrice),
                            num(invoiceDetail.getStorageCount()), num(storageTotalPrice), invoiceDetail.getReceiver()));
                }else{
                    bodyData.add(Arrays.asList(count, invoiceDetail.getGoodsName(), invoiceDetail.getOrderNum(), num(invoiceDetail.getPurPrice()), num(invoiceDetail.getPurCount()), invoiceDetail.getPurUnitCode(), num(totalPrice),
                            invoiceDetail.getStorageGoodsName(), num(invoiceDetail.getStoragePrice()), num(invoiceDetail.getStorageCount()), num(storageTotalPrice), invoiceDetail.getReceiver()));
                }
                if(count >= MAX_PAGE || !orgName.equals(invoiceDetail.getOrgName()) || totalCount == details.size()){
                    bodyData.add(setMerger(widths.length, mergeCount, new String[]{"合计", "￥"+ num(totalAmt), "", "￥"+ num(storageTotalAmt), "" }));
                    bodyData.add(setMerger(widths.length, mergeRemark, new String[]{join("备注", invoiceInfo.getRemark())}, 1));
                    element.add(Arrays.asList(title, billHeader, infoTable, header, getBody(bodyData, font.get(10), widths), signer, signDate));
                    element.add(new int[]{currentPage, totalPage});
                    if(totalCount != details.size()) {
                        totalAmt = new BigDecimal(0.0); //总计
                        storageTotalAmt = new BigDecimal(0.0);
                        bodyData = new ArrayList<>();
                        count = 0;
                        currentPage++;
                        if (!orgName.equals(invoiceDetail.getOrgName())) {
                            orgName = invoiceDetail.getOrgName();
                            infoData.get(1).set(1, join("到货机构", orgName));
                        }
                    }
                }
            }
            createPdf(pdfPath, ConstDefine.EMAIL_ATTACHMENT_INVOICE_FILENAME, element);
        } catch (Exception e) {
            logger.error("生成发货单文件失败 msg = {}", e);
        }
    }

    /**
     * 导出文件
     * @param data
     * @param pdfPath
     * @throws Exception
     */
    public static void exportOrder(List<OrderInfoPrintDto> data, String pdfPath) throws Exception {
        try {
            //字体设置
            Map<Integer, Font> font = createFont(Stream.of(10, 12, 14).collect(Collectors.toList()));
            //公共部分
            int totalPage;
            int currentPage;

            //1、title 2、billHeader 3、info 4、header 5、body 6、remark 7、signer 8、signDate
            List element = new ArrayList();
            Paragraph title =  getTitle("采购订单", font.get(14));

            PdfPTable header;
            PdfPTable orderHeaders = null;
            PdfPTable orderConsumableHeaders = null;

            float[] widths;
            //合并行
            int[] merge;

            for (OrderInfoPrintDto orderInfo : data) {
                currentPage = 1;
                totalPage = (orderInfo.getDetails().size() + MAX_PAGE_15 - 1 ) / MAX_PAGE_15;

                Paragraph billHeader = getBillHeader(join("订单号", orderInfo.getOrderNum()), font.get(10));
                if(ConstDefine.CATEGORY_CONSUME_ORDER.equals(orderInfo.getCategory())){
                    merge = new int[]{ 0, 9, 10 };
                    widths = new float[]{ 6, 14, 10, 10, 10, 8, 8, 10, 8, 10, 12, 10 };
                    if(orderConsumableHeaders == null) {
                        orderConsumableHeaders = getHeader(ORDER_CONSUMABLE_EXPORT_HEADERS, font.get(10), widths);
                    }
                    header = orderConsumableHeaders;
                }else{
                    merge = new int[]{ 0, 8, 9 };
                    widths = new float[]{ 6, 14, 8, 10, 8, 8, 10, 8, 10, 12, 10, 10 };
                    if(orderHeaders == null) {
                        orderHeaders = getHeader(ORDER_EXPORT_HEADERS, font.get(10), widths);
                    }
                    header = orderHeaders;
                }
                List infoData = Arrays.asList(
                        Arrays.asList(join("采购方", orderInfo.getCompanyName()), join("供应商", orderInfo.getVendorName())),
                        Arrays.asList(join("制单人", orderInfo.getBuyerName()), join("接单人", orderInfo.getConfirmer())),
                        Arrays.asList(join("制单日期",  date(orderInfo.getCreateTime())), ""),
                        Arrays.asList(join("审核人", orderInfo.getReviewer()), ""));
                PdfPTable infoTable = getInfo(infoData, font.get(10));

                BigDecimal totalAmt =  new BigDecimal(0.0); //总计
                List bodyData =  new ArrayList<>();
                int count = 0;
                int totalCount = 0;
                for (OrderDetailsPrintDto orderDetail : orderInfo.getDetails()) {
                    count++;
                    totalCount++;
                    BigDecimal totalPrice = orderDetail.getPurCount().multiply(orderDetail.getPurPrice());
                    totalAmt = totalAmt.add(totalPrice);
                    String purPrice = "--";
                    String purCount = "--";
                    String pcsPrice = "--";
                    String packages = "--";
                    String unitWeight = "--";
                    if("A".equals(orderDetail.getUnitType())){
                        purPrice = num(orderDetail.getPurPrice());
                        purCount = num(orderDetail.getPurCount());
                    }else{
                        pcsPrice = num(orderDetail.getPcsPrice());
                        packages = num(orderDetail.getPackages());
                        unitWeight = num(orderDetail.getUnitWeight());
                    }
                    if(ConstDefine.CATEGORY_CONSUME_ORDER.equals(orderInfo.getCategory())){
                        bodyData.add(Arrays.asList(count, orderDetail.getGoodsName(), orderDetail.getItemSpec(), purPrice, purCount, pcsPrice, packages, unitWeight,
                                orderDetail.getPurUnitCode(), num(totalPrice), orderDetail.getOrgName(), time(orderDetail.getArrivalTime())));
                    }else {
                        bodyData.add(Arrays.asList(count, orderDetail.getGoodsName(), purPrice, purCount, pcsPrice, packages, unitWeight,
                                orderDetail.getPurUnitCode(), num(totalPrice), orderDetail.getOrgName(), time(orderDetail.getArrivalTime()), orderDetail.getGoodsOrigin()));

                    }
                    if(count == MAX_PAGE_15 || totalCount == orderInfo.getDetails().size()){
                        bodyData.add(setMerger(widths.length, merge, new String[]{ "合计", "￥"+ num(totalAmt), "" }));
                        element.add(Arrays.asList(title, billHeader, infoTable, header, getBody(bodyData, font.get(10), widths)));
                        element.add(new int[]{currentPage, totalPage});
                        if(totalCount != orderInfo.getDetails().size()) {
                            totalAmt = new BigDecimal(0.0); //总计
                            bodyData = new ArrayList<>();
                            count = 0;
                            currentPage++;
                        }
                    }
                }
            }
            createPdf(pdfPath, ConstDefine.EMAIL_ATTACHMENT_ORDER_FILENAME, element);
        } catch (Exception e) {
            logger.error("生成订单文件失败 msg = {}", e);
        }
    }
    /**
     * 电子订单
     * @param orderInfo
     * @param pdfPath
     * @throws Exception
     */
    public static void printOrder(OrderInfoPrintDto orderInfo, String pdfPath) throws Exception {
        try {
            //字体设置
            Map<Integer, Font> font = createFont(Stream.of(10, 12, 14).collect(Collectors.toList()));
            //公共部分
            int totalPage = (orderInfo.getDetails().size() + MAX_PAGE - 1 ) / MAX_PAGE;
            int currentPage = 1;

            //1、title 2、billHeader 3、info 4、header 5、body 6、remark 7、signer 8、signDate
            List element = new ArrayList();
            Paragraph title =  getTitle("采购订单", font.get(14));
            Paragraph billHeader = getBillHeader(join("订单号", orderInfo.getOrderNum()), font.get(10));
            PdfPTable signDate = getSignDate(new Date(), font.get(12));
            PdfPTable signer = getSigner(2, orderInfo.getBuyerName(), orderInfo.getConfirmer(), font.get(12));

            int[] merge;
            float[] widths;
            PdfPTable header;
            Paragraph remark = new Paragraph();
            if(ConstDefine.CATEGORY_CONSUME_ORDER.equals(orderInfo.getCategory())){
                merge = new int[]{ 0, 7, 8 };
                widths = new float[]{ 6, 14, 8, 8, 10, 10, 8, 10, 12, 10, 10 };
                header = getHeader(ORDER_CONSUMABLE_HEADERS, font.get(10), widths);
                remark.addAll(ORDER_REMARK_CONSUMABLE.stream().map(r -> new Paragraph(r, font.get(10))).collect(Collectors.toList()));
            }else{
                merge = new int[]{ 0, 5, 6 };
                widths = new float[]{ 6, 14, 10, 10, 8, 10, 12, 10, 10 };
                header = getHeader(ORDER_HEADERS, font.get(10), widths);
                remark.addAll(ORDER_REMARK.stream().map(r -> new Paragraph(r, font.get(10))).collect(Collectors.toList()));
            }

            List infoData = Arrays.asList(
                    Arrays.asList(join("采购方", orderInfo.getCompanyName()), join("供应商", orderInfo.getVendorName())),
                    Arrays.asList(join("制单人", orderInfo.getBuyerName()), join("接单人", orderInfo.getConfirmer())),
                    Arrays.asList(join("制单日期",  date(orderInfo.getCreateTime())), ""),
                    Arrays.asList(join("审核人", orderInfo.getReviewer()), ""));
            PdfPTable infoTable = getInfo(infoData, font.get(10));

            BigDecimal totalAmt =  new BigDecimal(0.0); //总计
            List bodyData =  new ArrayList<>();
            int count = 0;
            int totalCount = 0;
            for (OrderDetailsPrintDto orderDetail : orderInfo.getDetails()) {
                count++;
                totalCount++;
                BigDecimal totalPrice = orderDetail.getPurCount().multiply(orderDetail.getPurPrice());
                totalAmt = totalAmt.add(totalPrice);

                if(ConstDefine.CATEGORY_CONSUME_ORDER.equals(orderInfo.getCategory())){
                    bodyData.add(Arrays.asList(count, orderDetail.getGoodsName(), orderDetail.getItemSpec(), orderDetail.getPackSpec(), num(orderDetail.getPurPrice()), num(orderDetail.getPurCount()),
                            orderDetail.getPurUnitCode(), num(totalPrice), orderDetail.getOrgName(), date(orderDetail.getArrivalTime()), "见附件"));
                }else {
                    bodyData.add(Arrays.asList(count, orderDetail.getGoodsName(), num(orderDetail.getPurPrice()), num(orderDetail.getPurCount()),
                            orderDetail.getPurUnitCode(), num(totalPrice), orderDetail.getOrgName(), date(orderDetail.getArrivalTime()), "见附件"));
                }
                if(count == MAX_PAGE || totalCount == orderInfo.getDetails().size()){
                    bodyData.add(setMerger(widths.length, merge, new String[]{ "合计", "￥"+ num(totalAmt), "" }));
                    element.add(Arrays.asList(title, billHeader, infoTable, header, getBody(bodyData, font.get(10), widths), remark, signer, signDate));
                    element.add(new int[]{currentPage, totalPage});
                    if(totalCount != orderInfo.getDetails().size()) {
                        totalAmt = new BigDecimal(0.0); //总计
                        bodyData = new ArrayList<>();
                        count = 0;
                        currentPage++;
                    }
                }
            }
            if(orderInfo.getGoodsStandards() != null && orderInfo.getGoodsStandards().size() >0){
                widths = new float[]{ 10, 20, 35, 35 };
                bodyData = new ArrayList();
                count = 0;
                totalCount = 0;
                currentPage = 1;
                totalPage = (orderInfo.getGoodsStandards().size() + MAX_PAGE_5 - 1 ) / MAX_PAGE_5;
                int[] alignments = { PdfPCell.ALIGN_CENTER, PdfPCell.ALIGN_CENTER, PdfPCell.ALIGN_LEFT, PdfPCell.ALIGN_LEFT };
                for (GoodsPartStrandDto standard : orderInfo.getGoodsStandards()){
                    count ++;
                    totalCount ++;
                    bodyData.add(Arrays.asList(count, standard.getGoods_name(),
                            standard.getItem_list().stream()
                                    .filter(item -> "核心标准".equals(item.getStandard_type_name()))
                                    .map(item -> item.getQc_item_name() + "：" + item.getStandard_desc()).collect(Collectors.joining("\n")),
                            standard.getItem_list().stream()
                                    .filter(item -> "主要标准".equals(item.getStandard_type_name()))
                                    .map(item -> item.getQc_item_name() + "：" + item.getStandard_desc()).collect(Collectors.joining("\n"))));
                    if(count == MAX_PAGE_5 || totalCount == orderInfo.getGoodsStandards().size()){
                        element.add(Arrays.asList(getTitle("附件", font.get(14)),
                                createTable(GOODS_HEADERS, font.get(10), widths, 30, new BaseColor(0xF0, 0xF0, 0xF0), PdfPCell.BOX, PdfPCell.ALIGN_CENTER),
                                createTable(bodyData, font.get(10), widths, 0, BaseColor.WHITE, PdfPCell.BOX, alignments)));
                        element.add(new int[]{currentPage, totalPage});
                        if(totalCount != orderInfo.getGoodsStandards().size()) {
                            bodyData = new ArrayList<>();
                            count = 0;
                            currentPage++;
                        }
                    }
                }

            }else if(orderInfo.getConsumableStandards() != null && orderInfo.getConsumableStandards().size() >0){
                widths = new float[]{ 10, 30, 60 };
                bodyData = new ArrayList();
                count = 0;
                totalCount = 0;
                currentPage = 1;
                totalPage = (orderInfo.getConsumableStandards().size() + MAX_PAGE_15 - 1 ) / MAX_PAGE_15;
                int[] alignments = { PdfPCell.ALIGN_CENTER, PdfPCell.ALIGN_CENTER, PdfPCell.ALIGN_LEFT };
                for (ConsumableStandardDto standard : orderInfo.getConsumableStandards()){
                    count ++;
                    totalCount ++;
                    bodyData.add(Arrays.asList(count, standard.getName(), standard.getStandard()));
                    if(count == MAX_PAGE_15 || totalCount == orderInfo.getConsumableStandards().size()){
                        element.add(Arrays.asList(getTitle("附件", font.get(14)),
                                createTable(CONSUMABLE_HEADERS, font.get(10), widths, 30, new BaseColor(0xF0, 0xF0, 0xF0), PdfPCell.BOX, PdfPCell.ALIGN_CENTER),
                                createTable(bodyData, font.get(10), widths, 0, BaseColor.WHITE, PdfPCell.BOX, alignments)));
                        element.add(new int[]{currentPage, totalPage});
                        if(totalCount != orderInfo.getConsumableStandards().size()) {
                            bodyData = new ArrayList<>();
                            count = 0;
                            currentPage++;
                        }
                    }
                }
            }
            createPdf(pdfPath, ConstDefine.EMAIL_ATTACHMENT_ORDER_FILENAME, element);
        } catch (Exception e) {
            logger.error("生成订单文件失败 msg = {}", e);
        }
    }

}