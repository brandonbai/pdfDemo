package com.github.brandonbai.pdfDemo.util;

import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

import javax.xml.bind.JAXBException;
import java.io.ByteArrayInputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * docx4j 工具类
 * @author brandon
 * @since 2017-08-01
 */
public class Docx4JUtil {
    /**
     * 普通文本段落
     */
    public final static String SAMPLE_TEXT = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            +"<w:pPr>\n" +
            "<w:jc w:val=\"${jc}\"/>\n" +
            "</w:pPr>"+
            "<w:r>"
            +"<w:rPr>"
            +"<w:rFonts w:ascii=\"Times New Roman\" w:cs=\"Times New Roman\" w:eastAsia=\"宋体\" w:hAnsi=\"Times New Roman\"/>"
            +"</w:rPr>"
            +"<w:t xml:space=\"preserve\">${content}</w:t>"
            +"</w:r>"
            +"</w:p>";

    /**
     * 粗体文本段落
     */
    public final static String SAMPLE_TEXT_BOLD = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            +"<w:r>"
            +"<w:rPr>"
            +"<w:rFonts w:ascii=\"Times New Roman\" w:cs=\"Times New Roman\" w:eastAsia=\"宋体\" w:hAnsi=\"Times New Roman\"/>"
            +"<w:b/>"
            +"</w:rPr>"
            +"<w:t xml:space=\"preserve\">${content}</w:t>"
            +"</w:r>"
            +"</w:p>";

    /**
     * 对象工厂，用于生成各种word文档实例
     */
    public static ObjectFactory objectFactory = new ObjectFactory();

    /**
     * 通过ftl和数据生成word文档
     * freemarker+docx4j
     * @param ftlName ftl文件名
     * @param obj 数据
     * @return
     * @throws Exception
     */
    public static WordprocessingMLPackage genaratePdfByFtlAndDocx4J(String ftlName, Object obj) throws Exception {
        String generate = FreemarkerUtil.generate(ftlName, obj);
        ByteArrayInputStream in = new ByteArrayInputStream(generate.getBytes());
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);
        return wordMLPackage;
    }

    /**
     * 通过ftl和数据生成word文档
     * freemarker+dom4j+docx4j
     * @param ftlName ftl文件名
     * @param obj 数据
     * @return
     * @throws Exception
     */
    public static WordprocessingMLPackage genaratePdfByFtlAndDom4jAndDoxc4J(String ftlName, Object obj, String header, String footer) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

        int left = (int)(567 * 1.8);
        int right = (int)(567 * 1.6);
        setDocMarginSpace(wordMLPackage, null, left, null, right);

        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

        if(header != null) {
            // 页眉
            Relationship styleRel = mainDocumentPart.getStyleDefinitionsPart().getSourceRelationships().get(0);
            RelationshipsPart relationshipsPart = mainDocumentPart.getRelationshipsPart();
            relationshipsPart.removeRelationship(styleRel);

            Relationship hRelationship = createHeaderPart(wordMLPackage, getHP(header));
            createHeaderReference(wordMLPackage, hRelationship);
        }

        if(footer != null) {
            // 页脚=页码+右下角文字
            Map<String, String> map = new HashMap<>();
            map.put("right", footer);
            String text = FreemarkerUtil.generate("bp_report_footer.ftl", map);
            Relationship fRelationship = createFooterPart(wordMLPackage, text);
            createFooterReference(wordMLPackage, fRelationship);
        }

        String generate = FreemarkerUtil.generate(ftlName, obj);
        Document document = DocumentHelper.parseText(generate);
        Element rootElement = document.getRootElement();
        List<Element> elements = rootElement.elements();
        for (Element element : elements) {
            String xmlString = element.asXML();
            Object o = XmlUtils.unmarshalString(xmlString);
            mainDocumentPart.addObject(o);
        }
        return wordMLPackage;
    }

    /**
     * 向文档中添加一段内容
     * @param mainDocumentPart 文档主体
     * @param template 模版
     * @param content 内容
     * @throws JAXBException
     */
    public static void addObject(MainDocumentPart mainDocumentPart, String template, String content) throws JAXBException {
        addObject(mainDocumentPart, template, content, JcEnumeration.LEFT);

    }

    /**
     * 向文档中添加一段内容
     * @param mainDocumentPart 文档主体
     * @param template 模版
     * @param content 内容
     * @param je 文本布局
     * @throws JAXBException
     */
    public static void addObject(MainDocumentPart mainDocumentPart, String template, String content, JcEnumeration je) throws JAXBException {
        if(content == null) {
            return;
        }
        final String brFlag = "<br/>";
        if(content.contains(brFlag)) {
            Br br = objectFactory.createBr();
            P p = objectFactory.createP();
            PPr ppr = objectFactory.createPPr();
            Jc jc = new Jc();
            jc.setVal(je);
            ppr.setJc(jc);

            RPr rpr = objectFactory.createRPr();
            RFonts rFonts = objectFactory.createRFonts();
            rFonts.setAscii("Times New Roman");
            rFonts.setEastAsia("宋体");
            rpr.setRFonts(rFonts);

            content = content.replace(brFlag, "|");
            String[] rArr = content.split("\\|");
            for (String rv : rArr) {
                R r = objectFactory.createR();
                Text text = objectFactory.createText();
                text.setValue(rv);
                text.setSpace("preserve");
                r.setRPr(rpr);
                r.getContent().add(br);
                r.getContent().add(text);
                p.getContent().add(r);
            }
            p.setPPr(ppr);

            mainDocumentPart.addObject(p);
        } else {
            HashMap<String, String> substitution = new HashMap<>();
            substitution.put("content", content.replaceAll("<[.[^>]]*>", ""));
            substitution.put("jc", je.value());
            Object o = XmlUtils.unmarshallFromTemplate(template, substitution);
            mainDocumentPart.addObject(o);
        }
    }

    /**
     * 向文档中添加多个段落的文本
     * @param mainDocumentPart 文档主体
     * @param template 模版
     * @param content 内容
     * @param splitStr 分隔符
     * @throws JAXBException
     */
    public static void addBrArray(MainDocumentPart mainDocumentPart, String template, String content, String splitStr) throws JAXBException {
        if(content == null) {
            return;
        }
        content = content.replace(splitStr, "|");
        String[] cArr = content.split("\\|");
        for (String c : cArr) {
            addObject(mainDocumentPart, template, c);
        }

    }

    /**
     * 设置页边距
     * @param wordPackage 文档
     * @param top
     * @param left
     * @param bottom
     * @param right
     */
    public static void setDocMarginSpace(WordprocessingMLPackage wordPackage, Integer top,
                                         Integer left, Integer bottom, Integer right) {
        SectPr sectPr = getDocSectPr(wordPackage);
        SectPr.PgMar pg = sectPr.getPgMar();
        if (pg == null) {
            pg = objectFactory.createSectPrPgMar();
            sectPr.setPgMar(pg);
        }
        if (top != null) {
            pg.setTop(new BigInteger(String.valueOf(top)));
        }
        if (bottom != null) {
            pg.setBottom(new BigInteger(String.valueOf(bottom)));
        }
        if (left != null) {
            pg.setLeft(new BigInteger(String.valueOf(left)));
        }
        if (right != null) {
            pg.setRight(new BigInteger(String.valueOf(right)));
        }
    }

    public static SectPr getDocSectPr(WordprocessingMLPackage wordPackage) {
        SectPr sectPr = wordPackage.getDocumentModel().getSections().get(0).getSectPr();
        return sectPr;
    }

    /**
     * 本方法创建一个单元格并将给定的内容添加进去.
     * 如果给定的宽度大于0, 将这个宽度设置到单元格.
     * 最后, 将单元格添加到行中.
     * @param wordMLPackage
     * @param row
     * @param content
     * @param width
     */
    public static void addTableCellWithWidth(WordprocessingMLPackage wordMLPackage, Tr row, String content, int width){
        Tc tableCell = objectFactory.createTc();
        tableCell.getContent().add(
                wordMLPackage.getMainDocumentPart().createParagraphOfText(
                        content));

        if (width > 0) {
            setCellWidth(tableCell, width);
        }
        row.getContent().add(tableCell);
    }

    /**
     * 创建一个单元格属性集对象和一个表格宽度对象. 将给定的宽度设置到宽度对象然后将其添加到
     * 属性集对象. 最后将属性集对象设置到单元格中.
     * @param tableCell
     * @param width
     */
    public static void setCellWidth(Tc tableCell, int width) {
        TcPr tableCellProperties = new TcPr();
        TblWidth tableWidth = new TblWidth();
        tableWidth.setW(BigInteger.valueOf(width));
        tableCellProperties.setTcW(tableWidth);
        tableCell.setTcPr(tableCellProperties);
    }

    /**
     * 为表格添加边框
     * @param table
     */
    public static void addBorders(Tbl table) {
        table.setTblPr(new TblPr());
        CTBorder border = new CTBorder();
        border.setColor("auto");
        border.setSz(new BigInteger("4"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);

        TblBorders borders = new TblBorders();
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setTop(border);
        borders.setInsideH(border);
        borders.setInsideV(border);
        table.getTblPr().setTblBorders(borders);
    }

    /**
     * 增加分页符
     *
     * @param documentPart
     */
    public static void addPageBreak(MainDocumentPart documentPart) {
        Br breakObj = new Br();
        breakObj.setType(STBrType.PAGE);

        P paragraph = objectFactory.createP();
        paragraph.getContent().add(breakObj);
        documentPart.getContent().add(paragraph);
    }

    public static Relationship createHeaderPart(
            WordprocessingMLPackage wordprocessingMLPackage, P p)
            throws Exception {

        HeaderPart headerPart = new HeaderPart();
        Relationship rel =  wordprocessingMLPackage.getMainDocumentPart()
                .addTargetPart(headerPart);

        // After addTargetPart, so image can be added properly
        headerPart.setJaxbElement(getHdr(wordprocessingMLPackage, headerPart, p));

        return rel;
    }

    public static Relationship createFooterPart(
            WordprocessingMLPackage wordprocessingMLPackage, String text)
            throws Exception {

        FooterPart headerPart = new FooterPart();
        Relationship rel =  wordprocessingMLPackage.getMainDocumentPart()
                .addTargetPart(headerPart);

        // After addTargetPart, so image can be added properly
        headerPart.setJaxbElement(getFtr(wordprocessingMLPackage, headerPart, text));

        return rel;
    }

    public static void createHeaderReference(
            WordprocessingMLPackage wordprocessingMLPackage,
            Relationship relationship )
            throws InvalidFormatException {

        List<SectionWrapper> sections = wordprocessingMLPackage.getDocumentModel().getSections();

        SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
        // There is always a section wrapper, but it might not contain a sectPr
        if (sectPr==null ) {
            sectPr = objectFactory.createSectPr();
            wordprocessingMLPackage.getMainDocumentPart().addObject(sectPr);
            sections.get(sections.size() - 1).setSectPr(sectPr);
        }

        HeaderReference headerReference = objectFactory.createHeaderReference();
        headerReference.setId(relationship.getId());
        headerReference.setType(HdrFtrRef.DEFAULT);
        sectPr.getEGHdrFtrReferences().add(headerReference);// add header or
        // footer references

    }

    public static void createFooterReference(
            WordprocessingMLPackage wordprocessingMLPackage,
            Relationship relationship )
            throws InvalidFormatException {

        List<SectionWrapper> sections = wordprocessingMLPackage.getDocumentModel().getSections();

        SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
        if (sectPr==null ) {
            sectPr = objectFactory.createSectPr();
            wordprocessingMLPackage.getMainDocumentPart().addObject(sectPr);
            sections.get(sections.size() - 1).setSectPr(sectPr);
        }
        FooterReference footerReference = objectFactory.createFooterReference();
        footerReference.setId(relationship.getId());
        footerReference.setType(HdrFtrRef.DEFAULT);
        sectPr.getEGHdrFtrReferences().add(footerReference);
    }

    public static Hdr getHdr(WordprocessingMLPackage wordprocessingMLPackage,
                             HeaderPart sourcePart, P p) throws Exception {

        Hdr hdr = objectFactory.createHdr();

        hdr.getContent().add(p);
        return hdr;

    }

    public static Ftr getFtr(WordprocessingMLPackage wordprocessingMLPackage,
                             FooterPart sourcePart, String text) throws Exception {
        Ftr ftr = objectFactory.createFtr();
        Document document = DocumentHelper.parseText(text);
        Element rootElement = document.getRootElement();
        List<Element> elements = rootElement.elements();
        for (Element element : elements) {
            String eStr = element.asXML();
            Object o = XmlUtils.unmarshalString(eStr);
            ftr.getContent().add(o);
        }
        return ftr;

    }

    public static P getHP(String up) {

        P headerP = objectFactory.createP();

        RPr rpr = objectFactory.createRPr();
        RFonts rFonts = objectFactory.createRFonts();
        rFonts.setAscii("Times New Roman");
        rFonts.setEastAsia("宋体");
        HpsMeasure sz = new HpsMeasure();
        sz.setVal(new BigInteger("30"));
        rpr.setSz(sz);
        rpr.setRFonts(rFonts);

        R run = objectFactory.createR();
        Text text = objectFactory.createText();
        text.setValue(up);

        run.setRPr(rpr);
        run.getContent().add(text);

        PPr ppr = objectFactory.createPPr();
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.CENTER);
        ppr.setJc(jc);

        PPrBase.PBdr pbdr = objectFactory.createPPrBasePBdr();
        CTBorder ctBorder = objectFactory.createCTBorder();
        ctBorder.setColor("auto");
        ctBorder.setSpace(new BigInteger("1"));
        ctBorder.setSz(new BigInteger("24"));
        ctBorder.setVal(STBorder.SINGLE);

        pbdr.setBottom(ctBorder);
        ppr.setPBdr(pbdr);

        headerP.setPPr(ppr);
        headerP.getContent().add(run);
        return headerP;
    }

    public static P getFP(String up) {

        P footerP = objectFactory.createP();

        RPr rpr = objectFactory.createRPr();
        RFonts rFonts = objectFactory.createRFonts();
        rFonts.setAscii("Times New Roman");
        rFonts.setEastAsia("宋体");
        HpsMeasure sz = new HpsMeasure();
        sz.setVal(new BigInteger("10"));
        rpr.setSz(sz);
        rpr.setRFonts(rFonts);

        R run = objectFactory.createR();
        Text text = objectFactory.createText();
        text.setValue(up);

        run.setRPr(rpr);
        run.getContent().add(text);

        PPr ppr = objectFactory.createPPr();
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.CENTER);
        ppr.setJc(jc);

        footerP.setPPr(ppr);
        footerP.getContent().add(run);
        return footerP;
    }

    /**
     * 创建图片
     * @param wordMLPackage
     * @param bytes
     * @param filenameHint
     * @param altText
     * @param id1
     * @param id2
     * @param cx
     * @return
     * @throws Exception
     */
    public static org.docx4j.wml.P newImage( WordprocessingMLPackage wordMLPackage,
                                             byte[] bytes,
                                             String filenameHint, String altText,
                                             int id1, int id2, long cx) throws Exception {

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        Inline inline = imagePart.createImageInline( filenameHint, altText,
                id1, id2, cx, false);

        org.docx4j.wml.P  p = objectFactory.createP();
        org.docx4j.wml.R  run = objectFactory.createR();
        p.getContent().add(run);
        org.docx4j.wml.Drawing drawing = objectFactory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);

        return p;

    }


}
