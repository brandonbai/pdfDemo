# Java 生成 PDF 文档方案整理
> 最近项目需要实现PDF下载的功能，由于没有这方面的经验，从网上花了很长时间才找到相关的资料。整理之后，发现有如下几个框架可以实现这个功能。

### 1. 开源框架支持
* iText，生成PDF文档，还支持将XML、Html文件转化为PDF文件；
* Apache PDFBox，生成、合并PDF文档；
* docx4j，生成docx、pptx、xlsx文档，支持转换为PDF格式。

**比较：**
* iText开源协议为AGPL，而其他两个框架协议均为Apache License v2.0。
* 使用PDFBox生成PDF就像画图似的，文字和图像根据页面坐标画上去的，需要根据字数手动换行。
* docx4j用来生成docx文档，提供了将WORD文档转换为PDF文档的功能，并不能直接生成PDF文档。

### 2. 实现方案
—|格式复杂|格式简单
-|-|-
数据量大|docx4j+freemarker|docx4j或PDFBox
数据量小|docx4j|PDFBox

#### 2.1 纯数据生成PDF
> 1.docx4j，适用于生成格式简单或格式复杂且数据量小的PDF文档；
> 2.Apache PDFBox，适用于生成格式简单且数据量小的PDF文档。

1.docx4j
docx4j是一个开源Java库，用于创建和操作Microsoft Open XML（Word docx，Powerpoint pptx和Excel xlsx）文件。它类似于Microsoft的OpenXML SDK，但适用于Java。docx4j使用JAXB来创建内存中的对象表示，程序员需要花时间了解JAXB和Open XML文件结构 。
```java
// word对象
WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
// 文档主体
MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
// 换行符
Br br = objectFactory.createBr();
// 段落
P p = objectFactory.createP();
// 段落设置
PPr ppr = objectFactory.createPPr();
// 文字位置
Jc jc = new Jc();
jc.setVal(je);
ppr.setJc(jc);
// 行设置
RPr rpr = objectFactory.createRPr();
// 字体设置
RFonts rFonts = objectFactory.createRFonts();
rFonts.setAscii("Times New Roman");
rFonts.setEastAsia("宋体");
rpr.setRFonts(rFonts);
// 行
R r = objectFactory.createR();
// 文本
Text text = objectFactory.createText();
text.setValue("这是一段普通文本");
r.setRPr(rpr);
r.getContent().add(br);
r.getContent().add(text);
p.getContent().add(r);
p.setPPr(ppr);
// 添加到正文中
mainDocumentPart.addObject(p);
// 导出
//..
```
2.Apache PDFBox
Apache PDFBox是处理PDF文档的一个开源的Java工具。该项目允许创建新的PDF文档，处理现有文档以及从文档中提取内容的功能。Apache PDFBox还包括几个命令行实用程序。
```java
String formTemplate = "/Users/xiaoming/Desktop/test_pdfbox.pdf";
// 定义文档对象
PDDocument document = new PDDocument();
// 定义一页，大小A4
PDPage page = new PDPage(PDRectangle.A4);
document.addPage(page);
// 获取字体
PDType0Font font = PDType0Font.load(document, new File("/Users/xiaoming/work/tmp/simsun.ttf"));
// 定义页面内容流
PDPageContentStream stream = new PDPageContentStream(document, page);
// 设置字体及文字大小
stream.setFont(font, 12);
// 设置画笔颜色
stream.setNonStrokingColor(Color.BLACK);
// 添加矩形
stream.addRect(29, 797, 100, 14);
// 填充矩形
stream.fill();
stream.setNonStrokingColor(Color.BLACK);
// 文本填充开始
stream.beginText();
// 设置行距
stream.setLeading(18f);
// 设置文字位置
stream.newLineAtOffset(30, 800);
// 填充文字
stream.showText("呵呵");
// 换行
stream.newLine();
stream.showText("哈哈");
stream.newLine();
stream.showText("嘻嘻");
// 文本填充结束
stream.endText();
// 关闭流
stream.close();
// 保存
document.save(formTemplate);
// 释放资源
document.close();
```
#### 2.2 模版+数据生成PDF
> FreeMarker+docx4j，适用于生成格式复杂且数据量大的PDF文档

Apache FreeMarker是一个模板引擎，用于根据模板和更改数据生成文本输出（HTML网页，电子邮件，配置文件，源代码等）。模板是用FreeMarker模板语言（FTL）编写的，是一种简单的专用语言。

Office2003以上，Word是可以以XML文本格式存储的。先将要生成的PDF转换为Word文档 ，再将其保存为XML文本，通过模版引擎将数据填充到XML文本中，最后再反向转换为PDF文档。简单来说就是PDF->Word->XML->Word->PDF的流程。

步骤|描述|工具
-|-|-
1|word -> xml|手动
2|xml -> ftl|手动，参考[《XML格式Word文档常用标签介绍》](https://www.jianshu.com/p/b7d7ba967383)
3|ftl + obj = xml|freemarker
4|xml -> pdf|docx4j

##### 步骤
* 1 把pdf文档对应的word(docx)制作出来
![简历.png](https://user-gold-cdn.xitu.io/2019/4/23/16a4828f8c21ce14?w=532&h=686&f=png&s=58612)
* 2 把word文档另存为xml文件
![另存为xml](https://user-gold-cdn.xitu.io/2019/4/23/16a4828f8c329db0?w=1022&h=384&f=png&s=44695)
* 3 将xml文件制作为freemarker模版(ftl)文件
![制作模版文件](https://user-gold-cdn.xitu.io/2019/4/23/16a4828f8c4c7af3?w=704&h=614&f=png&s=79086)
* 4 将数据和ftl文件组装为xml文本
```java
Map<String, Object> map = new HashMap<>();
map.put("name", "小明");
map.put("address", "北京市朝阳区");
map.put("email", "xiaoming@abc.com");
StringWriter stringWriter = new StringWriter();
BufferedWriter writer = new BufferedWriter(stringWriter);
template.process(map, writer);
String xmlStr = stringWriter.toString();
```
* 5 使用docx4j将xml文本加载为word文档对象
```java
ByteArrayInputStream in = new ByteArrayInputStream(xmlStr.getBytes());
WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);
```
* 6 使用docx4j将word文档转存为pdf文档
```java
String outputfilepath = "/Users/xiaoming/简历.pdf";
FileOutputStream os = new FileOutputStream(new File(outputFilePath));
FOSettings foSettings = Docx4J.createFOSettings();
foSettings.setWmlPackage(wordMLPackage);
Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
// Docx4J.toPDF(wordMLPackage, new FileOutputStream(new File(outputfilepath)));
```
#### 2.3 Word转PDF
> docx4j
```java
WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(new File("abc.docx"));
Mapper fontMapper = new IdentityPlusMapper();  
// fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));  
mlPackage.setFontMapper(fontMapper);  
OutputStream os = new java.io.FileOutputStream("abc.pdf");    
FOSettings foSettings = Docx4J.createFOSettings();  
foSettings.setWmlPackage(mlPackage);  
Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);  
```
#### 2.4 合并多个PDF
> Apache PDFBox，将多个PDF文档合并
```java
String folderName = "/Users/xiaoming/pdfs";
String destPath = "/Users/xiaoming/all.pdf";
PDFMergerUtility mergePdf = new PDFMergerUtility();
String[] filesInFolder = getFiles(folderName);
Arrays.sort(filesInFolder, new Comparator<String>() {
      @Override
      public int compare(String o1, String o2) {
          return o1.compareTo(o2);
      }
});
for (int i = 0; i < filesInFolder.length; i++) {
     mergePdf.addSource(folderName + File.separator + filesInFolder[i]);
}
mergePdf.setDestinationFileName(destPath);
mergePdf.mergeDocuments(MemoryUsageSetting.setupMainMemoryOnly());
```
