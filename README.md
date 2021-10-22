# zeonpad-pdfconvertor
zeonpad-pdfconvertor is library for converting word,excel,ppt,publisher,visio,outlook ,images to PDF. It works for  both Desktop and Server application. It generates the high quality PDF file that are adobe Acrobat  compatible. Supports both 32bit and 64 bit windows os. 

**MAIN FEATURES**<br />
<ul>
  <li>Convert Word documents (.docx, .docm, .doc, .dotx, .dotm, .dot, .rtf, .odt, .wps) to PDF </li>
<li>Convert Excel documents (.xlsx, .xlsm, .xlsb, .xls, .xlts, .xltm, .xlt, .xls, .csv, .xps, .xlsx, .ods) to PDF </li>
<li>Convert PowerPoint documents (.pptx, .pptm, .ppt, .potx, .potm, .pot, .ppsx, .ppsm, .pps, .ppam, .ppa, .odp) to PDF</li>
<li>Convert Publisher documents (.pub) to PDF</li>
<li>Convert Visio documents (.vsd, .vdw, .vss, .vst, .vdx, .svg, ) to PDF</li>
<li>Convert Outlook documents (.msg) to PDF</li>
<li>Convert Images (.png, .bmp, .jpg, .jpeg, .jpe, jfif, .gif, .tif, .tiff) to PDF</li>
<li>Convert Other files (.txt, .xml, .rft, .html, .mhtml, .mht,) to PDF</li>
<li>Convert AutoCAD DXF and DWG (2D) to PDF</li>
  </ul>

**SUPPORTED OPERATING SYSTEM**<br />
<ul>
  <li>Microsoft Windows XP</li>
<li>Microsoft Windows Server 2003 and 2003 R2</li>
<li>Microsoft Windows Vista</li>
<li>Microsoft Windows Server 2008 and 2008 R2</li>
<li>Microsoft Windows 7</li>
<li>Microsoft Windows 8 and 8.1</li>
<li>Microsoft Windows Server 2012 and 2012 R2</li>
</ul>

**SUPPORTED JAVA VERSION**<br />
<ul>
  <li>JDK 1.8 and higher</li>
 </ul>

**JAVA SETUP**<br />
Kindly copy the below given files from <b style="color:red;">jacob</b> folder.<br />
<ul>
  <li>jacob-1.18-x86.dll (JACOB native 32-bit binary)</li>
  <li>jacob-1.18-x64.dll (JACOB native 64-bit binary)</li>
  <li>jacob.jar (JACOB Java library)</li>
</ul>
Add a <b style='color:green;'>JAVA_HOME</b> environment variable. Add  jacob.jar into your classpath.<br />
If you are using 64 bit machine then Copy the jacob-1.18-x64.dll file to java_Home/bin (Example : C:\Program Files\Java\jdk1.8.0_181\jre\bin).<br />
If you are using 32 bit machine then copy Copy the jacob-1.18-x86.dll file to java_Home/bin (Example : C:\Program Files\Java\jdk1.8.0_181\jre\bin).


**Code Example :**<br/>

 **GenericConvertor**<br/> 
This object is used to convert Word, Excel, PowerPoint, Publisher, Visio, Outlook, AutoCad, Images, Text, Html to PDF.<br/>
GenericConvertor genericConv =new GenericConvertor();<br/>
genericConv.convert("C:\\Users\\Desktop\\input.doc","C:\\Users\\Desktop\\output.pdf");<br/><br/>

**WordToPdf**<br/>
This object is used to convert ms-word document to pdf. This object contains many ms-word related features. WordToPdf only works with Word 2007 or higher version. Only Word 2007 requires the free "Save as PDF or XPS" add-in for Office 2007 to be installed. This add-in is available from Microsoft.<br/>
WordToPdf wordToPdf = new WordToPdf();<br/>
wordToPdf.convert("C:\\Users\\Desktop\\input.docx","C:\\Users\\Desktop\\output.pdf");<br/><br/>

**ExcelToPdf**<br/>
This object is used to convert Microsoft Excel to pdf. This object contains many Microsoft Excel related features. ExcelToPdf only works with Microsoft Excel 2007 or higher version. Only Excel 2007 requires the free "Save as PDF or XPS" add-in for Office 2007 to be installed. This add-in is available from Microsoft.<br/>
// Create an Object<br/>
ExcelToPdf excelToPdf = new ExcelToPdf();<br/>
// Convert to Pdf<br/>
excelToPdf.convert("C:\\Users\\Desktop\\input. xls","C:\\Users\\Desktop\\output.pdf");<br/><br/>

**PptToPdf**<br/>
This object is used to convert Microsoft PowerPoint to pdf. PptToPdf only works with Microsoft PowerPoint 2007 or higher version. Only PowerPoint 2007 requires the free "Save as PDF or XPS" add-in for Office 2007 to be installed. This add-in is available from Microsoft.<br/>
// Create an Object<br/>
PptToPdf pptToPdf = new PptToPdf();<br/>
// Convert to Pdf<br/>
pptToPdf.convert("C:\\Users\\Desktop\\input.pptx","C:\\Users\\Desktop\\output.pdf");<br/>

**PublisherToPdf**<br/>
This object is used to convert Microsoft Publisher to pdf. PublisherToPdf only works with Microsoft Publisher 2007 or higher version. Only Publisher 2007 requires the free "Save as PDF or XPS" add-in for Office 2007 to be installed. This add-in is available from Microsoft.<br/>
// Create an Object<br/>
PublisherToPdf publisherToPdf = new PublisherToPdf();<br/>
// Convert to Pdf<br/>
publisherToPdf.convert("C:\\Users\\Desktop\\input.pub", " C:\\Users\\Desktop\\output.pdf");<br/>

**OutlookToPdf**<br/>
This object is used to convert Microsoft Outlook to pdf.<br/>
// Create an Object<br/>
OutlookToPdf outlookToPdf = new OutlookToPdf();<br/>
// Convert to Pdf<br/>
outlookToPdf.convert("C:\\Users\\Desktop\\input.msg", "C:\\Users\\Desktop\\output.pdf");<br/>

**VisioToPdf**<br/>
This object is used to convert Microsoft Visio to pdf. VisioToPdf only works with Microsoft Visio 2007 or higher version. Only Visio 2007 requires the free "Save as PDF or XPS" add-in for Office 2007 to be installed. This add-in is available from Microsoft.<br/>
// Create an Object<br/>
VisioToPdf visioToPdf = new VisioToPdf();<br/>
// Covert to Pdf<br/>
visioToPdf.convert("C:\\Users\\Desktop\\input.vsd", "C:\\Users\\Desktop\\output.pdf");<br/>

**ImageToPDF**<br/>
This object is used to convert images into PDF.<br/>
// Create an Object<br/>
ImageToPDF imageToPdf = new ImageToPDF();<br/>
// Covert to Pdf<br/>
imageToPdf.convert("C:\\Users\\Desktop\\input.jpg", "C:\\Users\\Desktop\\output.pdf");<br/>

