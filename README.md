# PDFConverterUtility
Converter Utility ("jpeg","jpg","gif","tif","png","doc","docx","xls","xlsx","ppt","pptx") format to PDF document.

Steps to RUN the Utility
------------------------
1. Java 1.7 or higher version and Maven installed system 
2. Checkout project from https://github.com/ramesh-iyer/PDFConverterUtility.git
3. Unzip into any drive
4. RUN command: "mvn clean install"
5. A JAR file is created under Project/target location
6. Can test standalone by executing "java -jar pdf-converter-utility-1.0.0 inFilePath, outFileDir

   Example-
   inFilePath = "C:\Users\USER\Desktop\resources\resume.jpg"
   outFileDir = <Location where the PDF file need to create>  "C:\Users\USER\Desktop\resources"
   
7. Can include the created JAR file with any Java application and invoke as below-
   
   PdfConverter converter = new PdfConverter();
   converter.convertToPdf(inFilePath, outFileDir);
   
                                               ---------------------------------------
   
