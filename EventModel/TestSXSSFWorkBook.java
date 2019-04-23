package com.archive.ssm.common.util.EventModel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

public class TestSXSSFWorkBook {
    public static void main(String[] args){
        TestSXSSFWorkBook test=new TestSXSSFWorkBook();
        test.test();
//        test.test2();
    }
    //SXSSF 流用戶模式
    public  void test(){
        try {
            long startTime=System.currentTimeMillis();
            SXSSFWorkbook sxssfWorkbook=new SXSSFWorkbook(100);
            sxssfWorkbook.createSheet("testSAX");
            SXSSFSheet sheet=sxssfWorkbook.getSheetAt(0);
            for (int rowNum=0;rowNum<1000000;rowNum++){
                SXSSFRow row=sheet.createRow(rowNum);
                SimpleDateFormat dateFormat=new SimpleDateFormat("yyyy-MM-dd");
                row.createCell(0).setCellValue(dateFormat.format(new Date()));
                row.createCell(1).setCellValue("星星");
                row.createCell(2).setCellValue(12.0);
//                Calendar calendar=new GregorianCalendar();
                row.createCell(3).setCellValue("wide");
                row.createCell(4).setCellValue(1);
            }

            // Rows with rownum < 900 are flushed and not accessible
            for(int rownum = 0; rownum <1000000-100; rownum++){
                Assert.assertNull(sheet.getRow(rownum));
            }

            // ther last 100 rows are still in memory
            for(int rownum = 1000000-100; rownum < 1000000; rownum++){
                Assert.assertNotNull(sheet.getRow(rownum));
            }
            OutputStream outputStream=new FileOutputStream("C:/Users/DELL/Downloads/sax.xlsx");
            sxssfWorkbook.write(outputStream);
            outputStream.close();
            sxssfWorkbook.dispose();
            long endTime=System.currentTimeMillis();

            System.out.println(sxssfWorkbook.getNumberOfSheets());
            System.out.println("SXSSFWorkbook  data size: 100w； time cost:"+(double)(endTime-startTime)+"s");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void test2(){
            try {
                long startTime=System.currentTimeMillis();
                XSSFWorkbook xssfWorkbook=new XSSFWorkbook();
                xssfWorkbook.createSheet("testSAX");
                XSSFSheet sheet=xssfWorkbook.getSheetAt(0);
                for (int rowNum=0;rowNum<1000000;rowNum++){
                    Row row=sheet.createRow(rowNum);
                    SimpleDateFormat dateFormat=new SimpleDateFormat("yyyy-MM-dd");
                    row.createCell(0).setCellValue(dateFormat.format(new Date()));
                    row.createCell(1).setCellValue("星星");
                    row.createCell(2).setCellValue(12.0);
//                Calendar calendar=new GregorianCalendar();
                    row.createCell(3).setCellValue("wide");
                    row.createCell(4).setCellValue(1);
                }


                OutputStream outputStream=new FileOutputStream("C:/Users/DELL/Downloads/sax2.xlsx");
                xssfWorkbook.write(outputStream);
                outputStream.close();
                long endTime=System.currentTimeMillis();
                System.out.println(xssfWorkbook.getNumberOfSheets());
                System.out.println("XSSFWorkbook  data size: 100w； time cost:"+(double)(endTime-startTime)+"s");
            } catch (Exception e) {
                e.printStackTrace();
            }
    }
}
