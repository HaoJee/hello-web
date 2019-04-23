package com.archive.ssm.common.util.ExcelPC;

import com.archive.ssm.archive.domain.Student;
import com.archive.ssm.common.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Controller;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;

/**
 * 生产者:通过流来获取excel数据
 */
@Component
public class Producter extends  Thread{

    @Autowired
    private DataWarehouse dataWarehouse;

    private InputStream inputStream;//输入流
    private String fileName;//文件名
    private CountDownLatch latch;//计时锁

    @Autowired
    private TransDataUtil transDataUtil;
//    private TransDataUtil transDataUtil=new TransDataUtil();
//    public Producter(DataWarehouse dataWarehouse,InputStream inputStream,String fileName) {
//        this.dataWarehouse = dataWarehouse;
//        this.inputStream=inputStream;
//        this.fileName=fileName;
//    }


    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public void setLatch(CountDownLatch latch){
        this.latch=latch;
    }

    @Override
    public void run() {
//        System.out.println("product。。");
        try {
            List<Student> list=null;
            Workbook wb= ExcelUtil.getWorkBook(inputStream,fileName);
            if (wb!=null){
                Sheet sheet=null;
                Row row=null;
                Cell cell=null;
                list=new ArrayList<Student>();

                for (int i=0;i<wb.getNumberOfSheets();i++){
                    sheet=wb.getSheetAt(i);
                    if (sheet==null){continue;}
    //                sheet.isColumnHidden();
                    //遍历每一行
                    for (int j=sheet.getFirstRowNum();j<=sheet.getLastRowNum();j++){
                        row=sheet.getRow(j);
                        //去除表头和空行
                        if (row==null || j==sheet.getFirstRowNum()){ continue; }
                        //去除被隐藏的行
                        if (row.getZeroHeight()){
                            continue;
                        }
                        //该行的列数
                        // Integer columns=row.getPhysicalNumberOfCells();//非空列的个数
                        Integer columns= (int)row.getLastCellNum();//最后一个非空列是第几个
                        //如果列数大于规定的数，则跳过该行
                        if (columns>9){
                            continue;
                        }else if (columns<9){
                            columns=9;
                        }
                        //遍历每一列
                        List<Object> rowData=new ArrayList<>();
                        for (int m=0;m<columns;m++){
                            cell=row.getCell(m);

                            if (cell!=null) {
//                                System.out.println("cell:"+cell+"  value:"+ExcelUtil.getCellValue(cell));
                                rowData.add(ExcelUtil.getCellValue(cell));
                            }else{
                                rowData.add("");
                            }
                        }
//                        System.out.println(columns+"  row:"+rowData+" "+rowData.size());
                        //将rowData转化成对象Student放到list中
                        Student t=transDataUtil.transData(rowData);
                        list.add(t);

                        if(list.size()>=1000){
                            dataWarehouse.addDatas(new ArrayList<>(list));
                            System.out.println("生产了："+list.size()+"条");
//                            System.out.println(list);
                            list.clear();
                        }
                    }
                }

                if (list.size()>0){
                    dataWarehouse.addDatas(list);
//                    list=new ArrayList<>();
                }
            }
            DataWarehouse.isFinish=true;
//            System.out.println("生产数据结束");
        } catch (Exception e) {
            e.printStackTrace();
        }
        latch.countDown();
    }
}
