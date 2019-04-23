package com.archive.ssm.common.util.ExcelPC;

import com.archive.ssm.archive.domain.Student;
import com.archive.ssm.common.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;

/**
 * 生产者:通过sheet和size来获取excel数据
 */
@Component
public class Producter2 extends  Thread{

    @Autowired
    private DataWarehouse dataWarehouse;

    private Sheet sheet;//输入流
    private int start;//开始行
    private int size;//数据源大小

    private static CountDownLatch latch;//计时锁

    @Autowired
    private TransDataUtil transDataUtil;


    public void setLatch(CountDownLatch latch){
        this.latch=latch;
    }
    public void setSheet(Sheet sheet){this.sheet=sheet;}
    public void setStart(int start){
        this.start=start;
    }
    public void setSize(int size){
        this.size=size;
    }

    @Override
    public void run() {
//        System.out.println("product。。");
        Row row=null;
        Cell cell=null;
        List<Student> list=new ArrayList<>();
        for (int i=this.start;i<this.start+size;i++) {
                //遍历每一行
                row=sheet.getRow(i);
                //去除表头和空行
                if (row==null || i==sheet.getFirstRowNum()){ continue; }
                //去除被隐藏的行
                if (row.getZeroHeight()){
                    continue;
                }
                //该行的列数
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
                    System.out.println(Thread.currentThread().getName()+" 生产了："+list.size()+"条");
//                            System.out.println(list);
                    list.clear();
                }
        }

        if (list.size()>0) {
            dataWarehouse.addDatas(list);
        }
        list=null;
            DataWarehouse.isFinish=true;
            System.out.println("生产数据结束");

        latch.countDown();
    }
}
