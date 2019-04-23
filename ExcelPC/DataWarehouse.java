package com.archive.ssm.common.util.ExcelPC;

import com.archive.ssm.archive.domain.Student;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;

/**
 * 数据仓库
 */
@Component
@Scope("singleton")
public class DataWarehouse {
    private static  int MAX_SIZE=5;
    public static boolean isFinish=false;
    private BlockingQueue<List<Student>> dw=new LinkedBlockingQueue<>();

    public synchronized void addDatas(List<Student> datas){
        while(dw.size()>=MAX_SIZE){
            try {
                this.wait();//使用while循环，来判断线程唤醒后，是继续等待，还是执行操作
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
        dw.add(datas);
        this.notifyAll();//唤醒所有等待的线程，主要唤醒消费者，防止死锁
    }

    public synchronized List consumeDatas(){
        while(dw.size()<=0){
            if(isFinish){
               this.notifyAll();
               return null;
            }
            try {
                this.wait();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
        System.out.println("dw-前："+dw.size());
        List<Student> list=dw.poll();
//        System.out.println("取出数据："+list.size());
        System.out.println("dw-后："+dw.size());
        this.notifyAll();//唤醒所有在等待的线程，主要唤醒生产者
        return list;
    }

    public BlockingQueue<List<Student>> getData(){
        return dw;
    }
}
