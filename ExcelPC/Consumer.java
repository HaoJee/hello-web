package com.archive.ssm.common.util.ExcelPC;

import com.archive.ssm.archive.domain.Student;
import com.archive.ssm.archive.mapper.StudentMapper;
import com.archive.ssm.archive.service.StudentService;
import com.archive.ssm.archive.service.impl.StudentServiceImpl;
import org.omg.SendingContext.RunTime;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Controller;

import java.util.List;
import java.util.concurrent.CountDownLatch;

/**
 * 消费者,将仓库中的数据添加到数据库中
 */
@Component
public class Consumer extends  Thread{

    @Autowired
    private DataWarehouse dataWarehouse;
    @Autowired
    private StudentService studentService;

    private CountDownLatch latch;

    public void setLatch(CountDownLatch latch){
        this.latch=latch;
    }
    /**
     *
     */
    @Override
    public void run() {
//        System.out.println("consumer。。");
        while(true){
            List<Student> studentList=dataWarehouse.consumeDatas();
            if(studentList!=null)
                 System.out.println(Thread.currentThread().getName()+" 消费了："+studentList.size()+"条");
            if (studentList==null) {
//                System.out.println("结束了");
                break;
            }else if (studentList.size()>0) {
                Runtime runtime=Runtime.getRuntime();
                System.out.println("自由空间1："+runtime.freeMemory());
                studentService.addStudentBatch(studentList);
                studentList=null;
                System.gc();
                System.out.println("自由空间2："+runtime.freeMemory());

//                System.out.println("数据持久化了");
            }
        }
//        System.out.println("over。。。");
        latch.countDown();
    }
}
