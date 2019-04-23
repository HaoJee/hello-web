package com.archive.ssm.common.util.ExcelPC;

import com.archive.ssm.archive.domain.Archive;
import com.archive.ssm.archive.domain.Department;
import com.archive.ssm.archive.domain.Profession;
import com.archive.ssm.archive.domain.Student;
import com.archive.ssm.archive.service.StudentService;
import com.archive.ssm.archive.service.impl.StudentServiceImpl;
import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Controller;

import java.math.BigInteger;
import java.util.List;

@Component("transDataUtil")
public class TransDataUtil {

    @Autowired
    private StudentService studentService;
//    private StudentService studentService=new StudentServiceImpl();

    public  Student transData(List<Object> rowData){
        Student student=new Student();
//        for (int i=0;i<rowData.size();i++){
//            System.out.println(rowList.get(0)+" rowList "+String.valueOf(rowList.get(0))+" "+String.valueOf(rowList.get(0)).replaceAll("\\s",""));
            //学院
//        System.out.println("rowData"+studentService);
            Integer depid = studentService.findDepidByDepname(String.valueOf(rowData.get(0)).replaceAll("\\s",""));
            if (depid!=null && depid !=0){
                Department dep=new Department();
                dep.setDepartmentId(depid);
                student.setDepartment(dep);
            }
            //专业
            Profession profession=new Profession();
            profession.setProName(String.valueOf(rowData.get(1)).replaceAll("\\s",""));
            profession.setDepartmentId(depid);
            Integer proid = studentService.findProidByPro(profession);
            if (proid != null && proid !=0){
                Profession pro=new Profession();
                pro.setProid(proid);
                student.setPro(pro);
            }
            //班级
            String classStr=String.valueOf(rowData.get(2)).replaceAll("\\s","");
            if (classStr.length()<=10) {
                student.setStuClass(classStr);
            }else{
                student.setStuClass(classStr.substring(0,10));
            }
            //姓名
            String nameStr=String.valueOf(rowData.get(3)).replaceAll("\\s","");
            if (nameStr.length()<=20) {
                student.setStuName(nameStr);
            }else{
                student.setStuName(null);
            }
            //学号
            String  s=rowData.get(4).toString();
            if (rowData.get(4).equals("")){
                student.setStuId(null);
            }else if (!StringUtils.isNumeric(rowData.get(4).toString().replaceAll("\\s",""))){//非数字
                student.setStuId(null);
            }else{
                String s2=((String) rowData.get(4)).replaceAll("\\s","");
                BigInteger stuId=new BigInteger(((String) rowData.get(4)).replaceAll("\\s",""));
                if (stuId.toString().length()<=20) {
                    student.setStuId(stuId);
                }else{
                    student.setStuId(null);
                }
            }
            //性别
            String sexStr=String.valueOf(rowData.get(5)).replaceAll("\\s","");
            if (sexStr.length()<=2) {
                student.setStuSex(sexStr);
            }else{
                student.setStuSex(sexStr.substring(0,2));
            }
            //生源地
            String goStr=String.valueOf(rowData.get(6)).replaceAll("\\s","");
            if (goStr.length()<=50) {
                student.setStuLocation(goStr);
            }else{
                student.setStuLocation(goStr.substring(0,50));
            }
            //派遣证号
            String sendNum=String.valueOf(rowData.get(7)).replaceAll("\\s","");
            if (sendNum.length()<=20){
                student.setStuSendnum(sendNum);
            }else{
                student.setStuSendnum(null);
            }

            Archive archive = new Archive();
            //档案去向
            String archiveGo=String.valueOf(rowData.get(8)).replaceAll("\\s","");
            if(archiveGo.length()<=30){
                archive.setArchiveGo(archiveGo);
            }else{
                archive.setArchiveGo(null);
            }
            student.setArchive(archive);
//        }
        return student;
    }
}
