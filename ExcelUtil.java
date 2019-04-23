package com.archive.ssm.common.util;

import com.archive.ssm.archive.domain.Department;
import com.archive.ssm.archive.domain.Profession;
import javafx.geometry.HorizontalDirection;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.time.DateFormatUtils;
import org.apache.commons.lang.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.springframework.web.servlet.mvc.method.annotation.ExceptionHandlerExceptionResolver;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;

/**
 * 解析excel工具类
 * 1、获取excel文件类型
 * 2、格式化单元格数据
 * 3、excel数据ToList集合
 * 导入数据
 */
public class ExcelUtil {
    private static  String excel2003L=".xls";
    private static  String excel2007U=".xlsx";

    /**
     * 根据excel文件后缀，创建相应版本的Workbook工作簿对象
     * @param in
     * @param fileName
     * @return
     * @throws Exception
     */
    public  static  Workbook getWorkBook(InputStream in,String fileName) throws Exception{
        Workbook wb=null;
        if (in != null) {
            if (fileName != null && !fileName.equals("")) {
                String type = fileName.substring(fileName.lastIndexOf("."));
                if (type.equals(excel2003L)) {
                    wb = new HSSFWorkbook(in);
                } else if (type.equals(excel2007U)) {
                    wb = new XSSFWorkbook(in);
                }
            }
        }
        System.out.println("sleep:60000");
        Thread.sleep(60000);
        return wb;
    }

    /**
     * 格式化传入的数据
     */
    public  static Object getCellValue(Cell cell){
        Object value=null;
        DecimalFormat df=new DecimalFormat("0");//格式化字符类型数字
        SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd");
        DecimalFormat df2=new DecimalFormat("0.00");//格式化数字
        switch (cell.getCellType()){
            case Cell.CELL_TYPE_STRING:
                value=cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if ("General".equals(cell.getCellStyle().getDataFormatString())){
                    value=df.format(cell.getNumericCellValue());
                }else if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())){
                    value=sdf.format(cell.getDateCellValue());
                }else{
                    value=df2.format(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                value=cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_BLANK:
                value="";
                break;
            default:
                    break;
        }
        return  value;
    }

    /**
     * 解析excel文件
     *根据文件名，获取文件
     * 遍历每一个sheet，遍历每一行（list）遍历每一列cell
     */
    public static List<List<Object>> getListByExcel(InputStream in,String fileName) throws Exception{
        List<List<Object>> list=null;
        Workbook wb=getWorkBook(in,fileName);
        if (wb!=null){
            Sheet sheet=null;
            Row row=null;
            Cell cell=null;
            list=new ArrayList<List<Object>>();

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
                        /*if(cell.getCellStyle().getHidden()){
                            System.out.println(j+"行"+m+"列 被隐藏");
                            continue;
                        }*/
                        if (cell!=null) {
//                            System.out.println("cell:"+cell+"  value:"+getCellValue(cell));
                            rowData.add(getCellValue(cell));
                        }else{
                            rowData.add("");
                        }
                    }
                    //将每一行的数据放到list中
                    list.add(rowData);
                }
            }
        }
        return list;
    }

    //递归设计、
    public static Queue<List<Object>> getListByExcel2(InputStream in,String fileName) throws Exception{
        Queue<List<Object>> list=null;
        Workbook wb=getWorkBook(in,fileName);
        if (wb!=null){
            Sheet sheet=null;
            Row row=null;
            Cell cell=null;
            list=new LinkedList<>();

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
                        /*if(cell.getCellStyle().getHidden()){
                            System.out.println(j+"行"+m+"列 被隐藏");
                            continue;
                        }*/
                        if (cell!=null) {
                            rowData.add(getCellValue(cell));
                        }else{
                            rowData.add("");
                        }
                    }
                    //将每一行的数据放到list中
                    list.add(rowData);
                }
            }
        }
        return list;
    }

    /**
     * 生成excel文件：XSSF生成的是.xlsx文件，单个sheet最多只能容纳104万条数据，该方法适用于104万条数据以内的操作，如果超过需要根据数据量分次导出
     * @param clazz 数据所属的对象类型（用于通过反射的方式调用对象的getXXX方法，从而获得属性）
     * @param map 存放excel的标题以及对应的属性字段
     * @param objs 是需要导出的数据，是student的list集合
     * @param sheetname 生成的excel工作表名称
     * @return 返回一个工作簿
     * @throws IllegalAccessException
     * @throws IntrospectionException
     * @throws InvocationTargetException
     */
    public static XSSFWorkbook createExcelFile(Class clazz,Map<Integer,List<ExcelBean>> map,List objs,String sheetname) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        XSSFWorkbook workbook=new XSSFWorkbook();

        int sheetSize=10000;//每个sheet的大小
        int counts=objs.size()/sheetSize;//sheet的个数
        int other=objs.size()%sheetSize;//最后一页不够size
        if (other>0){
            counts+=1;
        }
        //没有数据，就创建一个空sheet
        if (counts==0){
            XSSFSheet sheet=workbook.createSheet(sheetname);
            createFont(workbook);//设置表头和内容的样式
            createHeader(sheet, map);//设置表头信息
        }

        for(int i=0;i<counts;i++) {
            List objsTemp=objs.subList(0+i*sheetSize, Math.min(objs.size(), sheetSize*(i+1)));
            XSSFSheet sheet=workbook.createSheet(sheetname+"("+(i+1)+")");
            createFont(workbook);//设置表头和内容的样式
            createHeader(sheet, map);//设置表头信息
            createTableRows(sheet, map, objsTemp, clazz);//将数据放入excel的每一个cell中
        }
        return workbook;
    }
  /*  public static XSSFWorkbook createExcelFile(Class clazz,Map<Integer,List<ExcelBean>> map,List objs,String sheetname) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet(sheetname);
        createFont(workbook);//设置表头和内容的样式
        createHeader(sheet, map);//设置表头信息
        createTableRows(sheet, map, objs, clazz);//将数据放入excel的每一个cell中
        return workbook;
    }*/
    /**
     * 设置字体
     */
    private static XSSFCellStyle fontStyle1;
    private static XSSFCellStyle fontStyle2;
    public  static void createFont(XSSFWorkbook workbook){
        fontStyle1=workbook.createCellStyle();
        XSSFFont font1 = workbook.createFont();
        font1.setFontName("黑体");
        font1.setFontHeightInPoints((short) 14);// 设置字体大小
        fontStyle1.setFont(font1);
        fontStyle1.setAlignment(HorizontalAlignment.CENTER);

        // 内容
        fontStyle2=workbook.createCellStyle();
        XSSFFont font2 = workbook.createFont();
        font2.setFontName("宋体");
        font2.setFontHeightInPoints((short) 10);// 设置字体大小
        fontStyle2.setFont(font2);
        fontStyle2.setAlignment(HorizontalAlignment.CENTER); // 居中


    }

    /**
     * 生成表头
     */
    public  static final void createHeader(XSSFSheet sheet, Map<Integer,List<ExcelBean>> map){
        int startIndex=0;
        int endIndex=0;
        for (Map.Entry<Integer,List<ExcelBean>> entry:map.entrySet()) {
            XSSFRow row=sheet.createRow(entry.getKey());
            List<ExcelBean> excels=entry.getValue();
            for (int i=0;i<excels.size();i++){
                XSSFCell cell=row.createCell(i);
                cell.setCellValue(excels.get(i).getHeadTextName());
                if (excels.get(i).getCellStyle()!=null){
                    cell.setCellStyle(excels.get(i).getCellStyle());
                }
                cell.setCellStyle(fontStyle1);

            }
        }


    }

    /**
     * 生成每一行
     * @param sheet
     * @param map
     * @param objs
     * @param clazz
     */
    public  static  void createTableRows(XSSFSheet sheet, Map<Integer,List<ExcelBean>> map,List objs,Class clazz) throws IntrospectionException, InvocationTargetException, IllegalAccessException {
        Integer rowIndex=map.size();
        Integer maxKey=0;
        List<ExcelBean> ems = new ArrayList<>();
        for (Map.Entry<Integer, List<ExcelBean>> entry : map.entrySet()) {
            if (entry.getKey() > maxKey) {
                maxKey = entry.getKey();
            }
        }
        ems=map.get(maxKey);
        List<Integer> widths = new ArrayList<Integer>(ems.size());
        for (Object obj:objs) {
            XSSFRow row = sheet.createRow(rowIndex);
            for (int i = 0; i < ems.size(); i++) {
                ExcelBean excelBean = ems.get(i);
                //获得getXXX()方法
                PropertyDescriptor propertyDescriptor = new PropertyDescriptor(excelBean.getPropertyName(), clazz);
                Method getMethod = propertyDescriptor.getReadMethod();
                Object getValue = getMethod.invoke(obj);//obj对象调用get方法获得属性值
                String value = "";
                if (getValue != null) {
                    if (getValue instanceof Department){
                        value=((Department) getValue).getDepartmentName();
                        if (value==null){
                            value="";
                        }
                    }else if (getValue instanceof Profession){
                        value=((Profession) getValue).getProName();
                        if (value==null){
                            value="";
                        }
                    }else if (getValue instanceof Date) {
                        value = DateFormatUtils.format((Date) getValue, "yyyy-MM-dd");
                    } else if(getValue instanceof BigDecimal){
                        NumberFormat nf = new DecimalFormat("#,##0.00");
                        value=nf.format((BigDecimal)getValue).toString();
                    } else if ((getValue instanceof Integer) && (Integer.valueOf(getValue.toString()) < 0)) {
                        value = "--";
                    } else {
                        value = getValue.toString();
                    }

                }
                XSSFCell cell = row.createCell(i);
                cell.setCellValue(value);
                cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                cell.setCellStyle(fontStyle2);
                // 获得最大列宽
                if (value==null) {
                   value="";
                }
                int width = value.getBytes().length * 300;
                // 还未设置，设置当前
                if (widths.size() <= i) {
                    widths.add(width);
                    continue;
                }
                // 比原来大，更新数据
                if (width > widths.get(i)) {
                    widths.set(i, width);

                }
            }
            rowIndex++;
        }
        // 设置列宽
        for (int index = 0; index < widths.size(); index++) {
            Integer width = widths.get(index);
            width = width < 2500 ? 2500 : width + 300;
            width = width > 10000 ? 10000 + 300 : width + 300;
            sheet.setColumnWidth(index, width);
        }
    }
}
