package com.wmli.system.manager.util;

import com.wmli.system.manager.model.Customer;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @Auther: Administrator
 * @Date: 2018/6/16 0016 15:34
 * @Description:
 */
@Slf4j
@Data
public class ReadExcel {
    //总行数
    private static int totalRows = 0;
    //总条数
    private static int totalCells = 0;
    //错误信息接收器
    private static String errorMsg;
    //构造方法
    public ReadExcel(){}


    /**
     * 验证EXCEL文件
     * @param filePath
     * @return
     */
    public static boolean validateExcel(String filePath){
        if (filePath == null || !(WDWUtil.isExcel2003(filePath) || WDWUtil.isExcel2007(filePath))){
            errorMsg = "文件名不是excel格式";
            return false;
        }
        return true;
    }

    /**
     * 读EXCEL文件，获取客户信息集合
     * @return
     */
    public static List<Customer> getExcelInfo(String fileName){

        //把spring文件上传的MultipartFile转换成CommonsMultipartFile类型
//        CommonsMultipartFile cf= (CommonsMultipartFile)Mfile; //获取本地存储路径
        File file = new  File(fileName);
        //新建一个文件
//        File file1 = new File("D:\\fileupload" + new Date().getTime() + ".xlsx");
        //将上传的文件写入新建的文件中
//        try {
//            cf.getFileItem().write(file1);
//        } catch (Exception e) {
//            e.printStackTrace();
//        }

        //初始化客户信息的集合
        List<Customer> customerList=new ArrayList<>();
        //初始化输入流
        InputStream is = null;
        try{
            //验证文件名是否合格
            if(!validateExcel(fileName)){
                return null;
            }
            //根据文件名判断文件是2003版本还是2007版本
            boolean isExcel2003 = true;
            if(WDWUtil.isExcel2007(fileName)){
                isExcel2003 = false;
            }
            //根据新建的文件实例化输入流
            is = new FileInputStream(file);
            //根据excel里面的内容读取客户信息
            customerList = getExcelInfo(is, isExcel2003);
            is.close();
        }catch(Exception e){
            e.printStackTrace();
        } finally{
            if(is !=null)
            {
                try{
                    is.close();
                }catch(IOException e){
                    is = null;
                    e.printStackTrace();
                }
            }
        }
        return customerList;
    }
    /**
     * 根据excel里面的内容读取客户信息
     * @param is 输入流
     * @param isExcel2003 excel是2003还是2007版本
     * @return
     * @throws IOException
     */
    public static List<Customer> getExcelInfo(InputStream is, boolean isExcel2003){
        List<Customer> customerList=null;
        try{
            /** 根据版本选择创建Workbook的方式 */
            Workbook wb = null;
            //当excel是2003时
            if(isExcel2003){
                wb = new HSSFWorkbook(is);
            }
            else{//当excel是2007时
                wb = new XSSFWorkbook(is);
            }
            //读取Excel里面客户的信息
            customerList=readExcelValue(wb);
        }
        catch (IOException e)  {
            e.printStackTrace();
        }
        return customerList;
    }
    /**
     * 读取Excel里面客户的信息
     * @param wb
     * @return
     */
    private static List<Customer> readExcelValue(Workbook wb){
        int numbers = wb.getNumberOfSheets();
        List<Customer> customerList=new ArrayList<>();
        for (int i = 0; i<numbers;i++) {
            //得到第一个shell
            Sheet sheet=wb.getSheetAt(i);

            //得到Excel的行数
            totalRows=sheet.getPhysicalNumberOfRows();

            //得到Excel的列数(前提是有行数)
            if(totalRows>=1 && sheet.getRow(0) != null){
                totalCells=sheet.getRow(0).getPhysicalNumberOfCells();
            }


            Customer customer;
            //循环Excel行数,从第二行开始。标题不入库
            for(int r=1;r<totalRows;r++){
                Row row = sheet.getRow(r);
                if (row == null) continue;
                customer = new Customer();

                //循环Excel的列
                for(int c = 0; c <totalCells; c++){
                    row.getCell(c).setCellType(Cell.CELL_TYPE_STRING);
                    Cell cell = row.getCell(c);
                    if (null != cell){
                        if(c==0){//第一列不读
                        }else if(c==1){
                            customer.setName(cell.getStringCellValue());//客户名称
                        }else if(c==2){
                            customer.setSimpleName(cell.getStringCellValue());//客户简称
                        }else if(c==3){
                            customer.setTrade(cell.getStringCellValue());//行业
                        }else if(c==4){
                            customer.setSource(cell.getStringCellValue());//客户来源
                        }else if(c==5){
                            customer.setAddress(cell.getStringCellValue());//地址
                        }else if(c==6){
                            customer.setRemark(cell.getStringCellValue());//备注信息
                        }
                    }
                }
                //添加客户
                customerList.add(customer);
            }
        }
        return customerList;
    }
}

