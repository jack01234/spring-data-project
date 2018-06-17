package com.wmli.system.test;

import com.wmli.system.manager.model.Customer;
import com.wmli.system.manager.util.ReadExcel;

import java.util.List;

/**
 * @Auther: Administrator
 * @Date: 2018/6/16 0016 17:38
 * @Description:
 */
public class ReadExcelFileTest {
    public static void main(String[] args) {
        List<Customer> excelInfo = ReadExcel.getExcelInfo("E://test_scope/test.xlsx");
        System.out.println("custormer result "+excelInfo);
    }
}
