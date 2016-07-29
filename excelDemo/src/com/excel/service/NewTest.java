package com.excel.service;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * @author qchen
 * @version
 * @datetime 2016年7月18日 下午6:22:53
 */
public class NewTest {
	public static void main(String[] args) {
		List<String> titleStrs = new ArrayList<String>();
		titleStrs.add("ID");
		titleStrs.add("姓名");
		titleStrs.add("联系电话");
		titleStrs.add("地址");
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		
	    //  创建表  
	    HSSFSheet sheet = workbook.createSheet("用户信息");   
	    HSSFHeader header = sheet.getHeader();   
        header.setCenter("用户表");
        
        //创建一行
        HSSFRow headerRow = sheet.createRow(0);
        int i = 0;
        for(String titleStr:titleStrs){
        	HSSFCell headerCell = headerRow.createCell(i);
        	headerCell.setCellValue(titleStr);
        	
        	i++;
        }
        
        
        List<String> dataStrs1 = new ArrayList<String>();
        dataStrs1.add("001");
        dataStrs1.add("小陈");
        dataStrs1.add("13622221111");
        dataStrs1.add("江西");
        
        HSSFRow dataRow1 = sheet.createRow(1);
        for(String dataStr:dataStrs1){
        	HSSFCell dataCell = dataRow1.createCell(i);
        	dataCell.setCellValue(dataStr);
        	
        	i++;
        }
        
        
        String fileName = "D:\\javaexcel\\用户信息2.xls";   
	    FileOutputStream fos = null;   
	    try {   
	        fos = new FileOutputStream(fileName);   
	        sheet.setGridsPrinted(true);   
	    	HSSFFooter footer = sheet.getFooter();   
	    	footer.setRight("Page " + HSSFFooter.page() + " of " + HSSFFooter.numPages());   
	    	workbook.write(fos);
	        
	        JOptionPane.showMessageDialog(null, "表格已成功导出到 : "+fileName);   
	    } catch (Exception e) {   
	        JOptionPane.showMessageDialog(null, "表格导出出错，错误信息 ："+e+"\n错误原因可能是表格已经打开。");   
	        e.printStackTrace();   
	    } finally {   
	        try {   
	            fos.close();   
	        } catch (Exception e) {   
	            e.printStackTrace();   
	        }   
	    }
	}
}
