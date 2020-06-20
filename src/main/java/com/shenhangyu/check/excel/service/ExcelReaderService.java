/**
 *版权所有©微信公众号|视频号:深航渔
 */
package com.shenhangyu.check.excel.service;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExcelReaderService {
	private POIFSFileSystem fs;
	private HSSFWorkbook wb;
	private HSSFSheet sheet;
	private HSSFRow row;
	public String[] readExcelTitle(InputStream is){
		try {
			this.fs = new POIFSFileSystem(is);
			this.wb = new HSSFWorkbook(this.fs);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		this.sheet = this.wb.getSheetAt(0);
		this.row = this.sheet.getRow(0);
		int colNum = this.row.getPhysicalNumberOfCells();
		System.out.println("colNum:"+colNum);
		String[] title = new String[colNum];
		for(int i=0; i<colNum; i++){
			title[i] = getCellFormatValue(this.row.getCell((short)i));
		}
		return title;
	}
	public Map<Integer,String> readExcelContent(InputStream is){
		Map<Integer,String> content = new HashMap<Integer,String>();
		String str = "";
		try {
			this.fs = new POIFSFileSystem(is);
			this.wb = new HSSFWorkbook(this.fs);
		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		this.sheet = this.wb.getSheetAt(0);
		int rowNum = this.sheet.getLastRowNum();
		this.row = this.sheet.getRow(0);
		int colNum = this.row.getPhysicalNumberOfCells();
		for(int i=1; i<=rowNum; i++){
			this.row = this.sheet.getRow(i);
			int j=0;
			while(j<colNum){
				str = str + getCellFormatValue(this.row.getCell((short)j)).trim()+" ";
				j++;
			}
			content.put(Integer.valueOf(i), str);
			str = "";
		}
		return content;
	}
	
	public List<HashMap<Integer,String>> readExcelContent2(InputStream is){
		List<HashMap<Integer,String>> list = new ArrayList<HashMap<Integer,String>>();
		try {
			this.fs = new POIFSFileSystem(is);
			this.wb = new HSSFWorkbook(this.fs);
		} catch (IOException e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		this.sheet = this.wb.getSheetAt(0);
		
		
		int rowNum = this.sheet.getLastRowNum();
		this.row = this.sheet.getRow(0);
		int colNum = this.row.getPhysicalNumberOfCells();
		
		for(int i=1; i<=rowNum; i++){
			this.row = this.sheet.getRow(i);
			int j=0;
			HashMap<Integer,String> map = new HashMap<Integer,String>();
			while(j<colNum){
				String str = getCellFormatValue(this.row.getCell((short)j));
				
				if(null != str){
					str = str.trim();
				}
				map.put(Integer.valueOf(j), str);
				j++;
			}
			list.add(map);
		}
		return list;
	}
	
	private String getStringCellValue(HSSFCell cell){
		if( null == cell){
			return "";
		}
		String strCell = "";
		switch(cell.getCellType()){
		case 1:
			strCell = cell.getStringCellValue();
			break;
		case 0:
			strCell = String.valueOf(cell.getNumericCellValue());
			break;
		case 4:
			strCell = String.valueOf(cell.getBooleanCellValue());
			break;
		case 3:
			strCell = "";
			break;
		case 2:
		default:
			strCell = "";
		}
		if((strCell.equals("")) || ( null == strCell)){
			return "";
		}
		return strCell;
	}
	
	private String getDateCellValue(HSSFCell cell){
		String result = "";
		try {
			int cellType = cell.getCellType();
			if(0 == cellType){
				Date date = cell.getDateCellValue();
				result = date.getYear() + 1900 + "-" + (date.getMonth()+1) + "-" +date.getDate();
			}else if(1 == cellType){
				String date = getStringCellValue(cell);
				result = date.replace("[年月]", "-").replace("日", "-").trim();
			}else if(3 == cellType){
				result = "";
			}
		} catch (Exception e) {
			// TODO: handle exception
			System.out.println("日期格式不正确");
			e.printStackTrace();
		}
		return result;
	}
	
	private String getCellFormatValue(HSSFCell cell){
		String cellValue = "";
		if( null != cell){
			switch(cell.getCellType()){
			case 0:
			case 2:
				if(HSSFDateUtil.isCellDateFormatted(cell)){
					Date date = cell.getDateCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					cellValue = sdf.format(date);
					
					if("".equals(cellValue)){
						sdf = new SimpleDateFormat("yyyy-MM-dd");
						cellValue = sdf.format(date);
					}
				}else{
					cellValue = String.valueOf(cell.getNumericCellValue());
				}
				break;
			case 1:
				cellValue = cell.getRichStringCellValue().getString();
				break;
			default:
				cellValue = " ";
				break;
			}
		}else{
			cellValue = "";
		}
		return cellValue;
	}
	public static void main(String[] args){
		try {
			InputStream is = new FileInputStream("d:\\考勤分析\\员工信息.xls");
			ExcelReaderService excelReaderService = new ExcelReaderService();
			String[] title = excelReaderService.readExcelTitle(is);
			System.out.println("获得Excel表格的标题:");
			for(String s:title){
				System.out.print(s+"    ");
			}
			System.out.println("");
			InputStream is2 = new FileInputStream("d:\\考勤分析\\员工信息.xls");
			Map map = excelReaderService.readExcelContent(is2);
			for(int i = 1 ; i<= map.size(); i++){
				System.out.println((String)map.get(Integer.valueOf(i)));
			}
		} catch (FileNotFoundException e) {
			// TODO: handle exception
			System.out.println("未找到指定路径的文件");
			e.printStackTrace();
		}
	}
}
