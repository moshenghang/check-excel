/**
 *版权所有©微信公众号|视频号:深航渔
 */
package com.shenhangyu.check.excel.utils;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class DateUtils {
	
	/*
	 * 格式化Date数据为yyyy-MM-dd的String
	 */
	public static String dateToString(Date date){
		if( null == date){
			return null;
		}
		SimpleDateFormat simpleDateformat = new SimpleDateFormat("yyyy-MM-dd");
		return simpleDateformat.format(date);
	}
	
	public static Date getEndDateByMonths(Date startDate,int months){
		Calendar calendarStartDate = Calendar.getInstance();
		calendarStartDate.setTime(startDate);
		calendarStartDate.add(2, months);
		return calendarStartDate.getTime();
	}
	
	public static Date getFirstDate(Date today){
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(today);
		calendar.set(5, 1);
		return calendar.getTime();
	}
	
	public static String getEndDateByDays(String startDate,int days){
		String endDate = null;
		try {
			SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
			Date startDt = formatter.parse(startDate);
			Date endDt = getEndDateByDays(startDt,days);
			endDate	= formatter.format(endDt);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		return endDate;
	}
	
	public static Date getEndDateByDays(Date startDate,int days){
		Calendar calendarStartDate = Calendar.getInstance();
		calendarStartDate.setTime(startDate);
		calendarStartDate.add(6, days);
		return calendarStartDate.getTime();
	}
	
}
