package com.shenhangyu.check.excel;

import com.shenhangyu.check.excel.service.CheckAttendaceService;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	CheckAttendaceService service = new CheckAttendaceService();
		service.comp();
    }
}
