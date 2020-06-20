/**
 *版权所有©微信公众号|视频号:深航渔
 */
package com.shenhangyu.check.excel.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.shenhangyu.check.excel.utils.DateUtils;

public class CheckAttendaceService {
	HashMap<String, String> namemap = new HashMap<String, String>();
	HashMap<String, String> departmap = new HashMap<String, String>();
	HashMap<String, String> startmap = new HashMap<String, String>();
	HashMap<String, String> endmap = new HashMap<String, String>();
	HashMap<String, String> neimap = new HashMap<String, String>();
	HashMap<String, String> weimap = new HashMap<String, String>();
	ExcelReaderService service = new ExcelReaderService();
	int total = 0;

	public boolean getPesList() {
		boolean b = true;
		try {
			String xlsname = "D:\\考勤分析\\员工信息.xls";
			File sfile = new File(xlsname);
			if (sfile.exists()) {
				System.out.println("获取到：" + xlsname + "文件，正在读取中..........");
				InputStream sfileis = new FileInputStream(sfile);
				List<HashMap<Integer, String>> contlist = null;

				sfileis = new FileInputStream(sfile);
				contlist = this.service.readExcelContent2(sfileis);
				int i = 0;
				for (HashMap map : contlist) {
					String job_no = (String) map.get(Integer.valueOf(0));
					String job_name = (String) map.get(Integer.valueOf(1));
					String job_dept = (String) map.get(Integer.valueOf(2));
					String start_time = (String) map.get(Integer.valueOf(3));
					String end_time = (String) map.get(Integer.valueOf(4));
					this.namemap.put(job_no, job_name);
					this.departmap.put(job_no, job_dept);
					this.startmap.put(job_no, start_time);
					this.endmap.put(job_no, end_time);
					i++;
				}
				System.out.println("员工信息读取完成，共" + i + "人");
				this.total = i;
			} else {
				b = false;
				System.out.println("未获取到员工信息表，庆检查文件..........");
			}
			String xlsname1 = "D:\\考勤分析\\内部考勤数据.xls";
			File sfile1 = new File(xlsname1);
			if (sfile1.exists()) {
				System.out.println("已获取到 内部考勤数据文件");
			} else {
				b = false;
				System.out.println("未获取到 内部考勤数据文件，请检查目录是否存在" + sfile1.getName() + "文件");
			}
			String xlsname2 = "D:\\考勤分析\\外部考勤数据.xls";
			File sfile2 = new File(xlsname2);
			if (sfile2.exists()) {
				System.out.println("已获取到 外部考勤数据文件");
			} else {
				b = false;
				System.out.println("未获取到 外部考勤数据文件，请检查目录是否存在" + sfile2.getName() + "文件");
			}
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		return b;
	}

	public void comp() {
		boolean b = getPesList();
		if (b) {
			try {
				String xlsname = "D:\\考勤分析\\内部考勤数据.xls";
				File sfile = new File(xlsname);
				InputStream sfileis = new FileInputStream(sfile);
				List<HashMap<Integer, String>> nei = this.service.readExcelContent2(sfileis);

				String xlsname2 = "D:\\考勤分析\\外部考勤数据.xls";
				File sfile2 = new File(xlsname2);
				InputStream sfileis2 = new FileInputStream(sfile2);
				List<HashMap<Integer, String>> wai = this.service.readExcelContent2(sfileis2);

				System.out.println("今天是" + DateUtils.dateToString(new Date()) + "，上个月为"
						+ DateUtils.dateToString(DateUtils.getEndDateByMonths(new Date(), -1)).substring(0, 7));

				Date lm = DateUtils.getFirstDate(DateUtils.getEndDateByMonths(new Date(), -1));

				String jgxlsname = "D:\\考勤分析\\考勤结果.xls";
				Workbook work = new HSSFWorkbook();
				Sheet sheet = work.createSheet("sheet1");
				CellStyle style = work.createCellStyle();
				style.setFillForegroundColor(HSSFColor.RED.index);
				style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				Row row = null;
				Cell cell = null;
				String[] exceltitle = { "员工工号         ", "姓名     ", "部门       ", "考勤日期         ", "考勤时间          ",
						"考勤核对结果          " };
				row = sheet.createRow(0);
				for (int xx = 0; xx < exceltitle.length; xx++) {
					cell = row.createCell(xx);
					cell.setCellValue(exceltitle[xx]);
				}
				int i = 0;
				int xx = 1;
				for (Map.Entry ent : this.namemap.entrySet()) {
					String get_job_no = (String) ent.getKey();
					String get_job_dept = this.departmap.get(get_job_no) == null ? ""
							: (String) this.departmap.get(get_job_no);
					String get_job_name = (String) ent.getValue();
					String get_start_time = this.startmap.get(get_job_no) == null ? ""
							: (String) this.startmap.get(get_job_no);
					String get_end_time = this.endmap.get(get_job_no) == null ? ""
							: (String) this.endmap.get(get_job_no);
					int start = Integer.parseInt(get_start_time.substring(0, get_start_time.indexOf(".")));
					int end = Integer.parseInt(get_end_time.substring(0, get_end_time.indexOf(".")));
					for (int x = 0; x <= 32; x++) {
						Date pdate = DateUtils.getEndDateByDays(lm, x);
						int day = pdate.getDay();
						boolean zm = false;
						if ((day == 0) || (day == 6)) {
							zm = true;
						}
						String pdatestr = DateUtils.dateToString(pdate);
						if (DateUtils.dateToString(pdate)
								.equals(DateUtils.dateToString(DateUtils.getFirstDate(new Date()))))
							break;
						String check_job_no = "";
						String i_xls_device_no = "";
						String o_xls_device_no = "";
						String check_date = "";
						String check_time = "";
						String i_xls_check_time;
						for (HashMap map : nei) {
							String i_xls_job_no = (String) map.get(Integer.valueOf(0));
							String i_xls_device_no_str = (String) map.get(Integer.valueOf(3));
							String i_xls_check_date = map.get(Integer.valueOf(1)) == null ? ""
									: (String) map.get(Integer.valueOf(1));
							i_xls_check_time = map.get(Integer.valueOf(2)) == null ? ""
									: (String) map.get(Integer.valueOf(2));
							if ((i_xls_check_date.split(" ")[0].equals(pdatestr))
									&& (i_xls_job_no.equals(i_xls_job_no.equals(get_job_no)))) {
								check_job_no = get_job_no;
								i_xls_device_no = i_xls_device_no_str;
								check_date = i_xls_check_date;
								if (i_xls_check_time.split(" ").length == 2) {
									i_xls_check_time = i_xls_check_time.split(" ")[1];
								}
								check_time = check_time + i_xls_check_time + " ";
							}
						}
						String o_xls_check_date;
						for (HashMap map : wai) {
							String o_xls_check_job_no = (String) map.get(Integer.valueOf(0));
							String date = map.get(Integer.valueOf(1)) == null ? ""
									: (String) map.get(Integer.valueOf(1));
							String time = map.get(Integer.valueOf(2)) == null ? ""
									: (String) map.get(Integer.valueOf(2));
							o_xls_check_date = date + " " + time.substring(0,time.lastIndexOf(":"));
							if ((o_xls_check_date.split(" ")[0].equals(pdatestr))
									&& (o_xls_check_job_no.equals(get_job_no))) {
								check_job_no = get_job_no;
								check_date = o_xls_check_date;
								check_time = check_time + o_xls_check_date.split(" ")[1] + " ";
							}
						}
						row = sheet.createRow(xx);
						cell = row.createCell(0);
						cell.setCellValue(get_job_no);

						cell = row.createCell(1);
						cell.setCellValue(get_job_name);

						cell = row.createCell(2);
						cell.setCellValue(get_job_dept);

						cell = row.createCell(3);
						cell.setCellValue(pdatestr);

						cell = row.createCell(4);
						cell.setCellValue(check_time);
						cell = row.createCell(5);
						String type = "";
						if (zm) {
							if ((check_time.split(" ").length == 1) && (check_time.contains(":"))) {
								type = "周末单边";
								cell.setCellStyle(style);
							}
						}
						if (!zm) {
							if ((check_time.split(" ").length == 1) && (check_time.contains(":"))) {
								type = "工作日单边";
								cell.setCellStyle(style);
							}
							if (check_time.equals("")) {
								type = "工作日无考勤";
								cell.setCellStyle(style);
							}

							if ((!check_time.equals("")) && (check_time.split(":").length >= 3)) {
								int[] ks = new int[check_time.split(" ").length];
								int ksi = 0;
								String[] arrayOfString1;
								int count = (arrayOfString1 = check_time.split(" ")).length;
								for (int d = 0; d < count; d++) {
									String str = arrayOfString1[d];
									int t = Integer.parseInt(str.replaceAll(":", ""));
									ks[ksi] = t;
									ksi++;
								}
								Arrays.sort(ks);
								int c = ks[0];
								int t = ks[(ks.length - 1)];
								if (c > start) {
									type = "工作日迟到";
									cell.setCellStyle(style);
								}
								if (t < end) {
									type = "工作日早退";
									cell.setCellStyle(style);
								}
							}
						}
						cell.setCellValue(type);
						xx++;
					}
					i++;
					if (i % 10 == 0) {
						System.out.println("已核对 " + i + " 人");
					}
					if (i == this.total) {
						System.out.println("核对完毕 ，" + i);
					}
				}
				for (int y = 0; y <= 10; y++) {
					sheet.autoSizeColumn(y);
					sheet.autoSizeColumn(y, true);
				}
				System.out.println("正在准备写入数据>>>>>>>>");
				FileOutputStream os = new FileOutputStream(jgxlsname);
				work.write(os);
				os.flush();
				os.close();
				System.out.println("已生成，核对结果  D:\\考勤分析\\考勤结果.xls");

			} catch (Exception e) {
				// TODO: handle exception
			}
		}
	}
}
