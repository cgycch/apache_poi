package com.cch;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class App {
	public static void main(String[] args) {
		ExcelFileGenerator<ReportOne> generator = new ReportOneExcelGenerator();
		List<ReportOne> dataList = new ArrayList<>();
		dataList.add(new ReportOne("qwe", "dsfaf", "fdgg"));
		String filePath = "E:\\youdaoyunfile";
		String fileName = "myReport.xlsx";
		String sheetName ="sheetone";
		File file = generator.generateXLSX(dataList, filePath, fileName, sheetName);
		System.out.println(file==null);
	}
}
