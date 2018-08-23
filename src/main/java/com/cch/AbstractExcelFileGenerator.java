package com.cch;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public abstract class AbstractExcelFileGenerator<T> implements ExcelFileGenerator<T> {
	
	private static final Logger LOGGER = LoggerFactory.getLogger(AbstractExcelFileGenerator.class);
	
	@SuppressWarnings("unchecked")
	@Override
	public File generateXLSX(List<T> dataList, String filePath, String fileName, String sheetName) throws ExcelException{
		if(StringUtils.isEmpty(filePath) || StringUtils.isEmpty(fileName)) {
			LOGGER.debug("filePath and fileName could not be empty!");
			throw new ExcelException("filePath and fileName could not be empty! ");
		}
		if(StringUtils.isEmpty(sheetName)) {
			sheetName = fileName;
		}
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName);
		int rowIdx = 0;
		if(dataList != null && dataList.size()>0) {
			initMapper((Class<? extends T>) dataList.get(0).getClass());
			fillHeader(sheet.createRow(rowIdx++),dataList.get(0));
		}
		for (T data : dataList) {
			XSSFRow row = sheet.createRow(rowIdx++);
			fillDataRow(row,data);
		}
		String excelFileName = filePath + File.separatorChar + fileName;
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(excelFileName);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
		} catch (Exception e) {
			LOGGER.debug(e.getMessage());
			throw new ExcelException(e.getMessage());
		}
		return new File(excelFileName);
	}
	
	@Override
	public File generateXLS(List<T> data, String filePath, String fileName, String sheetName) throws ExcelException {
		throw new ExcelException("current not support .XLS file to generate");
	}
	public abstract void initMapper(Class<? extends T> clazz);
	public abstract void fillHeader(XSSFRow row, T data) throws ExcelException;
	public abstract void fillDataRow(XSSFRow row, T data) throws ExcelException; 

}
