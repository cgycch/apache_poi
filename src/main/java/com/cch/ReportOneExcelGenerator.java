package com.cch;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class ReportOneExcelGenerator extends AbstractExcelFileGenerator<ReportOne> {
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ReportOneExcelGenerator.class);
	
	List<Object[]> annotationList = new ArrayList<>();
	List<String> headerList = new ArrayList<>();
	List<Integer> headerWidthList = new ArrayList<>();
	@Override
	public void initMapper(Class<? extends ReportOne> clazz) {
		 Field[] fields = clazz.getDeclaredFields();
		 for (Field f : fields){
	            ExcelField ef = f.getAnnotation(ExcelField.class);
	            annotationList.add(new Object[]{ef,f});
	     }
		 Collections.sort(annotationList, new Comparator<Object[]>() {
				@Override
				public int compare(Object[] o1, Object[] o2) {
					return new Integer(((ExcelField)o1[0]).sort()).compareTo(
							new Integer(((ExcelField)o2[0]).sort()));
				};
			});		 
			for (Object[] os : annotationList){
				ExcelField ef = (ExcelField)os[0];
				headerList.add(ef.title());
				headerWidthList.add(ef.width());
			}
	}
	        
		
	
	@Override
	public void fillHeader(XSSFRow row, ReportOne data) throws ExcelException {
		int idx = 0;
		for (String header : headerList) {
			XSSFCell cell = row.createCell(idx++);
			cell.setCellValue(header);
		}
	}

	@Override
	public void fillDataRow(XSSFRow row, ReportOne data) throws ExcelException {
		if(headerList.size() == 0) {
			throw new ExcelException("header is not setting!");
		}
		int idx = 0;
		for (Object[] os : annotationList){
			Field field = (Field)os[1];
			try {
				System.out.println(field.get(data));
			} catch (IllegalArgumentException | IllegalAccessException e) {
				System.err.println("data get error "+field);
				LOGGER.debug("nonono");
			}
			XSSFCell cell = row.createCell(idx++);
			cell.setCellValue(data.getName());
		}
	}

	

}
