package com.cch;

import java.io.File;
import java.util.List;

public interface ExcelFileGenerator<T> {
	File generateXLSX(List<T> dataList,String filePath,String fileName,String sheetName) throws ExcelException;
	File generateXLS(List<T> data,String filePath,String fileName,String sheetName) throws ExcelException;
}
