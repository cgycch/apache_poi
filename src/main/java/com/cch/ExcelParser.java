package com.cch;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class ExcelParser<T> {
	
	private static final String EMPTY_STRING = "";
	
	private Class<T> entityClass = null;
	private Workbook workBook = null;
	private CellStyle errorCellStyle = null;
	private final List<ColumnMapper> columnMapperList = new ArrayList<ColumnMapper>();
	private final Map<Integer, String> errorMessage = new HashMap<Integer, String>();
	
	private int startRowIndex = 1;
	//private int promptColumnIndex = 1;
	
	private int lastRowNum = 0;
	private int successCount = 0;
	private int errorCount = 0;
	
	/**
	 * 构造方法
	 * @param excelFile Excel文件对象
	 * @param clazz 解析数据的封装类型
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public ExcelParser(File excelFile, Class<T> clazz, int startRowIndex, int promptColumnIndex) throws InvalidFormatException, IOException {
		this.entityClass = clazz;
		this.startRowIndex = startRowIndex;
		//this.promptColumnIndex = promptColumnIndex;
		this.createWorkbook(excelFile);
	}
	
	/**
	 * 创建poi的workbook对象
	 * @param excelFile
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private void createWorkbook(File excelFile) throws InvalidFormatException, IOException
	{
		FileInputStream fis = null;
 
		try
		{
			fis = new FileInputStream(excelFile);
			this.workBook = WorkbookFactory.create(fis);
		}
		finally
		{
			if(fis != null)
			{
				try
				{
					fis.close();
				}
				catch(IOException e)
				{
					e.printStackTrace();
				}
			}
		}
		
		if(this.workBook != null)
		{
			Sheet firstSheet = this.workBook.getSheetAt(0);
			
			if(firstSheet != null)
			{
				this.lastRowNum = firstSheet.getLastRowNum();
			}
			
			//设置错误提示列的样式、字体、颜色
			Font font = this.workBook.createFont();
			font.setColor(Font.COLOR_RED);
			this.errorCellStyle = this.workBook.createCellStyle();
			this.errorCellStyle.setFont(font);
		}
	}
	
	/**
	 * 解析Excel文件，将数据封装为List<T>并返回
	 * @param startRowIndex 解析起始行的索引，从0开始
	 * @return 解析结果集合对象
	 * @throws NoSuchFieldException
	 */
	public List<T> parseExcel() throws NoSuchFieldException
	{
		List<T> resultList = new ArrayList<T>();
		
		if(this.columnMapperList == null || this.columnMapperList.size() < 1 || this.startRowIndex < 0)
		{
			return resultList;
		}
		
		if(this.workBook != null)
		{
			Sheet firstSheet = this.workBook.getSheetAt(0);
			
			if(firstSheet != null)
			{
				Row titleRow = firstSheet.getRow(this.startRowIndex - 1);
				for(ColumnMapper mapper : this.columnMapperList)
				{
					mapper.setField(getFieldByName(this.entityClass, mapper.getFieldName()));
					
					//获取标题列名称
					if(titleRow != null && titleRow.getCell(mapper.getColumnIndex()) != null)
					{
						mapper.setColumnTitle(titleRow.getCell(mapper.getColumnIndex()).getStringCellValue().trim());
					}
				}
				
				Cell cell = null;
				Row row = null;
				Object fieldValue = null;
				T entityObject = null;
				
				for(int rowIndex = this.startRowIndex; rowIndex <= this.lastRowNum; rowIndex++)
				{
					row = firstSheet.getRow(rowIndex);
					if(row != null)
					{
						try
						{
							entityObject = this.entityClass.newInstance();
						}
						catch(Exception e1)
						{
							e1.printStackTrace();
							break;
						}
						
						boolean hasError = false;
						for(ColumnMapper mapper : this.columnMapperList)
						{
							cell = row.getCell(mapper.getColumnIndex());
							
							if(mapper.getValidator() != null)
							{
								try
								{
									mapper.getValidator().validate(cell);
								}
								catch(Exception e1)
								{
									hasError = true;
									this.errorMessage.put(rowIndex, mapper.getColumnTitle() + "[第" + (mapper.getColumnIndex()+1) + "列]：" + e1.getMessage());
									++this.errorCount;
									break;
								}
							}
							
							try
							{
								fieldValue = changeCellToObject(cell, mapper);
							}
							catch(Exception e)
							{
								hasError = true;
								this.errorMessage.put(rowIndex, mapper.getColumnTitle() + "[第" + (mapper.getColumnIndex()+1) + "列]：" + e.getMessage());
								++this.errorCount;
								break;
							}
							
							try
							{
								mapper.getField().set(entityObject, fieldValue);
							}
							catch(Exception e)
							{
								hasError = true;
								this.errorMessage.put(rowIndex, mapper.getColumnTitle() + "[第" + (mapper.getColumnIndex()+1) + "列]：" + e.getMessage());
								++this.errorCount;
								break;
							}
						}
						
						if(!hasError)
						{
							resultList.add(entityObject);
							++this.successCount;
						}
					}
				}
			}
		}
		
		return resultList;
	}
	
	/**
	 * 通过行和列的索引，获取对应单元格对象
	 * @param rowNum 行的索引，从0开始
	 * @param columnNum 列的索引，从0开始
	 * @return Cell 单元格对象
	 */
    public Cell getCellByIndex(int rowNum, int columnNum)
    {
    	if(this.workBook != null)
		{
			Sheet firstSheet = this.workBook.getSheetAt(0);
			
			if(firstSheet != null && firstSheet.getRow(rowNum) != null)
			{
				return firstSheet.getRow(rowNum).getCell(columnNum);
			}
		}
    	
    	return null;
    }
    
    /**
     * 获取错误提示文件的文件名
     * @param messageCellIndex Excel中错误提示列的索引，从1开始
     * @return 提示信息文件路径字符串
     * @throws Exception
     */
/*    public String getPromptFileName() throws Exception 
    {
    	if(this.promptColumnIndex < 1)
    	{
    		return null;
    	}
    	
    	for(Map.Entry<Integer, String> entry : this.errorMessage.entrySet())
    	{
    		addErrorMessage(entry.getKey(), this.promptColumnIndex, entry.getValue());
    	}
    	
    	File file = null;
    	OutputStream os = null;
		try
		{
			file = new File(getTemporaryFilePath(), getTemporaryFileName());
			os = new FileOutputStream(file);
			this.workBook.write(os);
		}
		catch(Exception e)
		{
			throw new Exception("创建错误提示文件出错", e);
		}
		finally
		{
			if(os != null)
			{
				try
				{
					os.close();
				}
				catch(IOException e)
				{
					e.printStackTrace();
				}
			}
		}
		
		return file.getName();
    }*/
    
    /**
     * 增加错误提示信息
     * @param rowIndex 行的索引，从0开始
     * @param cellIndex 列的索引，从0开始
     * @param message 提示信息
     */
   /* private void addErrorMessage(int rowIndex, int cellIndex, String message)
    {
    	if(this.workBook != null)
		{
			Sheet firstSheet = this.workBook.getSheetAt(0);
			
			if(firstSheet != null && firstSheet.getRow(rowIndex) != null)
			{
				Cell msgCell = firstSheet.getRow(rowIndex).createCell(cellIndex, Cell.CELL_TYPE_STRING);
				msgCell.setCellStyle(this.errorCellStyle);
				msgCell.setCellValue(message);
			}
		}
    }*/
    
    /**
     * 添加列映射对象
     * @param mapper
     */
    public void addColumnMapper(ColumnMapper mapper)
    {
    	this.columnMapperList.add(mapper);
    }
    
    /**
     * 解析成功的记录条数
     * @return 解析成功的记录条数
     */
	public int getSuccessCount()
	{
		return successCount;
	}
 
	/**
     * 解析失败的记录条数
     * @return 解析失败的记录条数
     */
	public int getErrorCount()
	{
		return errorCount;
	}
 
	/**
	 * 获取Excel文件最后一行的索引，从0开始
	 * @return 最后一行的索引，从0开始
	 */
	public int getLastRowNum()
	{
		return lastRowNum;
	}
    
	/**
	 * 将单元格转换为Java对象的工具方法
	 * @param cell 单元格对象
	 * @param mapper 列映射器
	 * @return 转换后的Java对象
	 * @throws Exception
	 */
	private static Object changeCellToObject(Cell cell, ColumnMapper mapper) throws Exception 
	{
		Object fieldValue = null;
		Class<?> fieldType = mapper.getField().getType();
		String cellValueString = changeCellToString(cell, fieldType);
		
		if(cellValueString != null && !(EMPTY_STRING.equals(cellValueString.trim())))
		{
			try
			{
				if(String.class.equals(fieldType))
				{
					fieldValue = cellValueString;
				}
				else if(Integer.class.equals(fieldType) || int.class.equals(fieldType))
				{
					fieldValue = Integer.valueOf(cellValueString);
				}
				else if(Long.class.equals(fieldType) || long.class.equals(fieldType))
				{
					fieldValue = Long.valueOf(cellValueString);
				}
				else if(Float.class.equals(fieldType) || float.class.equals(fieldType))
				{
					fieldValue = Float.valueOf(cellValueString);
				}
				else if(Double.class.equals(fieldType) || double.class.equals(fieldType))
				{
					fieldValue = Double.valueOf(cellValueString);
				}
				else if(Date.class.equals(fieldType))
				{
					switch(cell.getCellType())
					{
						case Cell.CELL_TYPE_STRING:
						{
							fieldValue = new SimpleDateFormat(mapper.getDateFormatString().trim()).parse(cellValueString);
							break;
						}
						case Cell.CELL_TYPE_NUMERIC:
						{
							if(DateUtil.isCellDateFormatted(cell))
							{
								fieldValue = cell.getDateCellValue();
							}
							break;
						}
					}
				}
			}
			catch(NumberFormatException ne)
			{
				throw new Exception("数字解析出错");
			}
			catch(ParseException pe)
			{
				throw new Exception("日期解析出错");
			}
			catch(IllegalArgumentException ae)
			{
				throw new Exception("日期格式参数无效");
			}
		}
		
		return fieldValue;
	}
	
	/**
	 * 将单元格转换为合适的字符串形式
	 * @param cell 单元格对象
	 * @param fieldType Java属性的类型
	 * @return 单元格转换后的字符串
	 */
	private static String changeCellToString(Cell cell, Class<?> fieldType)
	{
		String cellValueString = "";
		
		if(cell != null)
		{
			switch(cell.getCellType())
			{
				case Cell.CELL_TYPE_STRING:
				{
					cellValueString = cell.getStringCellValue().trim();
					break;
				}
				
				case Cell.CELL_TYPE_NUMERIC:
				{
					if(DateUtil.isCellDateFormatted(cell))
					{
						cellValueString = EMPTY_STRING;
					}
					else
					{
						if(Float.class.equals(fieldType) || float.class.equals(fieldType)
						  || Double.class.equals(fieldType) || double.class.equals(fieldType))
						{
							cellValueString = String.valueOf(cell.getNumericCellValue());
						}
						else
						{
							cell.setCellType(Cell.CELL_TYPE_STRING);
							cellValueString = cell.getStringCellValue().trim();
						}
					}
					break;
				}
				
				case Cell.CELL_TYPE_BLANK:
				{
					cellValueString = EMPTY_STRING;
					break;
				}
				
				case Cell.CELL_TYPE_FORMULA:
				{
					try
					{
						cellValueString = String.valueOf(cell.getNumericCellValue());
					}
					catch(Exception e)
					{
						cellValueString = cell.getRichStringCellValue().toString().trim();
					}
					break;
				}
				
				default:
				{
					cellValueString = EMPTY_STRING;
				}
			}
		}
		else
		{
			cellValueString = EMPTY_STRING;
		}
		
		return cellValueString;
	}
	
    private static Field getFieldByName(Class<?> clazz, String fieldName) throws NoSuchFieldException
    {
    	Field field = null;
    	try
		{
			field = clazz.getDeclaredField(fieldName);
		}
		catch(NoSuchFieldException e)
		{
			field = clazz.getSuperclass().getDeclaredField(fieldName);
		}
		
		field.setAccessible(true);
		
		return field;
    }
    
    /**
     * 获取临时文件保存的路径
     * @return 临时文件路径的File对象
     * @throws IOException
     */
   /* private static File getTemporaryFilePath() throws IOException
    {
    	File tempPath =new File((File)(ServletActionContext.getServletContext().getAttribute("javax.servlet.context.tempdir")),
    									SysConfigProperties.EXCEL_CHILD_DIR);
    	
		if(!tempPath.exists())
		{
			tempPath.mkdirs();
		}
 
		return tempPath.getCanonicalFile();
    }*/
    
    /**
     * 获取临时文件的名称
     * @return 临时文件的名称
     * @throws IOException
     */
  /*  private static String getTemporaryFileName()
	{
		StringBuilder fileName = new StringBuilder(UserContext.getContext()
															  .getCurrentUser()
															  .getUsername());
		
		fileName.append("_")
				.append(UUID.randomUUID().toString());
		
		return fileName.toString();
	}*/
    
    /**
     * Excel列与JAVA对象属性的映射器，
     * 每一列对应一个属性
     */
    public static class ColumnMapper
    {
    	private int columnIndex;
    	private String columnTitle = "";
    	private String fieldName;
    	private String dateFormatString;
    	private Field field;
    	private ColumnValidator validator;
    	
    	/**
    	 * 映射器构造方法
    	 * @param columnIndex 列在Excel文件的中的索引，从0开始
    	 * @param fieldName 列对应的Java属性名称
    	 * @param dateFormatString 日期格式（如“yyyy-MM-dd”），如需要解析为Java日期对象，则需设置该属性
    	 * @param validator 列对应的数据校验接口
    	 */
    	public ColumnMapper(int columnIndex, String fieldName, String dateFormatString, ColumnValidator validator){
			this.columnIndex = columnIndex;
			this.fieldName = fieldName;
			this.dateFormatString = dateFormatString;
			this.validator = validator;
		}
 
		public int getColumnIndex()
    	{
    		return columnIndex;
    	}
    	
    	public void setColumnIndex(int columnIndex)
		{
			this.columnIndex = columnIndex;
		}
    	
    	public String getColumnTitle()
		{
			return columnTitle;
		}
 
    	private void setColumnTitle(String columnTitle)
		{
			this.columnTitle = columnTitle;
		}
    	
    	public String getFieldName()
		{
			return fieldName;
		}
    	
    	public void setFieldName(String fieldName)
		{
			this.fieldName = fieldName;
		}
    	
    	public String getDateFormatString()
		{
			return dateFormatString;
		}
 
		public void setDateFormatString(String dateFormatString)
		{
			this.dateFormatString = dateFormatString;
		}
		
		private void setField(Field field)
		{
			this.field = field;
		}
		
		private Field getField()
    	{
    		return field;
    	}
    	
    	public ColumnValidator getValidator()
    	{
    		return validator;
    	}
    	
    	public void setValidator(ColumnValidator validator)
		{
			this.validator = validator;
		}
    }
    
    /**
     * Excel数据校验回调接口。
     * 校验失败时，将提示信息以异常形式抛出
     */
	public interface ColumnValidator
	{
		/**
		 * 数据校验回调方法。
		 * 校验失败时，将提示信息以异常形式抛出
		 * @param cell 单元格对象
		 * @throws Exception 校验提示信息
		 */
		public void validate(Cell cell) throws Exception;
	}
	
	
	/**
	 * 工具类使用示例
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception
	{
		//要解析的Excel文件对象
		File file = new File("E:/Download/DeptPriceTemplate_import.xls");
		
		//创建解析器（Object是用于封装解析数据的Java类，自行创建）
		//第三个参数为解析的起始行（从0开始），第四个参数为提示信息的列索引（从1开始）
		ExcelParser<Object> parser = new ExcelParser<Object>(file, Object.class, 1, 10);
		
		//获取Excel文档最后一行索引（如果行数超限，则返回提示，无需继续解析）
		System.out.println("最后一行索引：" + parser.getLastRowNum());
		
		//创建列校验对象
		ColumnValidator cValidator = new ColumnValidator(){
			public void validate(Cell cell) throws Exception{
				if(cell==null || "".equals(cell.toString().trim()))
				{
					throw new Exception("该字段不能为空！");
				}
			}
		};
		
		//添加列映射器（索引从0开始）
		parser.addColumnMapper(new ColumnMapper(0, "empCode", null, null));
		parser.addColumnMapper(new ColumnMapper(1, "workDate", "yyyy-MM-DD", null));
		parser.addColumnMapper(new ColumnMapper(2, "workTime", null, cValidator));
		
		//解析文件
		List<Object> list = parser.parseExcel();
		
		//获取正确条数和错误条数
		System.out.println("成功条数：" + list.size() + "失败条数：" + parser.getErrorCount());
		
		//获取错误提示信息文件名
		/*if(parser.getErrorCount() > 0)
		{
			String errorFilePath = parser.getPromptFileName();
			System.out.println(errorFilePath);
		}*/
	}


}
