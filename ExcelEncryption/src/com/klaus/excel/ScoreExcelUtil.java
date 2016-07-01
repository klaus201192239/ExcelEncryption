package com.klaus.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.security.MessageDigest;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ScoreExcelUtil {

	private Workbook workbook;
	

	public ScoreExcelUtil() {

		workbook = new HSSFWorkbook();

	}
	
	public ScoreExcelUtil(String fileUri) {

		try {			

			File excelFile = new File(fileUri); // 创建文件对象		
			
			FileInputStream is = new FileInputStream(excelFile); // 文件流

			
			this.workbook = WorkbookFactory.create(is); // 这种方式 Excel	
			// 2003/2007/2010
														// 都是可以处理
			
			
			
			
			
			
			
		} catch (Exception e) {

		}

	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}

	public void printExcel() {

		List<String[]> list=new ArrayList<String[]>();
		
		int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量

		// 遍历每个Sheet
		for (int s = 0; s < sheetCount; s++) {

			Sheet sheet = workbook.getSheetAt(s);

			if (sheet != null) {

				int rowCount = sheet.getPhysicalNumberOfRows(); // 获取总行数

				// 遍历每一行
				for (int r = 5; r < rowCount; r++) {
					
					if(r==5){
						
						System.out.println("     正在处理表格中的第1行到第10行数据～");
												
					}
					
					if(r%10==0){
						
						System.out.println("     正在处理表格中的第"+(r+1)+"行到第"+(r+10)+"行数据～");
												
					}

					Row row = sheet.getRow(r);

					if (row != null) {

						int cellCount = row.getPhysicalNumberOfCells(); // 获取总列数
						
						if(cellCount>=5){
							
							
							Cell cell1 = row.getCell(1);
							Cell cell4 = row.getCell(4);
							
							if(cell1==null||cell4==null){
								
							}else{
								
								String currentTime=String.valueOf(System.currentTimeMillis());
								
								int cellType1 = cell1.getCellType();
								String cellValue1 = getCellValues(cellType1, cell1).replaceAll("\\s*", "");										
								String str1 = encryptSHA(cellValue1+currentTime);
								cell1.setCellValue(str1);

								
								int cellType4 = cell4.getCellType();
								String cellValue4 = getCellValues(cellType4, cell4).replaceAll("\\s*", "");										
								String str4 = encryptSHA(cellValue4+currentTime);
								cell4.setCellValue(str4);
								
								
						//		System.out.println(cellValue1);
						//		System.out.println(str1);
								
								//System.out.println(cellValue4);
						//		System.out.println(str4);
								
								
								String[] ar={cellValue1,cellValue4,str1,str4,currentTime};
								
								list.add(ar);
								
								
							}							
							
						}

					}

				}

			}
		}
		
		
		if(list.size()>0){
			
			CSVUtil csv=new CSVUtil();
			csv.write(list);
			
		}
		

	}

	
	public void save(String url) throws IOException{
		FileOutputStream os = new FileOutputStream(url);  
		workbook.write(os);  
	    os.close();  
	}
	
	private String getCellValues(int cellType, Cell cell) {

		String cellValue = "";

		switch (cellType) {
		case Cell.CELL_TYPE_STRING: // 文本
			cellValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_NUMERIC: // 数字、日期
			if (DateUtil.isCellDateFormatted(cell)) {

				SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");

				cellValue = fmt.format(cell.getDateCellValue()); // 日期型
			} else {
				
				
				//cellValue = String.valueOf(cell.getNumericCellValue()); // 数字
				
				
				
				DecimalFormat df = new DecimalFormat("0");  
				
				cellValue= df.format(cell.getNumericCellValue());  
				
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 布尔型
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;

		case Cell.CELL_TYPE_BLANK: // 空白
			cellValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_ERROR: // 错误
			cellValue = "错误";
			break;
		case Cell.CELL_TYPE_FORMULA: // 公式
			cellValue = "错误";
			break;
		default:
			cellValue = "错误";
		}

		return cellValue;
	}
	
	
	private static final String KEY_SHA = "SHA";
	private static String encryptSHA(String coderDate) {
		
		if(coderDate==null){
			return null;
		}
		
		try{
			
			byte[] data=coderDate.getBytes();

			MessageDigest sha = MessageDigest.getInstance(KEY_SHA);
			sha.update(data);

			
			
			BigInteger shal = new BigInteger(sha.digest());
			
			
			return shal.toString(32).substring(1);
			
			
			//return sha.digest().toString();
			//
		}catch(Exception e){}
		
		return null;

	}


}
