package com.klaus.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class EmpExcelUtil {

	private Workbook workbook;

	public EmpExcelUtil() {

		workbook = new HSSFWorkbook();

	}

	public EmpExcelUtil(String fileUri) {

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
		
		int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量

		// 遍历每个Sheet
		for (int s = 0; s < sheetCount; s++) {

			Sheet sheet = workbook.getSheetAt(s);
			
			int cellName=-1,cellStuid=-1;

			if (sheet != null) {

				int rowCount = sheet.getPhysicalNumberOfRows(); // 获取总行数

				// 遍历每一行
				for (int r = 0; r < rowCount; r++) {
					
					if(r%10==0){
						
						System.out.println("     正在处理表格中的第"+(r+1)+"行到第"+(r+10)+"行数据～");
												
					}

					Row row = sheet.getRow(r);

					if (row != null) {

						int cellCount = row.getPhysicalNumberOfCells(); // 获取总列数

						
						if(cellStuid!=-1&&cellName!=-1){
							
							Cell cell1 = row.getCell(cellStuid);
							Cell cell4 = row.getCell(cellName);
							
							if(cell1==null||cell4==null){
								
								
							}else{
								
								int cellType1 = cell1.getCellType();
								String cellValue1 = getCellValues(cellType1, cell1).replaceAll("\\s*", "");

								
								String[] arr=getStuNew(cellValue1); 
								
								if(arr!=null){

									cell1.setCellValue(arr[2]);
									cell4.setCellValue(arr[3]);
									
								}

								
							}
							
								//		System.out.println(cellValue1);
								//		System.out.println(str1);
										
										//System.out.println(cellValue4);
								//		System.out.println(str4);
							
						}else{
							
							for(int c=0;c<cellCount;c++){
								
								
								Cell cell = row.getCell(c);
								
								if(cell!=null){
									
									int cellType = cell.getCellType();
									String cellValue = getCellValues(cellType, cell).replaceAll("\\s*", "");
									
									if(cellValue.equals("学号")){
										
										cellStuid=c;
										
									}
									
									if(cellValue.equals("姓名")){
										
										cellName=c;
										
									}
									
								//	if(cellStuid!=-1&&cellName!=-1){
										
								//		break;
										
								//	}
									
								}
								
							}
							
						}


					}

				}

			}
		}
		
	

	}

	public void save(String url) throws IOException {
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
				
				DecimalFormat df = new DecimalFormat("0");  
				
				//cellValue = String.valueOf(cell.getNumericCellValue()); // 数字
				
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

	 private String[] getStuNew(String stuId){		 
		 
		 CSVUtil csv=new CSVUtil();
		 return csv.read(stuId);

		 
	 }

}
