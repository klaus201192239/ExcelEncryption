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

			File excelFile = new File(fileUri); // �����ļ�����

			FileInputStream is = new FileInputStream(excelFile); // �ļ���

			this.workbook = WorkbookFactory.create(is); // ���ַ�ʽ Excel
			// 2003/2007/2010
			// ���ǿ��Դ���

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
		
		int sheetCount = workbook.getNumberOfSheets(); // Sheet������

		// ����ÿ��Sheet
		for (int s = 0; s < sheetCount; s++) {

			Sheet sheet = workbook.getSheetAt(s);
			
			int cellName=-1,cellStuid=-1;

			if (sheet != null) {

				int rowCount = sheet.getPhysicalNumberOfRows(); // ��ȡ������

				// ����ÿһ��
				for (int r = 0; r < rowCount; r++) {
					
					if(r%10==0){
						
						System.out.println("     ���ڴ������еĵ�"+(r+1)+"�е���"+(r+10)+"�����ݡ�");
												
					}

					Row row = sheet.getRow(r);

					if (row != null) {

						int cellCount = row.getPhysicalNumberOfCells(); // ��ȡ������

						
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
									
									if(cellValue.equals("ѧ��")){
										
										cellStuid=c;
										
									}
									
									if(cellValue.equals("����")){
										
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
		case Cell.CELL_TYPE_STRING: // �ı�
			cellValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_NUMERIC: // ���֡�����
			if (DateUtil.isCellDateFormatted(cell)) {

				SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");

				cellValue = fmt.format(cell.getDateCellValue()); // ������
			} else {
				
				DecimalFormat df = new DecimalFormat("0");  
				
				//cellValue = String.valueOf(cell.getNumericCellValue()); // ����
				
				cellValue= df.format(cell.getNumericCellValue());  
				
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN: // ������
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;

		case Cell.CELL_TYPE_BLANK: // �հ�
			cellValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_ERROR: // ����
			cellValue = "����";
			break;
		case Cell.CELL_TYPE_FORMULA: // ��ʽ
			cellValue = "����";
			break;
		default:
			cellValue = "����";
		}

		return cellValue;
	}

	 private String[] getStuNew(String stuId){		 
		 
		 CSVUtil csv=new CSVUtil();
		 return csv.read(stuId);

		 
	 }

}
