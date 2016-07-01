package com.klaus.excel;

import java.io.File;

public class FileUtil {

	public void startScoreFileEncrypt(String foldName, String targetFold) {

		File file = new File(foldName);

		String[] fileList = file.list();

		int length = fileList.length;

		for (int i = 0; i < length; i++) {

			String str=fileList[i];
			
			try {

				System.out.println("正在处理文件：  " + fileList[i]);

				ScoreExcelUtil ex = new ScoreExcelUtil(foldName + fileList[i]);
				ex.printExcel();
				ex.save(targetFold + fileList[i]);

			} catch (Exception e) {

				System.out.println("执行文件 "+str+"  出错");
				
			}

		}

	}

	public void startEmpFileEncrypt(String foldName, String targetFold) {

		File file = new File(foldName);

		String[] fileList = file.list();

		int length = fileList.length;

		for (int i = 0; i < length; i++) {

			
			String str=fileList[i];
			
			try {

				System.out.println("正在处理文件：  " + fileList[i]);

				EmpExcelUtil ex = new EmpExcelUtil(foldName + fileList[i]);
				ex.printExcel();
				ex.save(targetFold + fileList[i]);

			} catch (Exception e) {

				
				System.out.println("执行文件 "+str+"  出错");
				
			}

		}

	}

}
