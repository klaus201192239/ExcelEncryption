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

				System.out.println("���ڴ����ļ���  " + fileList[i]);

				ScoreExcelUtil ex = new ScoreExcelUtil(foldName + fileList[i]);
				ex.printExcel();
				ex.save(targetFold + fileList[i]);

			} catch (Exception e) {

				System.out.println("ִ���ļ� "+str+"  ����");
				
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

				System.out.println("���ڴ����ļ���  " + fileList[i]);

				EmpExcelUtil ex = new EmpExcelUtil(foldName + fileList[i]);
				ex.printExcel();
				ex.save(targetFold + fileList[i]);

			} catch (Exception e) {

				
				System.out.println("ִ���ļ� "+str+"  ����");
				
			}

		}

	}

}
