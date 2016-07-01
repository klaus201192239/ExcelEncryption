package com.klaus.excel;

import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.Writer;
import java.util.List;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;

public class CSVUtil {

	private String csvFilePath;
	private File file;

	public CSVUtil() {

		String s = Thread.currentThread().getContextClassLoader().getResource(".").getPath();

		String ss = s.substring(1).replace("bin", "files");

		csvFilePath = ss + "temp.csv";

		file = new File(csvFilePath);
		
		try{
			if (!file.exists()){
				
				file.createNewFile();
				
			}
		}catch(Exception e){
			
		}

	}

	public void write(String stuId, String stuName, String stuIdNew, String stuNameNew) {

		try {

			Writer writer = new FileWriter(file,true);
			
			//CSVWriter csvWriter = new 
			
			
			CSVWriter csvWriter = new CSVWriter(writer, ',');
			String[] strs = { stuId, stuName, stuIdNew, stuNameNew };
			csvWriter.writeNext(strs);
			csvWriter.close();

		} catch (Exception e) {

			System.out.println(stuName + " 信息出现错误");

		}

	}

	public void write(List<String[]> list) {

		String name = "";

		try {

			Writer writer = new FileWriter(file,true);
			
			CSVWriter csvWriter = new CSVWriter(writer, ',');

			for(int i=0;i<list.size();i++){	

				csvWriter.writeNext(list.get(i));

			}

			csvWriter.close();

		} catch (Exception e) {

			System.out.println(name + " 信息出现错误");

		}

	}

	
	public String[] read(String stuId) {

		try {

			FileReader fReader = new FileReader(file);

			CSVReader csvReader = new CSVReader(fReader);

			String[] strs = csvReader.readNext();

			while (strs != null) {

				if (strs[0].equals(stuId)) {

					csvReader.close();
					
					return strs;

				}

				strs = csvReader.readNext();

			}

			csvReader.close();

		} catch (Exception e) {

		}

		return null;
	}

	public void read() throws Exception {

		FileReader fReader = new FileReader(file);
		CSVReader csvReader = new CSVReader(fReader);

		List<String[]> list = csvReader.readAll();
		
		for (String[] ss : list) {

			if (null != ss) {

				System.out.println(ss[0] + ss[1] + ss[2] + ss[3]);

			}

		}
		csvReader.close();
	}

}
