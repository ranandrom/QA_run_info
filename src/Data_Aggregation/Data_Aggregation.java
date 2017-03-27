package Data_Aggregation;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ch.ethz.ssh2.Connection;
import ch.ethz.ssh2.SCPClient;

import com.jcraft.jsch.ChannelExec;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;

public class Data_Aggregation {
	public static void Data_Aggregation_Main(String dir, String Path, int Cover, String PutPath, int Uploadtag, int Upload ){
		System.out.println();
		Calendar now_star = Calendar.getInstance();
		SimpleDateFormat formatter_star = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		System.out.println("Data_Aggregation程序开始时间: "+now_star.getTime());
		System.out.println("Data_Aggregation程序开始时间: "+formatter_star.format(now_star.getTime()));
		System.out.println("===============================================");
		
    	SimpleDateFormat formatter_Date = new SimpleDateFormat("yyyyMMdd");
    	String Day = formatter_Date.format(now_star.getTime());
		//String dir = "./Collect";
		String Plasma_File = dir + "/" + "Plasma_" + Day + ".xlsx";
		String Tissue_File = dir + "/" + "Tissue_" + Day + ".xlsx";
		String Unknown_File = dir + "/" + "Unknown_" + Day + ".xlsx";
		//System.out.println("dir = " + dir);
		
		my_mkdir( dir );
		
		//如果文件不存在，则创建
		if(!new File(Plasma_File).exists() && !new File(Plasma_File).isFile()){
			CreateXlsx(new File(Plasma_File));
		}
		//如果文件不存在，则创建
		if(!new File(Tissue_File).exists() && !new File(Tissue_File).isFile()){
			CreateXlsx(new File(Tissue_File));
		}
		//如果文件不存在，则创建
		if(!new File(Unknown_File).exists() && !new File(Unknown_File).isFile()){
			CreateXlsx(new File(Unknown_File));
		}
		
		ArrayList<String> Plasma_File_List = new ArrayList<String>();
		ArrayList<String> Tissue_File_List = new ArrayList<String>();
		ArrayList<String> Unknown_File_List = new ArrayList<String>();
		ArrayList<String> All_File_List = new ArrayList<String>();
		ArrayList<String> Upload_All_File_List = new ArrayList<String>();
		ArrayList<String> File_List = new ArrayList<String>();
		
		ArrayList<String> Plasma_Data_List = new ArrayList<String>();
		ArrayList<String> Tissue_Data_List = new ArrayList<String>();
		ArrayList<String> Unknown_Data_List = new ArrayList<String>();
		
		ArrayList<String> Plasma_Porject_File_List = new ArrayList<String>();//血浆项目文件列表
		ArrayList<String> Tissue_Porject_File_List = new ArrayList<String>();//组织项目文件列表
		
		SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
		String day = formatter.format(now_star.getTime());
		
		String Plasma_cmd = "find " + Path + " -type f -name QA_run_info_*_Plasma_" + day + "*.xlsx";
		String Tissue_cmd = "find " + Path + " -type f -name QA_run_info_*_Tissue_" + day + "*.xlsx";
		String Unknown_cmd = "find " + Path + " -type f -name QA_run_info_*_Unknown_" + day + "*.xlsx";
		
		Plasma_File_List = Linux_Cmd(Plasma_cmd);
		Tissue_File_List = Linux_Cmd(Tissue_cmd);
		Unknown_File_List = Linux_Cmd(Unknown_cmd);
		
		//血浆表
		for(int i = 0; i < Plasma_File_List.size(); i++){
			//System.out.println(Plasma_File_List.get(i));
			ReadExcelData (new File(Plasma_File_List.get(i)), Plasma_Data_List);
		}
		if(Cover == 0){
			//Cleared_Row(new File(Plasma_File));//清空所有数据行
			CreateXlsx(new File(Plasma_File));
		}
		WriteExcelData (new File(Plasma_File), Plasma_Data_List);
		All_File_List.add(Plasma_File);
		for(int j = 0; j < Plasma_Data_List.size(); j++){
			//System.out.println(Plasma_Data_List.get(j));
			String str_row[] = Plasma_Data_List.get(j).split("\t");
			String porject_name[] = str_row[0].split("-");
			
			String This_Plasma_File = dir + "/" + "Plasma_" + porject_name[0] + "_" + Day + ".xlsx";
			File file = new File(This_Plasma_File);
			//如果文件不存在，则创建
			if(!file.exists() && !file.isFile()){
				CreateXlsx(file);
			}
			if( !(All_File_List.contains(This_Plasma_File)) ){
				All_File_List.add(This_Plasma_File);
				Plasma_Porject_File_List.add(This_Plasma_File + "\t" + porject_name[0]);
				if(Cover == 0){
					//Cleared_Row(new File(This_Plasma_File));//清空所有数据行
					CreateXlsx(new File(This_Plasma_File));
				}
			}
			//CreateXlsx(new File(This_Plasma_File));
			Write_Row_Data (file, Plasma_Data_List.get(j));
		}
		System.out.println("血浆表已完成！");
		
		//组织表
		for(int i = 0; i < Tissue_File_List.size(); i++){
			ReadExcelData (new File(Tissue_File_List.get(i)), Tissue_Data_List);
		}
		if(Cover == 0){
			//Cleared_Row(new File(Tissue_File));//清空所有数据行
			CreateXlsx(new File(Tissue_File));
		}
		WriteExcelData (new File(Tissue_File), Tissue_Data_List);
		All_File_List.add(Tissue_File);
		for(int j = 0; j < Tissue_Data_List.size(); j++){
			String str_row[] = Tissue_Data_List.get(j).split("\t");
			String porject_name[] = str_row[0].split("-");
			
			String This_Tissue_File = dir + "/" + "Tissue_" + porject_name[0] + "_" + Day + ".xlsx";
			File file = new File(This_Tissue_File);
			//如果文件不存在，则创建
			if(!file.exists() && !file.isFile()){
				CreateXlsx(file);
			}
			if( !(All_File_List.contains(This_Tissue_File)) ){
				All_File_List.add(This_Tissue_File);
				Tissue_Porject_File_List.add(This_Tissue_File + "\t" + porject_name[0]);
				if(Cover == 0){
					//Cleared_Row(new File(This_Tissue_File));//清空所有数据行
					CreateXlsx(new File(This_Tissue_File));
				}
			}
			Write_Row_Data (file, Tissue_Data_List.get(j));
		}
		System.out.println("组织表已完成！");
		
		//其他数据表
		for(int i = 0; i < Unknown_File_List.size(); i++){
			ReadExcelData (new File(Unknown_File_List.get(i)), Unknown_Data_List);
		}
		if(Cover == 0){
			//Cleared_Row(new File(Unknown_File));//清空所有数据行
			CreateXlsx(new File(Unknown_File));
		}
		WriteExcelData (new File(Unknown_File), Unknown_Data_List);
		All_File_List.add(Unknown_File);
		System.out.println("其他数据表已完成！");
		
		//去除空行和重复行
		for(int i = 0; i < All_File_List.size(); i++){
			//System.out.println(All_File_List.get(i));
			while(RemoveNullRow(new File(All_File_List.get(i))) != 0){
				RemoveNullRow(new File(All_File_List.get(i)));//去除空行
			}
			RewriteExcelData(new File(All_File_List.get(i)));
			WriteToTsv(All_File_List.get(i));
			//SSh_Upload_File(All_File_List.get(i), PutPath);
		}
		
		//决定上传模式
		String findfile_cmd =  "find " + dir + " -type f -name *" + Day + "*.xlsx";
		Upload_All_File_List = Linux_Cmd(findfile_cmd);
		if( Uploadtag == 1 ){
			File_List = All_File_List;
		}else{
			File_List = Upload_All_File_List;
		}
		
		ArrayList<String> All_File_Path = new ArrayList<String>();
		// 生成血浆项目汇总矩阵
		for(int i = 0; i < Plasma_Porject_File_List.size(); i++){
			String Str_Plasma_Porject[] = Plasma_Porject_File_List.get(i).split("\t");
			//ArrayList<String> All_File_Path = new ArrayList<String>();
			//System.out.println("Plasma_Porject_File_List.get(i) = "+ i + "::" +Plasma_Porject_File_List.get(i));
			//String All_File_Path =  WMstat_File_Path(Str_Plasma_Porject[0]);
			//String All_File_Path =  WMstat_File_Path("/Scratch/analysis/Ironman/170301_E00495_0103_AHHWKLALXX/Ironman-LC002/04.clean_alignment/");
			//String All_File_Path = "/Scratch/analysis/Ironman/170301_E00495_0103_AHHWKLALXX/Ironman-LC002/04.clean_alignment/*sorted.deduplicated.WM.stat";
			String OutPutfile = dir + "/" + "Plasma_" + Str_Plasma_Porject[1] + "_" + Day +"_WM" + ".stat";
			All_File_Path.clear();
			//All_File_Path = WMstat_File_Path("/Scratch/analysis/Ironman/170301_E00495_0103_AHHWKLALXX/Ironman-LC002/04.clean_alignment/");
			All_File_Path =  WMstat_File_Path(Str_Plasma_Porject[0]);
			//if(All_File_Path != null ){
			//System.out.println("P0: " + All_File_Path.size());
			if(All_File_Path.size() != 0 ){
				//System.out.println("P: " + All_File_Path.size());
				//String cmd = "/home/jiacheng_chuan/Ironman/IRONMAN3/ComethylationParser/tag_paste_for_logcpm_for_zhirong.sh " + OutPutfile + " " + All_File_Path;
				//String cmd[] = {"/home/jiacheng_chuan/Ironman/IRONMAN3/ComethylationParser/tag_paste_for_logcpm_for_zhirong.sh", OutPutfile, All_File_Path};
				String cmd[] = new String[All_File_Path.size()+2];
				cmd[0] = "/home/jiacheng_chuan/Ironman/IRONMAN3/ComethylationParser/tag_paste_for_logcpm_for_zhirong.sh";
				cmd[1] = OutPutfile;
				for(int t = 0; t < All_File_Path.size(); t++){
					cmd[t+2] = All_File_Path.get(t);
					//System.out.println("cmd[t+2]: " + All_File_Path.get(t));
				}
				try{
					//Runtime.getRuntime().exec(cmd);
					Process process = Runtime.getRuntime().exec(cmd);
					BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
					String line = null;
					while( (line = input.readLine()) != null ){
						//data.add(line23);
					}
				}catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				Upload_All_File_List.add(OutPutfile);
			}
		}
		
		// 生成组织项目汇总矩阵
		for(int i = 0; i < Tissue_Porject_File_List.size(); i++){
			String Str_Tissue_Porject[] = Tissue_Porject_File_List.get(i).split("\t");		
			String OutPutfile = dir + "/" + "Tissue_" + Str_Tissue_Porject[1] + "_" + Day +"_WM" + ".stat";
			//System.out.println("Tissue_Porject_File_List.get(i) = "+ i + "::" +Tissue_Porject_File_List.get(i));
			All_File_Path.clear();
			//All_File_Path = WMstat_File_Path("/Scratch/analysis/Ironman/170301_E00495_0103_AHHWKLALXX/Ironman-LC002/04.clean_alignment/");
			All_File_Path =  WMstat_File_Path(Str_Tissue_Porject[0]);
			//System.out.println("T0: " + All_File_Path.size());
			if(All_File_Path.size() != 0 ){
				//System.out.println("T: "+All_File_Path.size());
				//System.out.println("+++++++++++++++++++++++++ ");
				String cmd[] = new String[All_File_Path.size()+2];
				cmd[0] = "/home/jiacheng_chuan/Ironman/IRONMAN3/ComethylationParser/tag_paste_for_logcpm_for_zhirong.sh";
				cmd[1] = OutPutfile;
				for(int t = 0; t < All_File_Path.size(); t++){
					cmd[t+2] = All_File_Path.get(t);
					//System.out.println("cmd[t+2]: " + All_File_Path.get(t));
				}
				try{
					//Runtime.getRuntime().exec(cmd);
					Process process = Runtime.getRuntime().exec(cmd);
					BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
					String line = null;
					while( (line = input.readLine()) != null ){
						//System.out.println("line: "+line);
					}
				}catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				Upload_All_File_List.add(OutPutfile);
			}
		}
		
		//上传所有文件到\\wdmycloud
		if( Upload == 1 ){
			for(int i = 0; i < File_List.size(); i++){
				SSh_Upload_File(File_List.get(i), PutPath);
			}
		}
		
		//CreateXlsx(new File(Plasma_File));
		
		Calendar now_end = Calendar.getInstance();
		SimpleDateFormat formatter_end = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		System.out.println();
		System.out.println("==============================================");
		System.out.println("Data_Aggregation程序结束时间: "+now_end.getTime());
		System.out.println("Data_Aggregation程序结束时间: "+formatter_end.format(now_end.getTime()));
		System.out.println();
	}
	
	/*//获取项目表中的(log2(CPM+1)列数据
	public static ArrayList<String> WMstat_File_Path(String filename)
	{
		ArrayList<String> data = new ArrayList<String>();
		String cmd23 =  "find " + filename + " -name " + "*WM.stat";
		try{
			Process process23 = Runtime.getRuntime().exec(cmd23);
			BufferedReader input23 = new BufferedReader(new InputStreamReader(process23.getInputStream()));
			String line23 = null;
			while( (line23 = input23.readLine()) != null ){
				data.add(line23);
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			//System.out.println("filename = " + filename);
			e.printStackTrace();
		}
		return data;
	}*/
	
	//获取项目表中的(log2(CPM+1)列数据
	public static ArrayList<String> WMstat_File_Path(String filename)
	{
		//String All_File_Path = null;
		ArrayList<String> All_File_Path = new ArrayList<String>();
		File file = new File(filename);
		//System.out.println("file = " + filename);
		int cell = 0;
		//读表数据
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(is);
			XSSFSheet sheet = wb.getSheetAt(0);	//获取第1个工作薄
				
			// 获取当前工作薄的每一行
			for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
				XSSFRow xssfrow = sheet.getRow(i);
				if(i == 0){// 从第一行获取log2(CPM+1)所在列数
					for (int j = xssfrow.getFirstCellNum(); j < xssfrow.getLastCellNum(); j++) {
						XSSFCell xssfcell = xssfrow.getCell(j);
						if( xssfcell != null){
							xssfcell.setCellType(Cell.CELL_TYPE_STRING);//设置单元格类型为String类型，以便读取时候以string类型，也可其它
							String cellValue = xssfcell.getStringCellValue().trim();
							if(cellValue.equals("log2(CPM+1)")){
								cell = j;
								break;
							}else{
								continue;
							}
						}
					}
				}else{
					XSSFCell xssfcell = xssfrow.getCell(cell);
					if( xssfcell != null){
						xssfcell.setCellType(Cell.CELL_TYPE_STRING);//设置单元格类型为String类型，以便读取时候以string类型，也可其它
						String cellValue = xssfcell.getStringCellValue().trim();
						if(cellValue.equals("NA")){
							continue;
						}else{
							All_File_Path.add(cellValue);
						}
					}
				}
			}
			is.close();
			wb.close();		
		} catch (IOException e) {
			// TODO Auto-generated catch block
			//System.out.println("filename = " + filename);
			e.printStackTrace();
		}
		return All_File_Path;
	}
	
	//用SSh上传文件
	public static void SSh_Upload_File(String filename, String PutPath)
	{
		String user = "zhirong_lu";
		String pass = "zhirong_lu";
		String host = "192.192.192.220";
		int port = 22;

		try {
			if( !(new File(PutPath).exists()) && !(new File(PutPath).isDirectory()) ){
				String command = "mkdir " + PutPath;
				JSch jsch = new JSch();
				Session session = jsch.getSession(user, host, port);
				Hashtable<String, String> config = new Hashtable<String, String>();
	            config.put("StrictHostKeyChecking", "no");
				//session.setConfig("StrictHostKeyChecking","no");
				session.setConfig(config);
				session.setPassword(pass);
				session.connect();
				/*ChannelExec channelExec1 = (ChannelExec)session.openChannel("exec");
				channelExec1.setCommand(command1);
				channelExec1.connect();*/
				ChannelExec channelExec = (ChannelExec)session.openChannel("exec");
				//InputStream in = channelExec.getInputStream();
				channelExec.setCommand(command);
				channelExec.connect();
			
				/*//channelExec.setInputStream(null);
	            BufferedReader input = new BufferedReader(new InputStreamReader(channelExec.getInputStream()));
				//channelExec.setErrStream(System.err);
				channelExec.connect();
				//接收远程服务器执行命令的结果
	            String line;
	            while ((line = input.readLine()) != null) {  
	                System.out.println(line); 
	            }  
	            input.close(); */
				//String out = IOUtils.toString(in, "UTF-8");
				channelExec.disconnect();
				session.disconnect();
				//Thread.sleep(500);
			}
			Thread.sleep(100);
			
			Connection con = new Connection(host);
			con.connect();
			boolean isAuthed = con.authenticateWithPassword(user, pass); 
			//System.out.println("文件已上传："+filename);
			SCPClient scpClient = con.createSCPClient();
			scpClient.put(filename, PutPath);//从本地复制文件到远程目录
			//scpClient.get("/wdmycloud/anchordx_cloud/杨莹莹/项目-生信-汇总表/项目-生信-进展更新-模板.xlsx","./66");//从远程获取文件
			con.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	//写成tsv格式文本
	public static void WriteToTsv(String inputfile){
		ArrayList<String> Data_list = new ArrayList<String>();
		String data = null;
		String Suffix = inputfile.substring(inputfile.lastIndexOf(".")); //获取后缀名
		String Remove_suffix =  inputfile.replaceAll(Suffix, ""); //去除后缀名
		String outputfile = Remove_suffix + ".tsv";
		File file  = new File(inputfile);
		//读表数据
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(is);
			XSSFSheet sheet = wb.getSheetAt(0);	//获取第1个工作薄
				
			// 获取当前工作薄的每一行
			for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
				//String TemplateArr[] = TemplateData.split("\t");
				XSSFRow xssfrow = sheet.getRow(i);
				int log = 0;
				// 获取当前工作薄的每一列
				for (int j = xssfrow.getFirstCellNum(); j < xssfrow.getLastCellNum(); j++) {
					XSSFCell xssfcell = xssfrow.getCell(j);
					if( xssfcell != null){
						xssfcell.setCellType(Cell.CELL_TYPE_STRING);//设置单元格类型为String类型，以便读取时候以string类型，也可其它
						String cellValue = xssfcell.getStringCellValue().trim();
						if(log == 0){
							data = cellValue;
						}else{
							data += "\t" + cellValue;
						}
						log++;
					}
				}
				Data_list.add(data);
			}
				is.close();
				wb.close();		
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
			
		//写到输出文件
		try{
			FileWriter fw = new FileWriter(outputfile);
			BufferedWriter bw = new BufferedWriter(fw);
			//bw.write("#Head"+"\r\n");// 往文件上写头信息
			for(int i = 0; i < Data_list.size(); i++){
				bw.write(Data_list.get(i)+"\r\n");
			}
			bw.close();
			fw.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//删除空行
	public static int RemoveNullRow (File file){
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = workbook.getSheetAt(0);	//获取第1个工作薄

			// 获取当前工作薄的每一行
			for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
				XSSFRow xssfrow = sheet.getRow(i);				
				if ( xssfrow == null || (CheckRowNull(xssfrow) == 0) ) {
					System.out.println("删除空行：" + i);
					//continue;
						
					sheet.shiftRows(i+1, sheet.getLastRowNum(), -1);
						
					// 新建一输出文件流
					FileOutputStream fOut = new FileOutputStream(file);
					// 把相应的Excel 工作簿存盘
					workbook.write(fOut);
					fOut.flush();
					// 操作结束，关闭文件
					fOut.close();
					is.close();
					workbook.close();
						
					return 1;
				}else{
					continue;
				}
			}
			is.close();
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return 0;
	}
	
	//判断行为空,如果为空，则返回0
	public static int CheckRowNull(XSSFRow xssfRow){
		int num = 0;
		// 获取当前工作薄的每一列
		for (int j = xssfRow.getFirstCellNum(); j < xssfRow.getLastCellNum(); j++) {
			XSSFCell xssfcell = xssfRow.getCell(j);
			//String cellValue = String.valueOf(xssfcell);
			if( xssfcell == null  || (("") == String.valueOf(xssfcell)) || xssfcell.toString().equals("") || xssfcell.getCellType() == HSSFCell.CELL_TYPE_BLANK ){
				continue;
			}else{
				num++;
			}
		}
		return num;
	}
	
	//写回数据
	public static void RewriteExcelData (File file){
		ArrayList<String> Data_list = new ArrayList<String>();
		ReadExcelData(file, Data_list);
		CreateXlsx(file);//新建同名文件覆盖原文件，达到清空数据效果
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = workbook.getSheetAt(0);	//获取第1个工作薄
			//清空所有数据行
			/*for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
				XSSFRow xssfrow = sheet.getRow(i);
				sheet.removeRow(xssfrow);
				
				FileOutputStream fOut = new FileOutputStream(file);
				// 把相应的Excel 工作簿存盘
				workbook.write(fOut);
				fOut.flush();
				// 操作结束，关闭文件
				fOut.close();
			}*/
			//写回数据
			for(int j = 0; j < Data_list.size(); j++){
				XSSFRow row = sheet.createRow((short) j+1);
				String str_row[] = Data_list.get(j).split("\t");
				for(int i = 0; i < str_row.length; i++ ){
					// 在索引0的位置创建单元格（左上端）
					XSSFCell cell = row.createCell(i);
					if(str_row[i].equals("null")){
						cell.setCellValue("");
					}else{
						cell.setCellValue(str_row[i]);
					}
				}
			}
			// 新建一输出文件流
			FileOutputStream fOut = new FileOutputStream(file);
			// 把相应的Excel 工作簿存盘
			workbook.write(fOut);
			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
			//System.out.println("文件生成...");
			is.close();
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//写一行数据
	public static void Write_Row_Data (File file, String Data){
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = workbook.getSheetAt(0);	//获取第1个工作薄
			
			//写回数据
			XSSFRow row = sheet.createRow((short) sheet.getLastRowNum()+1);
			String str_row[] = Data.split("\t");
			for(int i = 0; i < str_row.length; i++ ){
				// 在索引0的位置创建单元格（左上端）
				XSSFCell cell = row.createCell(i);
				if(str_row[i].equals("null")){
					cell.setCellValue("");
				}else{
					cell.setCellValue(str_row[i]);
				}
			}

			// 新建一输出文件流
			FileOutputStream fOut = new FileOutputStream(file);
			// 把相应的Excel 工作簿存盘
			workbook.write(fOut);
			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
			//System.out.println("文件生成...");
			is.close();
			workbook.close();
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//清空所有行数据
	public static void  Cleared_Row (File file){
		while(RemoveNullRow(file) == 1){
			RemoveNullRow(file);//去除空行
		}
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = workbook.getSheetAt(0);	//获取第1个工作薄
		
			//清空所有数据行
			for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
				XSSFRow xssfrow = sheet.getRow(i);
				sheet.removeRow(xssfrow);
				
				FileOutputStream fOut = new FileOutputStream(file);
				// 把相应的Excel 工作簿存盘
				workbook.write(fOut);
				fOut.flush();
				// 操作结束，关闭文件
				fOut.close();
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//写数据
	public static void WriteExcelData (File file, ArrayList<String> Data_list){
		//ArrayList<String> Data_list = ReadExcelData(file);
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = workbook.getSheetAt(0);	//获取第1个工作薄
			
			//Cleared_Row(file);//清空所有数据行
			//清空所有数据行
			/*for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
				XSSFRow xssfrow = sheet.getRow(i);
				sheet.removeRow(xssfrow);
				
				FileOutputStream fOut = new FileOutputStream(file);
				// 把相应的Excel 工作簿存盘
				workbook.write(fOut);
				fOut.flush();
				// 操作结束，关闭文件
				fOut.close();
			}*/
			//写回数据
			for(int j = 0; j < Data_list.size(); j++){
				XSSFRow row = sheet.createRow((short) sheet.getLastRowNum()+1);
				String str_row[] = Data_list.get(j).split("\t");
				for(int i = 0; i < str_row.length; i++ ){
					// 在索引0的位置创建单元格（左上端）
					XSSFCell cell = row.createCell(i);
					if(str_row[i].equals("null")){
						cell.setCellValue("");
					}else{
						cell.setCellValue(str_row[i]);
					}
				}
			}
			// 新建一输出文件流
			FileOutputStream fOut = new FileOutputStream(file);
			// 把相应的Excel 工作簿存盘
			workbook.write(fOut);
			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
			//System.out.println("文件生成...");
			is.close();
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//读表数据到列表，去除重复行
	public static void ReadExcelData (File file, ArrayList<String> Data_list){
		//ArrayList<String> Data_list = new ArrayList<String>();
		String TemplateData  = null;
		String data = null;
		try{
			FileInputStream is = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(is);
			XSSFSheet sheet = wb.getSheetAt(0);	//获取第1个工作薄
			
			XSSFRow xssfrow0 = sheet.getRow(0);
			for (int j = xssfrow0.getFirstCellNum(); j < xssfrow0.getLastCellNum(); j++) {
				if(j == xssfrow0.getFirstCellNum()){
					TemplateData = "null";
				}else{
					TemplateData += "\t" + "null";
				}
			}
			// 获取当前工作薄的每一行
			for (int i = sheet.getFirstRowNum()+1; i <= sheet.getLastRowNum(); i++) {
				String TemplateArr[] = TemplateData.split("\t");
				XSSFRow xssfrow = sheet.getRow(i);	
				
				// 获取当前工作薄的每一列
				for (int j = xssfrow.getFirstCellNum(); j < xssfrow.getLastCellNum(); j++) {
					XSSFCell xssfcell = xssfrow.getCell(j);
					if( xssfcell == null  || (("") == String.valueOf(xssfcell)) || xssfcell.toString().equals("") || xssfcell.getCellType() == HSSFCell.CELL_TYPE_BLANK ){
						continue;
					}else{
						String cellValue = String.valueOf(xssfcell);
						TemplateArr[j] = cellValue;
					}
				}
				for( int x = 0; x < TemplateArr.length; x++){
					if(x == 0){
						data = TemplateArr[x];
					}else{
						data += "\t" + TemplateArr[x];
					}
				}
				if( Data_list.contains(data) ){
					continue;
				}else{
					Data_list.add(data);
				}
			}
				is.close();
				wb.close();
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//return Data_list;
	}
	
	//创建目录
	public static void my_mkdir( String dir_name){
		File file = new File( dir_name );
		
		//如果文件不存在，则创建
		if(!file.exists() && !file.isDirectory()){
			//System.out.println("//目录不存在");
			file.mkdirs();
		}
	}
	
	//调用linux命令
	public static ArrayList<String> Linux_Cmd(String cmd){
		ArrayList<String> Data_list = new ArrayList<String>();
		String line = null;
		//System.out.println("linux cmd");
		try{
			//String cmd1 = "awk 'NR==8 {print $7 }' /Src_Data1/analysis/Ironman/161231_E00490_0087_AHF75NALXX/Ironman-CR003/06.clean_summary/S11-DPM-CR003-097-PS1-IRM_S11_R1_001.clean.fastq.gz_bismark_bt2_pe.sorted.bam.hsmetrics.txt";
			Process process = Runtime.getRuntime().exec(cmd);
			BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
			//line = input.readLine();
			//System.out.println("linux cmd1");
			while ((line = input.readLine()) != null) {
				//System.out.println(line);
				Data_list.add(line);
			}
		}catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return Data_list;
	}
	
	//新建xlsx格式文件
	public static void CreateXlsx(File file) {
		//File file = new File(New_File);
		try{
		//FileInputStream is = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook();
		// 创建Excel的工作sheet,对应到一个excel文档的tab  
		XSSFSheet sheet = workbook.createSheet("sheet1");
		
		// 在索引0的位置创建行（最顶端的行）
		XSSFRow row0 = sheet.createRow((short) 0);
		// 在索引0的位置创建单元格（左上端）
		//XSSFCell cell = row.createCell((short) 0);
		
		String head_row0 = "Sample ID"+"\t"+
		"Pre-lib name"+"\t"+
				"Identification name"+"\t"+
		"Sequencing info"+"\t"+
				"Sequencing file name"+"\t"+
		"Mapping%"+"\t"+
				"Total PF reads"+"\t"+
		"Mean_insert_size"+"\t"+
				"Median_insert_size"+"\t"+
		"On target%"+"\t"+
				"Pre-dedup mean bait coverage"+"\t"+
		"Deduped mean bait coverage"+"\t"+
				"Deduped mean target coverage"+"\t"+
		"% target bases > 30X"+"\t"+
				"Uniformity (0.2X mean)"+"\t"+
		"C methylated in CHG context"+"\t"+
				"C methylated in CHH context"+"\t"+
		"C methylated in CpG context"+"\t"+
				"QC result"+"\t"+
		"Date of QC"+"\t"+
				"Path to sorted.deduped.bam"+"\t"+
		"Date of path update"+"\t"+
				"Bait set"+"\t"+
		"log2(CPM+1)"+"\t"+
				"Sample QC"+"\t"+
		"Failed QC Detail"+"\t"+
				"Warning QC Detail"+"\t"+
		"Check"+"\t"+
				"Note1"+"\t"+
		"Note2"+"\t"+
				"Note3";
		
		String head_row1 = "样本编号"+"\t"
				+"预文库样本名"+"\t"
				+ "上机"+"\t"
				+"e.g. with S01 prefix before a Pre-lib name - this is to separate multiple sequencing files derived from the same pre-library";
		
		//1、创建字体，设置其为红色：
		 XSSFFont font = workbook.createFont();
		 font.setColor(HSSFFont.COLOR_RED);
		 font.setFontHeightInPoints((short)10);
		 font.setFontName("Palatino Linotype");
		 //font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		 //2、创建格式
		 XSSFCellStyle cellStyle= workbook.createCellStyle();
		 cellStyle.setFont(font);
		 cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		 cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		 cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		 cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		 cellStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
		 cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		 
		//1、创建字体，设置其为粗体，背景蓝色：
		 XSSFFont font1 = workbook.createFont();
		 //font1.setColor(HSSFFont.COLOR_RED);
		 font1.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		 font1.setFontHeightInPoints((short)10);
		 font1.setFontName("Palatino Linotype");
		 //2、创建格式
		 XSSFCellStyle cellStyle1= workbook.createCellStyle();
		 cellStyle1.setFont(font1);
		 cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
		 cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		 cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		 cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
		 cellStyle1.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
		 cellStyle1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		 
		//1、创建字体，设置其为红色、粗体，背景绿色：
		 XSSFFont font2 = workbook.createFont();
		 font2.setColor(HSSFFont.COLOR_RED);
		 font2.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		 font2.setFontHeightInPoints((short)10);
		 font2.setFontName("Palatino Linotype");
		 //2、创建格式
		 XSSFCellStyle cellStyle2= workbook.createCellStyle();
		 cellStyle2.setFont(font2);
		 cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
		 cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		 cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		 cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
		 cellStyle2.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
		 cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		 
		//1、创建字体大小为10，背景蓝色：
		 XSSFFont font3 = workbook.createFont();
		 font3.setFontHeightInPoints((short)10);
		 font3.setFontName("Palatino Linotype");
		 //2、创建格式
		 XSSFCellStyle cellStyle3= workbook.createCellStyle();
		 cellStyle3.setFont(font3);
		 cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
		 cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		 cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		 cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);
		 cellStyle3.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
		 cellStyle3.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		 
		//1、创建字体大小为10，背景黄色：
		 XSSFFont font4 = workbook.createFont();
		 font4.setFontHeightInPoints((short)10);
		 font4.setFontName("Palatino Linotype");
		 //2、创建格式
		 XSSFCellStyle cellStyle4= workbook.createCellStyle();
		 cellStyle4.setFont(font4);
		 cellStyle4.setBorderTop(XSSFCellStyle.BORDER_THIN);
		 cellStyle4.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		 cellStyle4.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		 cellStyle4.setBorderRight(XSSFCellStyle.BORDER_THIN);
		 cellStyle4.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
		 cellStyle4.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		 
		//1、创建字体，设置其为粗体，背景黄色：
		 XSSFFont font5 = workbook.createFont();
		 //font1.setColor(HSSFFont.COLOR_RED);
		 font5.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		 font5.setFontHeightInPoints((short)10);
		 font5.setFontName("Palatino Linotype");
		 //2、创建格式
		 XSSFCellStyle cellStyle5= workbook.createCellStyle();
		 cellStyle5.setFont(font5);
		 cellStyle5.setBorderTop(XSSFCellStyle.BORDER_THIN);
		 cellStyle5.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		 cellStyle5.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		 cellStyle5.setBorderRight(XSSFCellStyle.BORDER_THIN);
		 cellStyle5.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
		 cellStyle5.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		 
		String str_head_row0[] = head_row0.split("\t");
		// 在单元格中输入一些内容
		for(int i = 0; i < str_head_row0.length; i++ ){
			// 在索引0的位置创建单元格（左上端）
			XSSFCell cell = row0.createCell(i);
			if( i < 3 ){
				cell.setCellValue(str_head_row0[i]);
				cell.setCellStyle(cellStyle2);
			}else if( i == str_head_row0.length-11 || i == str_head_row0.length-10 ){
				cell.setCellStyle(cellStyle5);
				cell.setCellValue(str_head_row0[i]);
			}else{
				cell.setCellStyle(cellStyle1);
				cell.setCellValue(str_head_row0[i]);
			}
		}
		/*XSSFRow row1 = sheet.createRow((short) 1);
		String str_head_row1[] = head_row1.split("\t");
		for(int i = 0; i < str_head_row0.length; i++ ){
			// 在索引0的位置创建单元格（左上端）
			XSSFCell cell = row1.createCell(i);
			if( i < str_head_row1.length ){
				if (i < 3 ){
					cell.setCellStyle(cellStyle);
					cell.setCellValue(str_head_row1[i]);
				}else{
					cell.setCellStyle(cellStyle3);
					cell.setCellValue(str_head_row1[i]);
				}
			}else if( i == str_head_row0.length-3 || i == str_head_row0.length-2 ){
				cell.setCellStyle(cellStyle4);
			}else{
				cell.setCellStyle(cellStyle3);
			}
		}*/
		// 新建一输出文件流
		FileOutputStream fOut = new FileOutputStream(file);
		// 把相应的Excel 工作簿存盘
		workbook.write(fOut);
		fOut.flush();
		// 操作结束，关闭文件
		fOut.close();
		//System.out.println("文件生成...");
		//is.close();
		workbook.close();
		}catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
