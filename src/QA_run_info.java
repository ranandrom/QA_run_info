import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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

import Data_Aggregation.Data_Aggregation;

public class QA_run_info {

	public static void main(String[] args) throws InterruptedException
    {
		System.out.println();
		Calendar now_star = Calendar.getInstance();
		SimpleDateFormat formatter_star = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		System.out.println("程序开始时间: "+now_star.getTime());
		System.out.println("程序开始时间: "+formatter_star.format(now_star.getTime()));
		System.out.println("===============================================");
		System.out.println("QA_run_info.1.7.3");
		//System.out.println();
		System.out.println("***********************************************");
		System.out.println();
		
		SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
		String day = formatter.format(now_star.getTime());
		
		int args_len = args.length;
		int Cover = 1;//0代表覆盖汇总表，1代表追加
		int Uploadtag = 0;//0代表所有表上传，1代表只上传更新表
		int Upload = 1;//设置是否需要上传至/wdmycloud/anchordx_cloud/杨莹莹/项目-生信-汇总表/，0代表不上传，1代表上传
		String dir = "./Ironman";
		String ExcelFormat = "xlsx";
		String Input = "/Src_Data1/analysis/Ironman/";
		String PutPath = "/wdmycloud/anchordx_cloud/杨莹莹/项目-生信-汇总表/" + day;
		String Path = null;
		
		for(int len = 0; len < args_len; len++){
			if( args[len].equals("-P") || args[len].equals("-p") ){
				Input = args[len+1];
			}else if(args[len].equals("-C") || args[len].equals("-c")){
				Cover = Integer.valueOf(args[len+1]);
			}else if(args[len].equals("-O") || args[len].equals("-o")){
				dir = args[len+1];
			}else if(args[len].equals("-F") || args[len].equals("-f")){
				Uploadtag = Integer.valueOf(args[len+1]);
			}else if(args[len].equals("-U") || args[len].equals("-u")){
				Upload = Integer.valueOf(args[len+1]);
			}
		}
		
		String Data_Aggregation_Path = dir + "/Data_Aggregation/";
		File DAP = new File(Data_Aggregation_Path);
		if(Cover == 1){
			if(DAP.exists() && DAP.isDirectory()){
				Copy_Old_File(Data_Aggregation_Path);//复制最新日期的文件
				//System.out.println ("Copy_Old_File");
			}else{
				System.out.println(Data_Aggregation_Path + "目录不存在");
			}
		}

		File fileInput = new File(Input);
		ExecutorService exe = Executors.newFixedThreadPool(15);//设置线程池最大线程数为15
		
		int Input_length = 0;
		String InputArr[] = Input.split("/");
		for(int i = 0; i < InputArr.length; i++){
			if(InputArr[InputArr.length-1].equals("Ironman")){
				Input_length = 0;
			}else if(InputArr[InputArr.length-2].equals("Ironman")){
				Input_length = 1;
			}else if(InputArr[InputArr.length-3].equals("Ironman")){
				Input_length = 2;
			}
		}
		if( Input_length == 0 ){
			Path = dir;
			for (File pathname : fileInput.listFiles())
			{
				exe.execute(new SubThread(dir, pathname, ExcelFormat, Input_length));
			}
		}else if( Input_length == 1 ){
			Path = dir + "/" + InputArr[InputArr.length - 1];
			exe.execute(new SubThread(dir, fileInput, ExcelFormat, Input_length));
			
		}else if( Input_length == 2 ){
			//dir = dir+"_6";
			Path = dir + "/" + InputArr[InputArr.length - 2]+ "/" + InputArr[InputArr.length - 1];
			exe.execute(new SubThread(dir, fileInput, ExcelFormat, Input_length));
			
		}else{
			System.out.println(Input + "是非法输入，请重新输入！");
			return;
		}
		
        //System.out.println("Path = " + Path);
        exe.shutdown();
		while (true)
        {
            if (exe.isTerminated()) //先让所有的子线程运行完，再运行主线程
            {
            	Data_Aggregation.Data_Aggregation_Main(dir+"/Data_Aggregation/"+day, Path, Cover, PutPath, Uploadtag, Upload);
                //System.out.println("结束了");
                break;
            }
            Thread.sleep(500);
        }
		
		Thread.sleep(3000);
		Upload_File(dir);//上传文件到阿里云端
		
		Calendar now_end = Calendar.getInstance();
		SimpleDateFormat formatter_end = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		System.out.println();
		System.out.println("==============================================");
		System.out.println("程序结束时间: "+now_end.getTime());
		System.out.println("程序结束时间: "+formatter_end.format(now_end.getTime()));
		System.out.println();
	}
	
	//上传文件
	public static void Upload_File(String PutPath)
	{
		String cmd = "/opt/local/bin/python35/python /var/script/alan/10k_api_script/qa_run_info_collections.py -path " + PutPath;
		try{
			Runtime.getRuntime().exec(cmd);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//获取最新文件所在目录
	public static String Get_New_FilePAth(String Path)
	{
		File file = new File(Path);
		int daynum = 0;
		for (File dir : file.listFiles()){
			if (dir.isDirectory()) { //如果是目录
				String dir_name = dir.getName(); //目录名
				if(daynum < Integer.valueOf(dir_name)){
					daynum = Integer.valueOf(dir_name);
				}else{
					continue;
				}
			}else{
				continue;
			}
		}
		return String.valueOf(daynum);
	}
	
	//复制最新日期的文件
	public static void Copy_Old_File(String Path){
		Calendar now_star = Calendar.getInstance();
		SimpleDateFormat formatter_Date = new SimpleDateFormat("yyyyMMdd");
    	String Day = formatter_Date.format(now_star.getTime());
		/*String cmd = "find " + Path + " -type f -name Unknown_*.xlsx";
		int daynum = 0;
		try {
			Process process = Runtime.getRuntime().exec(cmd);
			BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
			String line = null;
				
			while ((line = input.readLine()) != null) {
					File pathname = new File(line);
					String file_name = pathname.getName();
					String Suffix = file_name.substring(file_name.lastIndexOf(".")); //获取后缀名
					String Remove_suffix =  file_name.replaceAll(Suffix, ""); //去除后缀名
					String Arr[] = Remove_suffix.split("_");
					String newday = Arr[Arr.length - 1];
					
					if(daynum < Integer.valueOf(newday)){
						daynum = Integer.valueOf(newday);
						//System.out.println(daynum);
					}else{
						continue;
					}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		//System.out.println("big = "+daynum);*/
		String daynum = Get_New_FilePAth(Path);
		String cmd1 = "find " + Path + daynum + " -type f -name *" + daynum + "*.xlsx";
		my_mkdir( Path + "/" + Day );
		try {
			Process process = Runtime.getRuntime().exec(cmd1);
			BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
			String line = null;
			//System.out.println(cmd1);	
			while ((line = input.readLine()) != null) {
				//System.out.println(line);
				File pathname = new File(line);
				//String Folder = pathname.getParent();
				String file_name = pathname.getName();
				String Suffix = file_name.substring(file_name.lastIndexOf(".")); //获取后缀名
				String Remove_suffix =  file_name.replaceAll(Suffix, ""); //去除后缀名
				String Arr[] = Remove_suffix.split("_");
				String newname = null;
				for(int i = 0; i < Arr.length - 1; i++){
					if(i == 0){
						newname = Arr[i];
					}else{
						newname += "_" + Arr[i];
					}
				}
				String cmd2 = "cp " + line + " " + Path + "/" + Day + "/" + newname + "_" + Day + ".xlsx";
				Runtime.getRuntime().exec(cmd2);
				//System.out.println(newname);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
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
}

class SubThread extends Thread
{
	private String dir;
	private File pathname;
	private String ExcelFormat;
	private int inputlenght;

    public SubThread(String dir, File pathname, String ExcelFormat, int inputlenght)
    {
        this.dir = dir;
		this.pathname = pathname;
		this.ExcelFormat = ExcelFormat;
		this.inputlenght = inputlenght;
    }

    @Override
    public void run()
    {
    	Calendar now = Calendar.getInstance();
    	SimpleDateFormat formatter_Date = new SimpleDateFormat("yyyyMMdd");
    	String Day = formatter_Date.format(now.getTime());
    	if( inputlenght == 0 || inputlenght == 1){
			if (pathname.isDirectory()) { //如果是目录
				String dir_name = pathname.getName(); //目录名
				//System.out.println("dir_name = "+dir_name);
				String Sequencing_Info = dir + "/" + dir_name;
				//my_mkdir(Sequencing_Info);
				if( !(dir_name.startsWith("0.")) ){
					for (File porject : pathname.listFiles())
					{
						if (porject.isDirectory()) { //如果是目录
							//获取文件的绝对路径
							String Folder = porject.getParent();
							String Por_name = porject.getName(); //子目录名
							String Path = Folder + "/" + Por_name;
							//System.out.println("Path = "+Path);
							String Por_dir = Sequencing_Info + "/" + Por_name;
							String Plasma_ExcelName = "QA_run_info_" + Por_name + "_Plasma_" + Day + "." + ExcelFormat;//血浆表
							String Tissue_ExcelName = "QA_run_info_" + Por_name + "_Tissue_" + Day + "." + ExcelFormat;//组织表
							String Unknown_ExcelName = "QA_run_info_" + Por_name + "_Unknown_" + Day + "." + ExcelFormat;//其他的数据表
							String Plasma_Tsv = Por_dir + "/" + "QA_run_info_" + Por_name + "_Plasma_" + Day + ".tsv";//血浆
							String Tissue_Tsv = Por_dir + "/" + "QA_run_info_" + Por_name + "_Tissue_" + Day + ".tsv";//组织
							String Unknown_Tsv = Por_dir + "/" + "QA_run_info_" + Por_name + "_Unknown_" + Day + ".tsv";//其他
							String Plasma_Excel = Por_dir + "/" + Plasma_ExcelName;
							String Tissue_Excel = Por_dir + "/" + Tissue_ExcelName;
							String Unknown_Excel = Por_dir + "/" + Unknown_ExcelName;
							File Plasma_excel = new File(Plasma_Excel);
							File Tissue_excel = new File(Tissue_Excel);
							File Unknown_excel = new File(Unknown_Excel);
							my_mkdir(Por_dir);
							
							if(!Plasma_excel.exists() && !Plasma_excel.isFile()){
								CreateXlsx(Plasma_excel);//创建血浆表
							}
							if(!Tissue_excel.exists() && !Tissue_excel.isFile()){
								CreateXlsx(Tissue_excel);//创建组织表							
							}
							if(!Unknown_excel.exists() && !Unknown_excel.isFile()){
								CreateXlsx(Unknown_excel);//不明白的数据表
							}

							Excel_main(Plasma_Excel, Tissue_Excel, Unknown_Excel, Plasma_Tsv, Tissue_Tsv, Unknown_Tsv, Path, 1, null);
						}else{
							continue;
						}
					}		
				}
			}
    	}else if(inputlenght == 2){
			//获取文件的绝对路径
			String Folder = pathname.getParent();//父目录
			String Foldername = new File(Folder).getName();//父目录名
			String Por_name = pathname.getName(); //子目录名
			String Path = Folder + "/" + Por_name;
			//System.out.println("Path = "+Path);
			//String Sequencing_Info = dir + "/" + Por_name;
			String Por_dir = dir + "/" + Foldername + "/" + Por_name;
			String Plasma_ExcelName = "QA_run_info_" + Por_name + "_Plasma_" + Day + "." + ExcelFormat;//血浆表
			String Tissue_ExcelName = "QA_run_info_" + Por_name + "_Tissue_" + Day + "." + ExcelFormat;//组织表
			String Unknown_ExcelName = "QA_run_info_" + Por_name + "_Unknown_" + Day + "." + ExcelFormat;//其他的数据表
			String Plasma_Tsv = Por_dir + "/" + "QA_run_info_" + Por_name + "_Plasma_" + Day + ".tsv";//血浆
			String Tissue_Tsv = Por_dir + "/" + "QA_run_info_" + Por_name + "_Tissue_" + Day + ".tsv";//组织
			String Unknown_Tsv = Por_dir + "/" + "QA_run_info_" + Por_name + "_Unknown_" + Day + ".tsv";//其他
			String Plasma_Excel = Por_dir + "/" + Plasma_ExcelName;
			String Tissue_Excel = Por_dir + "/" + Tissue_ExcelName;
			String Unknown_Excel = Por_dir + "/" + Unknown_ExcelName;
			File Plasma_excel = new File(Plasma_Excel);
			File Tissue_excel = new File(Tissue_Excel);
			File Unknown_excel = new File(Unknown_Excel);
			my_mkdir(Por_dir);
			
			if(!Plasma_excel.exists() && !Plasma_excel.isFile()){
				CreateXlsx(Plasma_excel);//创建血浆表
			}
			if(!Tissue_excel.exists() && !Tissue_excel.isFile()){
				CreateXlsx(Tissue_excel);//创建组织表							
			}
			if(!Unknown_excel.exists() && !Unknown_excel.isFile()){
				CreateXlsx(Unknown_excel);//不明白的数据表
			}
			
			Excel_main(Plasma_Excel, Tissue_Excel, Unknown_Excel, Plasma_Tsv, Tissue_Tsv, Unknown_Tsv, Path, 1, null);
    	}
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
			if( i < 4 ){
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
	
	//调用linux命令获取符合要求的文件列表(跳过链接文件)
	public static ArrayList<String> SearchFile(String Input){
		ArrayList <String> data_ID = new ArrayList <String>();
		ArrayList <String> data_IDP = new ArrayList <String>();
		File fileInput = new File(Input);
		try{
			String InputArr[] = Input.split("/");
			if(InputArr.length == 3){
				for (File pathname : fileInput.listFiles())
				{
					if (pathname.isDirectory()) { //如果是目录
						String dir_name = pathname.getName(); //子目录名
						//System.out.println("dir_name = "+dir_name);
						if( !(dir_name.startsWith("0.")) ){
							//System.out.println("dir_name000 = "+dir_name);
							String cmd = "find " + Input + "/" + dir_name + " -type f";//查找该目录下所有文件（链接文件除外）
							Process process = Runtime.getRuntime().exec(cmd);
							BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
							String line = "";
							while ((line = input.readLine()) != null) {
								File file = new File(line);
								//获取文件的绝对路径
								String Folder = file.getParent();
								// 把文件名（basename）添加进列表
								String FileName = file.getName();
								//String regEx = "S.*[DR]P\\w(\\-DNA)?\\-[^\\-]*\\-[^\\-]*\\-[PFB][SREMC]\\d(\\-\\d)?";
								String regEx = "S.*_R1_001";
								String ID = Regular_Expression(FileName, regEx);
								String IDP = ID + "\t" + Folder;
								if( data_ID.contains(ID) || ID == null){
									continue;
								}else{
									data_ID.add(ID);
									data_IDP.add(IDP);
								}
							}
						}else{
							continue;
						}	
					}else{
						continue;
					}
				}
			}else{
				String cmd = "find " + Input + " -type f";//查找该目录下所有文件（链接文件除外）
				Process process = Runtime.getRuntime().exec(cmd);
				BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
				String line = "";
				while ((line = input.readLine()) != null) {
					File file = new File(line);
					//获取文件的绝对路径
					String Folder = file.getParent();
					// 把文件名（basename）添加进列表
					String FileName = file.getName();
					//String regEx = "S.*[DR]P\\w(\\-DNA)?\\-[^\\-]*\\-[^\\-]*\\-[PFB][SREMC]\\d(\\-\\d)?";
					String regEx = "S.*_R1_001";
					String ID = Regular_Expression(FileName, regEx);
					String IDP = ID + "\t" + Folder;
					if( data_ID.contains(ID) || ID == null){
						continue;
					}else{
						data_ID.add(ID);
						data_IDP.add(IDP);
					}
				}
			}
		}catch(Exception e){
			System.out.println("linux命令异常！！！！");
		}
		return data_IDP;
	}
	
	//调用正则表达式
	public static String Regular_Expression(String str, String regEx){
		String data = null;
		//编译正则表达式
		Pattern pattern = Pattern.compile(regEx);
		Matcher matcher = pattern.matcher(str);
		if( matcher.find() ){
			data = matcher.group();
		}
		return data;
	}
	
	//提取Sample_ID
	public static String Extract_Sample_ID (String Pre_lib_name){
		String str[] = Pre_lib_name.split("-");
		String strr = null;
		for(int i = 0; i <str.length; i++){
			if(str[i].equals("DPM")){
				if(str[i+1].equals("DNA")){
					strr = str[i+2]+"-"+str[i+3];
					break;
				}else{
					strr = str[i+1]+"-"+str[i+2];
					break;
				}
			}else{
				continue;
			}
		}
		return strr;
	}
	
	//提取Sequencing_Info
	public static String Extract_Sequencing_Info (String inputstr){
		String str[] = inputstr.split("/");
		String strr = null;
		for(int i = 0; i <str.length; i++){
			if(str[i].equals("Ironman")){	
				strr = str[i+1];
			}else{
				continue;
			}
		}
		return strr;
	}
	
	//调用linux命令
	public static String Linux_Cmd(String[] cmd){
		//String cmd = "ln -s -f " + Source_File + " " + Link_Path + "/" + str[3];
		String line = null;
		//System.out.println("linux cmd");
		try{
			//String cmd1 = "awk 'NR==8 {print $7 }' /Src_Data1/analysis/Ironman/161231_E00490_0087_AHF75NALXX/Ironman-CR003/06.clean_summary/S11-DPM-CR003-097-PS1-IRM_S11_R1_001.clean.fastq.gz_bismark_bt2_pe.sorted.bam.hsmetrics.txt";
			Process process = Runtime.getRuntime().exec(cmd);
			BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
			line = input.readLine();
			//System.out.println("linux cmd1");
			/*while ((line = input.readLine()) != null) {
				System.out.println(line);
			}*/
		}catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return line;
	}
	
	//判断一个Linux下的文件是否为链接文件，是返回true ,否则返回false
	public static boolean isLink(File file) {
		 String cPath = "";
			try {
			  cPath = file.getCanonicalPath();
		} catch (Exception ex) {
			System.out.println("文件异常："+file.getAbsolutePath());
		}
		return !cPath.equals(file.getAbsolutePath());
	}
	
	//4
	public static ArrayList<String> Find_flagstat_xls(String Path, String Start, String End){
		ArrayList<String> Data_list = new ArrayList<String>();
		String line = null;
		String data = null;
		int loog = 0;
		int filelog1 = 0;
		int filelog2 = 0;
		try{
			String cmd = "find " +  Path + " -type f -name flagstat.xls";
			Process process = Runtime.getRuntime().exec(cmd);
			BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
			//System.out.println("linux cmd1");
			while ((line = input.readLine()) != null) {
				//System.out.println(line);
				//Data_list.add(line);
				File file = new File(line);
				int log = 0;
				if( isLink(file) ){
					System.out.println("链接文件："+line);
					continue;
				}else{
					String encoding = "GBK";
					InputStreamReader read = new InputStreamReader(new FileInputStream(file),encoding);//考虑到编码格式
					BufferedReader bufferedReader = new BufferedReader(read);
					String lineTxt = null;
					while((lineTxt = bufferedReader.readLine()) != null){
						String str[] = lineTxt.split("\t");
						if( str[0].contains(Start) && str[0].endsWith(End)){
							//data = str[1];
							if( filelog1 == 0 ){
								if(Data_list.contains(str[1])){
									continue;
								}else{
									Data_list.add(str[1]);
								}
								filelog2++;
							}else if( (filelog1 != 0) && (filelog2 != 0) ){
								if(Data_list.contains(str[1])){
									continue;
								}else{
									Data_list.add(str[1]);
									System.out.println(line+"异常！其符合" +Start + "*"+ End + "的行的第二列的数据与另一个文件不一致！");
								}
							}
							log++;
						}else{
							continue;
						}
					}
				}
				filelog1++;
				if( log == 0 ){
					continue;
				}else if( log == 1 ){
					loog = 1;
				}else if( log > 1 ){
					loog = 1;
					System.out.println(line+"异常！包含多行" +Start + "*" + End + "的行！");
				}
			}
			if( loog == 0 ){
				String cmd1 = "find " +  Path + " -name " + Start +"*sorted.bam.flagstat";
				Process process1 = Runtime.getRuntime().exec(cmd1);
				BufferedReader input1 = new BufferedReader(new InputStreamReader(process1.getInputStream()));
				//System.out.println("linux cmd1");
				int i = 0;
				while ((line = input1.readLine()) != null) {
					String[] cmd4 = {"awk", "NR==5 {print $5 }", line};
					data = Linux_Cmd(cmd4);
					Data_list.add(data);
					i++;
				}
				if( i > 1 ){
					System.out.println(Path+"目录下有" + i + "多个符合" + Start + "*sorted.bam.flagstat" + "的文件！");
				}else if( i == 0 ){
					Data_list.add("NA");
				}
			}
			
			if(data == null){
				Data_list.add("NA");
			}
		}catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return Data_list;
	}
	
	//读取修改时间的方法  
	public static String getModifiedTime(String file){
		File f = new File(file);
		Calendar cal = Calendar.getInstance();
		long time = f.lastModified();
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy/MM/dd");
		cal.setTimeInMillis(time);
		return formatter.format(cal.getTime());
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
	
	//读表数据到列表，去除重复行
	public static ArrayList<String> ReadExcelData (File file){
		ArrayList<String> Data_list = new ArrayList<String>();
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
		return Data_list;
	}
	
	//写回数据
	public static void RewriteExcelData (File file){
		ArrayList<String> Data_list = ReadExcelData(file);
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
	
	/**
	 * 
	 * @param wb:excel文件对象
	 */
	//写xlsx格式文件
	public static void WriteXlsx(File file, String logo, String data, int rownum) throws Exception {
		FileInputStream is = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(is);
		XSSFSheet sheet = wb.getSheetAt(0);	//获取第1个工作薄

		int cellIndex = 0;
		XSSFRow xssfrow = sheet.getRow(0);
		
		// 获取当前工作薄的每一列
		for (int j = xssfrow.getFirstCellNum(); j < xssfrow.getLastCellNum(); j++) {
			XSSFCell xssfcell = xssfrow.getCell(j);

			if( xssfcell != null){
				//String cellValue = xssfcell.getStringCellValue().trim();
				String cellValue = String.valueOf(xssfcell).trim();
				if(cellValue.equals(logo)){
					cellIndex = j;	
				}else{
					continue;
				}
			}
		}
		try {
			is.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		int addrownum = rownum;
		// 指定行索引，创建一行数据, 行索引为当前最后一行的行索引 + 1
		int currentLastRowIndex = sheet.getLastRowNum();
		if( CheckRowNull(sheet.getRow(currentLastRowIndex)) == 0 ){
			addrownum = 0;
			//System.out.println("nullnull: " + currentLastRowIndex);
		}
		int newRowIndex = currentLastRowIndex + addrownum;
		XSSFRow newRow = null;
		if( addrownum == 0 ){
			newRow = sheet.getRow(newRowIndex);
		}else{
			newRow = sheet.createRow(newRowIndex);
		}
		
		// 创建一个单元格，设置其内的数据格式为字符串，并填充内容，其余单元格类同
		XSSFCell newGenderCell = newRow.createCell(cellIndex, Cell.CELL_TYPE_STRING);
		newGenderCell.setCellValue(data);

		// 首先要创建一个原始Excel文件的输出流对象！
		FileOutputStream excelFileOutPutStream = new FileOutputStream(file);
		// 将最新的 Excel 文件写入到文件输出流中，更新文件信息！
		wb.write(excelFileOutPutStream);
		// 执行 flush 操作， 将缓存区内的信息更新到文件上
		excelFileOutPutStream.flush();
		// 使用后，及时关闭这个输出流对象， 好习惯，再强调一遍！
		excelFileOutPutStream.close();
		wb.close();
	}
	
	//2
	public static String Extract_Pre_lib_name(String input){
		String Pre_lib_name = null;
		String Pre_lib_name_Arr[] = input.split("-");
		if( Pre_lib_name_Arr.length == 5 && !input.contains("-IRM")){
			for(int i = 1; i < Pre_lib_name_Arr.length-1; i++){
				if(i == 1){
					Pre_lib_name = Pre_lib_name_Arr[i];
				}else{
					Pre_lib_name += "-" + Pre_lib_name_Arr[i];
				}
			}
			String EndArr[] = Pre_lib_name_Arr[Pre_lib_name_Arr.length-1].split("_");
			Pre_lib_name += "-" + EndArr[0];
		}else{
			for(int i = 1; i < Pre_lib_name_Arr.length-1; i++){
				if(i == 1){
					Pre_lib_name = Pre_lib_name_Arr[i];
				}else{
					Pre_lib_name += "-" + Pre_lib_name_Arr[i];
				}
			}
		}
		return Pre_lib_name;
	}
	
	//写成tsv格式文本
	public static void WriteToTsv(String inputfile, String outputfile){
		ArrayList<String> Data_list = new ArrayList<String>();
		String data = null;
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
	
	//总调用
	public static void Excel_main(String Plasma_Excel, String Tissue_Excel, String Unknown_Excel, String Plasma_Tsv, String Tissue_Tsv, String Unknown_Tsv, String Input, int log, String PutPath){
		String data = null;
		File Plasma_File = new File(Plasma_Excel);
		File Tissue_File = new File(Tissue_Excel);
		File Unknown_File = new File(Unknown_Excel);
		HashMap<String, String> map_logo =  new HashMap<String, String>();
		ArrayList <String> ID_data = SearchFile(Input);
		ArrayList <String> Warning_List = new ArrayList <String>();
		ArrayList <String> Fail_List = new ArrayList <String>();
		String regEx = null;
		try{
			for(int i = 0; i < ID_data.size(); i++){
				//regEx = "[DR]P\\w(\\-DNA)?\\-[^\\-]*\\-[^\\-]*\\-[PFB][SREMC]\\d(\\-\\d)?";
				Warning_List.clear();
				Fail_List.clear();
				regEx = "[A-Z]{2}\\d{3}\\-\\d{3}";
				File file = null;
				int underlog = 0;
				int nulllog = 0;
				String ID_dataArr[] = ID_data.get(i).split("\t");
				//System.out.println("ID_dataArr[0] = " + ID_dataArr[0] );
				//System.out.println("ID_dataArr[1] = " + ID_dataArr[1] );
				String Sample_ID = null;
				//String Pre_lib_name = Regular_Expression(ID_dataArr[0], regEx);
				String Pre_lib_name = null;
				if(ID_dataArr[0].contains("-DPM") || ID_dataArr[0].contains("-DNA")){
					Pre_lib_name = Extract_Pre_lib_name(ID_dataArr[0]);
					String Pre_lib_name_Arr[] = Pre_lib_name.split("-");
					if(Pre_lib_name_Arr.length < 3){
						file = Unknown_File;
						Pre_lib_name = ID_dataArr[0];
						underlog = 1;
					}
					//System.out.println("Pre_lib_name = " + Pre_lib_name );
				}else{
					file = Unknown_File;
					Pre_lib_name = ID_dataArr[0];
					underlog = 1;
					//continue;
				}
				if( Pre_lib_name != null){
					if((Pre_lib_name.contains("-BD") || Pre_lib_name.contains("-PS")) && file == null ){
						file = Plasma_File;
					}else if( file == null ){
						file = Tissue_File;
					}
					Sample_ID = Regular_Expression(Pre_lib_name, regEx);
					if(Sample_ID == null){
						if(underlog == 0){
							Sample_ID = Extract_Sample_ID(Pre_lib_name);
						}
					}
				}else{
					continue;
				}
				String Sequencing_Info = Extract_Sequencing_Info(ID_dataArr[1]);
				
				map_logo.put("Sample ID", Sample_ID); // 0
				map_logo.put("Pre-lib name", Pre_lib_name); // 1
				map_logo.put("Identification name", ID_dataArr[0]); // 2
				map_logo.put("Sequencing info", Sequencing_Info); // 3
									
				String cmd = "find /Src_Data1/nextseq500 /Src_Data1/x10/ -name " + "*" + ID_dataArr[0] + "*";
				Process process = Runtime.getRuntime().exec(cmd);
				BufferedReader input = new BufferedReader(new InputStreamReader(process.getInputStream()));
				String Sequencing_file_name = null;
				if( (Sequencing_file_name = input.readLine()) != null ){
					map_logo.put("Sequencing file name", Sequencing_file_name); //4
				}else{
					map_logo.put("Sequencing file name", "NA"); // 4
				}
					
				String deduped_cvg = null;
				String deduped_hsmetrics = null;
				String origin_hsmetrics = null;
				String bisulfite = null;
				String QC_result = null;
				String deduped_bam = null;
				String deduped_insertSize = null;
				int tag = 0;
				
				String InputArr[] = Input.split("/");
				String cmd1 = null;
				if(InputArr.length < 3){
					cmd1 = "find " + Input + "/" + Sequencing_Info + " -name " + ID_dataArr[0] + "*";
				}else{
					cmd1 = "find " + Input + " -name " + ID_dataArr[0] + "*";
				}
				Process process1 = Runtime.getRuntime().exec(cmd1);
				BufferedReader input1 = new BufferedReader(new InputStreamReader(process1.getInputStream()));
				String line1 = null;
				while( (line1 = input1.readLine()) != null ){
					if(line1.endsWith("sorted.deduplicated.bam.perTarget.coverage")){
						deduped_cvg = line1;
						tag++;
					}else if(line1.endsWith("sorted.deduplicated.bam.hsmetrics.txt")){
						deduped_hsmetrics = line1;
						tag++;
					}else if(line1.endsWith("sorted.deduplicated.bam.insertSize.txt")){
						deduped_insertSize = line1;
						tag++;
					}else if(line1.endsWith("sorted.bam.hsmetrics.txt")){
						origin_hsmetrics = line1;
						tag++;
					}else if(line1.endsWith("fastq.gz_bismark_bt2_PE_report.txt")){
						bisulfite = line1;
						tag++;
					}else if(line1.endsWith("hsmetrics.QC.xls") || line1.endsWith("hsmetrics.QC.xlsx")){
						QC_result = line1;
						tag++;
					}else if(line1.endsWith("gz_bismark_bt2_pe.sorted.deduplicated.bam")){
						deduped_bam = line1;
						tag++;
					}else{
						continue;
					}	
				}
				if(tag == 0){
					continue;
				}
				String Map = null;
				String PF = null;
				String OnTarget = null;
				String BaitCvg = null;
				String DedupBaitCvg = null;
				String DedupCvg = null;
				String Target30X = null;
				String Uniformity = null;
				String CHG = null;
				String CHH = null;
				
				//5
				data = null;
				String Start = ID_dataArr[0];
				String End = "sorted.bam";
				int tar = 0;
				ArrayList<String> data4 = Find_flagstat_xls(Input, Start, End);
				for(int x = 0; x < data4.size(); x++){
					if(x == 0){
						if( !(data4.get(x).equals("NA")) ){
							data = data4.get(x);
							tar++;
						}
					}else{
						if( !(data4.get(x).equals("NA")) ){
							data += "__"+ data4.get(x);
							tar++;
						}
					}
					//map_logo.put("Mapping%", data);
				}
				if( tar != 0 ){
					map_logo.put("Mapping%", data);
				}else{
					map_logo.put("Mapping%", "NA");
				}
				if( tar == 1){
					Map = data;
				}else{
					Map = "NA";
				}
					 
				//6
				data = null;
				if( origin_hsmetrics != null ){
					String[] cmd5 = {"awk", "NR==8 {print $7 }", origin_hsmetrics};
					data = Linux_Cmd(cmd5);
					//System.out.println(data);
					map_logo.put("Total PF reads", data);
					PF = data;
					nulllog++;
				}else{
					PF = "NA";
					map_logo.put("Total PF reads", "NA");
				}
				
				//7
				data = null;
				if( deduped_insertSize != null ){
					String[] cmd6 = {"awk", "-F", "\t", "NR==8 {print $5}", deduped_insertSize};
					data = Linux_Cmd(cmd6);
					//System.out.println(data);
					map_logo.put("Mean_insert_size", data);
					nulllog++;
				}else{
					map_logo.put("Mean_insert_size", "NA");
				}
				
				//8
				data = null;
				if( deduped_insertSize != null ){
					String[] cmd7 = {"awk", "-F", "\t", "NR==8 {print $1}", deduped_insertSize};
					data = Linux_Cmd(cmd7);
					//System.out.println(data);
					map_logo.put("Median_insert_size", data);
					nulllog++;
				}else{
					map_logo.put("Median_insert_size", "NA");
				}
				   
				//9
				data = null;
				if( origin_hsmetrics != null ){
					String[] cmd8 = {"awk", "NR==8 {print $19}", origin_hsmetrics};
					data = Linux_Cmd(cmd8);
					//System.out.println(data);
					map_logo.put("On target%", data);
					OnTarget = data;
					nulllog++;
				}else{
					OnTarget = "NA";
					map_logo.put("On target%", "NA");
				}
								   
				//10
				data = null;
				if( origin_hsmetrics != null ){
					String[] cmd9 = {"awk", "NR==8 {print $22}", origin_hsmetrics};
					data = Linux_Cmd(cmd9);
					//System.out.println(data);
					map_logo.put("Pre-dedup mean bait coverage", data);
					BaitCvg = data;
					nulllog++;
				}else{
					BaitCvg = "NA";
					map_logo.put("Pre-dedup mean bait coverage", "NA");
				}
								   
				//11
				data = null;
				if( deduped_hsmetrics != null ){
					String[] cmd10 = {"awk", "NR==8 {print $22 }", deduped_hsmetrics};
					data = Linux_Cmd(cmd10);
					//System.out.println(data);
					map_logo.put("Deduped mean bait coverage", data);
					DedupBaitCvg = data;
					nulllog++;
				}else{
					DedupBaitCvg = "NA";
					map_logo.put("Deduped mean bait coverage", "NA");
				}
				
				   
				//12
				data = null;
				if( deduped_hsmetrics != null ){
					String[] cmd11 = {"awk", "NR==8 {print $23 }", deduped_hsmetrics};
					data = Linux_Cmd(cmd11);
					//System.out.println(data);
					map_logo.put("Deduped mean target coverage", data);
					DedupCvg = data;
					nulllog++;
				}else{
					DedupCvg = "NA";
					map_logo.put("Deduped mean target coverage", "NA");
				}				
				   
				//13
				data = null;
				if( deduped_hsmetrics != null ){
					String[] cmd12 = {"awk", "NR==8 {print $39 }", deduped_hsmetrics};
					data = Linux_Cmd(cmd12);
					//System.out.println(data);
					map_logo.put("% target bases > 30X", data);
					Target30X = data;
					nulllog++;
				}else{
					Target30X = "NA";
					map_logo.put("% target bases > 30X", "NA");
				}								
				
				//14
				data = null;
				if( QC_result != null ){
					String[] cmd13 = {"awk", "-F", "\t", "/UNIFORMITY/ {print $4}", QC_result};
					Process process13 = Runtime.getRuntime().exec(cmd13);
					BufferedReader input13 = new BufferedReader(new InputStreamReader(process13.getInputStream()));
					//System.out.println("linux cmd1");
					String line = null;
					while ((line = input13.readLine()) != null) {
						if(data == null){
							data = line;
						}else{
							data += "\t" + line;
						}
					}
					if(data == null){
						if(deduped_cvg != null){
							String[] cmd1_3 = {"Rscript", "/home/jiacheng_chuan/Ironman/DataArrangement/calcUniformity.R", deduped_cvg};
							Process process1_3 = Runtime.getRuntime().exec(cmd1_3);
							BufferedReader input1_3 = new BufferedReader(new InputStreamReader(process1_3.getInputStream()));
							//System.out.println("linux cmd1");
							String line1_3 = null;
							data = null;
							int log13 = 0; 
							while ((line1_3 = input1_3.readLine()) != null) {
								if(log13 == 1){
									String data13[] = line1_3.split("\t");
									data = data13[1];
									map_logo.put("Uniformity (0.2X mean)", data);
									Uniformity = data;
									//System.out.println("linux cmd1_3: "+data);
									break;
								}else{
									data += "\t" + line1_3;
								}
								log13++;
							}
						}else{
							map_logo.put("Uniformity (0.2X mean)", "NA");
							Uniformity = "NA";
						}
					}else{
						map_logo.put("Uniformity (0.2X mean)", data);
						Uniformity = data;
						nulllog++;
						//System.out.println("linux cmd111: "+data);
					}
				}else{
					//map_logo.put("Uniformity (0.2X mean)", "NA");
					if(deduped_cvg != null){
						String[] cmd1_3 = {"Rscript", "/home/jiacheng_chuan/Ironman/DataArrangement/calcUniformity.R", deduped_cvg};
						Process process1_3 = Runtime.getRuntime().exec(cmd1_3);
						BufferedReader input1_3 = new BufferedReader(new InputStreamReader(process1_3.getInputStream()));
						//System.out.println("linux cmd1");
						String line1_3 = null;
						data = null;
						int log13 = 0; 
						while ((line1_3 = input1_3.readLine()) != null) {
							if(log13 == 1){
								String data13[] = line1_3.split("\t");
								data = data13[1];
								map_logo.put("Uniformity (0.2X mean)", data);
								Uniformity = data;
								//System.out.println("linux cmd1_3: "+data);
								break;
							}else{
								data += "\t" + line1_3;
							}
							log13++;
						}
					}else{
						map_logo.put("Uniformity (0.2X mean)", "NA");
						Uniformity = "NA";
					}
				}
				//System.out.println("===================================== ");
				   
				//15
				data = null;
				if( bisulfite != null ){
					String[] cmd14 = {"awk", "/C methylated in CHG context/ {print $6}", bisulfite};
					data = Linux_Cmd(cmd14);
					//System.out.println(data);
					map_logo.put("C methylated in CHG context", data);
					CHG = data;
					nulllog++;
				}else{
					map_logo.put("C methylated in CHG context", "NA");
					CHG = "NA";
				}
				
				   
				//16
				data = null;
				if( bisulfite != null ){
					String[] cmd15 = {"awk", "/C methylated in CHH context/ {print $6}", bisulfite};
					data = Linux_Cmd(cmd15);
					//System.out.println(data);
					map_logo.put("C methylated in CHH context", data);
					CHH = data;
					nulllog++;
				}else{
					map_logo.put("C methylated in CHH context", "NA");
					CHH = "NA";
				}
				
				//17
				data = null;
				if( bisulfite != null ){
					String[] cmd15 = {"awk", "/C methylated in CpG context/ {print $6}", bisulfite};
					data = Linux_Cmd(cmd15);
					//System.out.println(data);
					map_logo.put("C methylated in CpG context", data);
					nulllog++;
				}else{
					map_logo.put("C methylated in CpG context", "NA");
				}
				  
				//18	
				data = null;
				if( QC_result != null ){
					data = QC_result;
					//System.out.println(data);
					map_logo.put("QC result", data);
					nulllog++;
				}else{
					map_logo.put("QC result", "NA");
				}
				
				//19
				data = null;
				if( QC_result != null ){
					//File file15 = new File(QC_result);
					data = getModifiedTime(QC_result);
					map_logo.put("Date of QC", data);
					nulllog++;
				}else{
					map_logo.put("Date of QC", "NA");
				}
				
				//20	
				data = null;
				if( deduped_bam != null ){
					data = deduped_bam;
					//System.out.println(data);
					map_logo.put("Path to sorted.deduped.bam", data);
					nulllog++;
				}else{
					map_logo.put("Path to sorted.deduped.bam", "NA");
				}
				
				//21
				data = null;
				if( QC_result != null ){
					//File file15 = new File(QC_result);
					data = getModifiedTime(QC_result);
					map_logo.put("Date of path update", data);
					nulllog++;
				}else{
					map_logo.put("Date of path update", "NA");
				}
				   
				//22	
				data = null;
				if( origin_hsmetrics != null ){
					String[] cmd20 = {"awk", "NR==8 {print $1 }", origin_hsmetrics};
					data = Linux_Cmd(cmd20);
					//System.out.println(data);
					map_logo.put("Bait set", data);
					nulllog++;
				}else{
					map_logo.put("Bait set", "NA");
				}
				
				//23 log2(CPM+1)
				data = null;
				String cmd23 =  "find " + Input + " -name " + ID_dataArr[0] + "*WM.stat";
				Process process23 = Runtime.getRuntime().exec(cmd23);
				BufferedReader input23 = new BufferedReader(new InputStreamReader(process23.getInputStream()));
				String line23 = null;
				while( (line23 = input23.readLine()) != null ){
					if(data == null){
						data = line23;
					}else{
						data += " " + line23;
					}
				}
				if(data == null){
					map_logo.put("log2(CPM+1)", "NA");
				}else{
					map_logo.put("log2(CPM+1)", data);
				}
				
				//24 	
				String Result = null;
				//Map
				if( !(Map.equals("NA")) ){
					String NumStrArr[] = Map.split("%");
					double num = Double.valueOf(NumStrArr[0]);
					if( num > 50 ){
						Result = "Pass";
					}else{
						Result = "Warning";
						Warning_List.add("Map");
					}
				}
				//PF
				if( !(PF.equals("NA")) ){
					int num = Integer.valueOf(PF);
					if( !(num > 10) ){
						Result = "Warning";
						Warning_List.add("PF");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//OnTarget
				if( !(OnTarget.equals("NA")) ){
					double num = Double.valueOf(OnTarget);
					if( !(num > 0.4) ){
						Result = "Warning";
						Warning_List.add("OnTarget");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//BaitCvg
				if( !(BaitCvg.equals("NA")) ){
					double num = Double.valueOf(BaitCvg);
					if( !(num > 500) ){
						Result = "Warning";
						Warning_List.add("BaitCvg");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//DedupBaitCvg
				if( !(DedupBaitCvg.equals("NA")) ){
					double num = Double.valueOf(DedupBaitCvg);
					if( !(num > 100) ){
						Result = "Warning";
						Warning_List.add("DedupBaitCvg");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//DedupCvg
				if( !(DedupCvg.equals("NA")) ){
					double num = Double.valueOf(DedupCvg);
					if( !(num > 50) ){
						Result = "Fail";
						Fail_List.add("DedupCvg");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//Target30X
				if( !(Target30X.equals("NA")) ){
					double num = Double.valueOf(Target30X);
					if( !(num > 0.5) ){
						Result = "Fail";
						Fail_List.add("Target30X");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//Uniformity
				if( !(Uniformity.equals("NA")) ){
					double num = Double.valueOf(Uniformity);
					if( !(num > 0.85) ){
						if( Result == null ){
							Result = "Warning";
						}else if( !(Result.equals("Fail")) ){
							Result = "Warning";
						}
						Warning_List.add("Uniformity");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//CHG
				if( !(CHG.equals("NA")) ){
					String NumStrArr[] = CHG.split("%");
					double num = Double.valueOf(NumStrArr[0]);
					if( num >= 3 ){
						Result = "Fail";
						Fail_List.add("CHG");
					}else if( num < 3 && num >= 1 ){
						if( Result == null ){
							Result = "Warning";
						}else if( !(Result.equals("Fail")) ){
							Result = "Warning";
						}
						Warning_List.add("CHG");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//CHH
				if( !(CHH.equals("NA")) ){
					String NumStrArr[] = CHH.split("%");
					double num = Double.valueOf(NumStrArr[0]);
					if( num >= 3 ){
						Result = "Fail";
						Fail_List.add("CHH");
					}else if( num < 3 && num >= 1 ){
						if( Result == null ){
							Result = "Warning";
						}else if( !(Result.equals("Fail")) ){
							Result = "Warning";
						}
						Warning_List.add("CHH");
					}else{
						if(Result == null){
							Result = "Pass";
						}
					}
				}
				//last
				if(Result == null){
					Result = "NA";
				}
				map_logo.put("Sample QC", Result);
				nulllog++;
				
				//25 
				data = null;
				if( Fail_List.size() != 0 ){
					for(int t = 0; t < Fail_List.size(); t++){
						if(t == 0){
							data = Fail_List.get(t);
						}else{
							data += ";" + Fail_List.get(t);
						}
					}
					map_logo.put("Failed QC Detail", data);
				}else{
					map_logo.put("Failed QC Detail", "NA");
				}
				nulllog++;
				
				//26 
				data = null;
				if( Warning_List.size() != 0 ){
					for(int t = 0; t < Warning_List.size(); t++){
						if(t == 0){
							data = Warning_List.get(t);
						}else{
							data += ";" + Warning_List.get(t);
						}
					}
					map_logo.put("Warning QC Detail", data);
				}else{
					map_logo.put("Warning QC Detail", "NA");
				}
				nulllog++;
						
				int rownum = 1;
				//String LastRowDdata = null;
				//String LastRowDdataDefore = ReadLastRow(file);
				//System.out.println("LastRowDdataDefore: "+LastRowDdataDefore);
				for ( String key : map_logo.keySet() ){
					//getFromExcel(New_File, key, map_logo.get(key), rownum);
					if( log == 0 ){
						//Excel_xls.WriteXls(file, key, map_logo.get(key), rownum);
					}else if( log == 1 ){
						WriteXlsx(file, key, map_logo.get(key), rownum);
					}
					rownum = 0;
					//System.out.println(key);
					//System.out.println(map_logo.get(key));
				}	
			}
		}catch(Exception e){
			e.printStackTrace();
		}
		if( log == 0 ){
			//while(Excel_xls.RemoveNullRow_xls(file) != 0){
				//Excel_xls.RemoveNullRow_xls(file);//去除空行
				//System.out.println("i= "+i);
			//}
			//Excel_xls.RewriteExcelData_xls(file);
			//Excel_xls.WriteToTxt_xls(New_File, Txt_File);
		}else if( log == 1 ){
			// String LastRowDdataLater = ReadLastRow(file);
			// System.out.println("LastRowDdataLater: "+LastRowDdataLater);
			
			while(RemoveNullRow(Plasma_File) != 0){
				RemoveNullRow(Plasma_File);//去除空行
			}
			RewriteExcelData(Plasma_File);
			WriteToTsv(Plasma_Excel, Plasma_Tsv);
			
			while(RemoveNullRow(Tissue_File) != 0){
				RemoveNullRow(Tissue_File);//去除空行
			}
			RewriteExcelData(Tissue_File);
			WriteToTsv(Tissue_Excel, Tissue_Tsv);
			
			while(RemoveNullRow(Unknown_File) != 0){
				RemoveNullRow(Unknown_File);//去除空行
			}
			RewriteExcelData(Unknown_File);
			WriteToTsv(Unknown_Excel, Unknown_Tsv);
			
			/*if(LastRowDdataDefore.equals(LastRowDdataLater)){
			RemoveLastRow(file);//如果最后两行重复，则去除最后一行
		 	System.out.println("已删除。");
			}*/
			
		}
		//SSh_Upload_File(New_File, PutPath);
		//System.out.println(Input + "已完成。");
	}	
	
}
