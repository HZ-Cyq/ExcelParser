package com.tools.parserconfig;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.sql.Ref;
import java.util.ArrayList;

import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;

public class XlsxToServerLog {
	
	static String strPackageName = "aaa";

	public static void main(String[] args) throws IOException 
	{
		if(args.length > 0)
		{
			strPackageName = args[0];
		}
		try 
		{
			File fi = new File("数据打点索引.xlsx");
			readXlsx(fi, "./");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void readXlsx(File file, String javaPath) throws Exception {
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(file.getPath());
		
		XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
		if (xssfSheet == null)
			return;

		String fileNameString = xssfSheet.getSheetName();
		String str = fileNameString.substring(0, 1);

		System.err.println("解析文件 " + fileNameString);
		createJava(xssfSheet, javaPath);
	}


	private static String CreateCodes(XSSFRow curRow)
	{
		int id = Integer.parseInt(curRow.getCell(0).getStringCellValue());
		String funcName = curRow.getCell(1).getStringCellValue();
		String simpleDesc = curRow.getCell(2)!=null ? curRow.getCell(2).getStringCellValue() : "";
		String fullDesc = curRow.getCell(3)!=null ? curRow.getCell(3).getStringCellValue() : "";
		
		fullDesc = fullDesc.replaceAll("\r", "");
		fullDesc = fullDesc.replaceAll("\n", "\n     *   ");

		ArrayList<String> arrParamName = new ArrayList<String>();
		ArrayList<String> arrParamDesc = new ArrayList<String>();
		ArrayList<String> arrParamType = new ArrayList<String>();
		ArrayList<String> arrNodeName  = new ArrayList<String>();
		String[] arr = { "long", "long", "long", "int", "int", "int", "int", "int", "String", "String",  "Timestamp", "Timestamp" };
		String[] noder = { "setLong0", "setLong1", "setLong2", "setInt0", "setInt1", "setInt2", "setInt3", "setInt4",  "setStr0", "setStr1", "setDatetime0", "setDatetime1" };
		for(int n=0; n<arr.length; n++)
		{
			int i=n+4;
			if(curRow.getCell(i) == null)
				continue;
			
			//curRow.getCell(i).setCellType(XSSFCell.CELL_TYPE_STRING);
			String s = curRow.getCell(i).getStringCellValue();
			if(s.equals(""))
				continue;
			String[] ss = s.split("\\|");
			arrParamDesc.add(ss[0]);
			arrParamName.add(ss[1]);
			arrParamType.add(arr[i-4]);
			arrNodeName.add(noder[i-4]);
		}
		
		String ret = "";
		ret += String.format("    /** 类型:%s 说明:%s 特殊说明: %s\n     *", id, simpleDesc, fullDesc.equals("")?"无":fullDesc);
		for(int i=0; i<arrParamName.size(); i++)
		{
			ret += String.format(" (%s=%s %s)", arrParamName.get(i), arrParamType.get(i), arrParamDesc.get(i));
		}
		
		//数据库查询语句
		{
			ret += "\n     * 数据库查询:";
			
			
			String params = "";
			for(int i=0; i<arrParamType.size(); i++)
			{
				String p = String.format("%s as %s", arrNodeName.get(i), arrParamName.get(i));
				if(i != arrParamType.size()-1)
					p += ",";
				params += p;
			}
			
			String sql = String.format("select time, %s from hc_records where type=%s;", params, id);
			ret += sql;
		}
		
		ret += String.format("*/\n");
		ret += String.format("    public static void %s(", funcName);
		for(int i=0; i<arrParamType.size(); i++)
		{
			ret += String.format("%s %s", arrParamType.get(i), arrParamName.get(i));
			if(i != arrParamName.size()-1)
			{
				ret += ", ";
			}
		}
		ret += "){ ";
		ret += String.format("int n=%s; DatabaseLogManager.LogNode node = DatabaseLogManager.GetInstance().createLogNode(n); ", id);
		for(int i=0; i<arrParamType.size(); i++)
		{
			ret += String.format("node.%s(%s);", arrNodeName.get(i), arrParamName.get(i));
		}
		if(arrParamName.size() > 0 && arrParamName.get(0).equals("playerid") && arrNodeName.size() > 0 && arrNodeName.get(0).equals("l0"))//可选择记录机器人的log
		{
			ret += " if(!DatabaseLogManager.USE_ROBOT_LOG && (node.l0/1000000000L)%10>=1){ return; }";
		}
		ret += "DatabaseLogManager.GetInstance().addLog(node);}\n\n";
		
		return ret;
	}
	
	private static void createJava(XSSFSheet xssfSheet, String javaPath) 
	{
		String javaName = xssfSheet.getSheetName();
		javaName = "DBGameLog";
		String fileName = javaPath + javaName + ".java";

		StringBuffer sb = new StringBuffer();
		sb.append("package com.mobi.log;\n");
		sb.append("import com.mobi.database.DatabaseLogManager;\n\n");
		sb.append("public class DBGameLog\n{\n");
		int lastRowNumber = GetSheetLastRowNum(xssfSheet);
		int lastCellNumber = GetSheetLastCellNum(xssfSheet);
		
		for (int i = 0; i <= lastRowNumber; i++) 
		{
			if(i == 0)
				continue;
			String codes = CreateCodes(xssfSheet.getRow(i));
			sb.append(codes);
		}
		sb.append("}");
		writeFile(sb.toString(), fileName, false);
	}

	
	public static int GetSheetLastRowNum(XSSFSheet xssfSheet)//最后一行行号内部索引  0-xxx
	{
		int lastRowNum = -1;
		int rr = xssfSheet.getLastRowNum();
		
		for (int rowNum = 0; rowNum <= xssfSheet.getLastRowNum(); rowNum++)
		{
			XSSFRow curRow = xssfSheet.getRow(rowNum);
			if(curRow == null)
				break;
			XSSFCell cell0 = curRow.getCell(0);
			if(cell0 == null)
				break;
			cell0.setCellType(Cell.CELL_TYPE_STRING);
			if(cell0.getStringCellValue().equals(""))
				break;
			lastRowNum = rowNum;
		}
		return lastRowNum;
	}
	
	public static int GetSheetLastCellNum(XSSFSheet xssfSheet)//列的条目数
	{
		int lastCellNum = 0;
		
		if(xssfSheet.getLastRowNum() >= 0)
		{
			XSSFRow curRow = xssfSheet.getRow(0);
			int cc = curRow.getLastCellNum();
			for (int cellNum = 0; cellNum < curRow.getLastCellNum(); cellNum++) 
			{
				XSSFCell c = curRow.getCell(cellNum);
				if(c == null)
					break;
				c.setCellType(Cell.CELL_TYPE_STRING);
				if(c.getStringCellValue().equals(""))
				{
					break;
				}
				lastCellNum++;
			}
		}
		return lastCellNum;
	}
	
	

	public static void writeFile(String data, String filePath, boolean flag) {

		try {
			File file = new File(filePath);

			if (!file.exists()) 
			{
				file.createNewFile();
			}

			final byte[] bom = new byte[] { (byte) 0xEF, (byte) 0xBB,
					(byte) 0xBF };

			FileOutputStream fos = new FileOutputStream(file, false);
			if (flag)
				fos.write(bom);
			OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
			osw.write(data);
			osw.flush();
			osw.close();

			System.out.println("创建" + file.getName() + "完成！文件位置："
					+ file.getAbsolutePath());

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/** 首字母大写 */
	public static String getFirstWordUp(String str) {
		return str.replaceFirst(str.substring(0, 1), str.substring(0, 1)
				.toUpperCase());
	}

}