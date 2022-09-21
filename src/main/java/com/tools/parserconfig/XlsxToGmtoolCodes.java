package com.tools.parserconfig;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.sql.Ref;
import java.util.ArrayList;

import javax.sql.rowset.CachedRowSet;

import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.omg.CosNaming.NamingContextPackage.NotEmpty;

public class XlsxToGmtoolCodes 
{
	public class ParamNode
	{
		String type = "";
		String name = "";
		String desc = "";
	}
	
	
	static String strPackageName = "aaa";
	static int MAX_PARAM_COUNT = 8;
	static int OFFSET = 9;

	public static void main(String[] args) throws IOException 
	{
		if(args.length > 0)
		{
			strPackageName = args[0];
		}
		try 
		{
			File fi = new File("GM工具表.xlsx");
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
		createcsharp(xssfSheet, javaPath);
	}

	enum LANG
	{
		JAVA,
		CSHARP,
	}
	
	private static String ConvertToPrimitiveType(String type, LANG lang) 
	{
		if(lang == LANG.JAVA)
		{
			switch (type) {
			case "string":
				return "String";
			case "json":
				return "JSONObject";
			default:
				break;
			}
		}
		else if(lang == LANG.CSHARP)
		{
			switch (type) {
			case "json":
				return "JsonData";
			default:
				break;
			}
		}
		return type;
	}
	
	private static String ConvertToGetPrimitiveJsonType(String type) 
	{
		if(type.equals("string"))
			return "String";
		if(type.equals("int"))
			return "Int";
		if(type.equals("float"))
			return "Float";
		if(type.equals("json"))
			return "JSONObject";
		return type;
	}

	private static String CreateJavaCodes(XSSFRow curRow)
	{
		if(curRow.getCell(1) != null) curRow.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
		if(curRow.getCell(2) != null) curRow.getCell(2).setCellType(Cell.CELL_TYPE_STRING);
		if(curRow.getCell(6) != null) curRow.getCell(6).setCellType(Cell.CELL_TYPE_STRING);
		if(curRow.getCell(8) != null) curRow.getCell(8).setCellType(Cell.CELL_TYPE_STRING);
		
		int id = Integer.parseInt(curRow.getCell(0).getStringCellValue());
		String funcName = curRow.getCell(1).getStringCellValue();
		String simpleDesc = curRow.getCell(2)!=null ? curRow.getCell(2).getStringCellValue() : "";
		String groupName = curRow.getCell(6)!=null ? curRow.getCell(6).getStringCellValue() : "";
		String caseName = curRow.getCell(OFFSET+MAX_PARAM_COUNT)!=null? curRow.getCell(OFFSET+MAX_PARAM_COUNT).getStringCellValue() : "";
		boolean isSingleServerCmd = curRow.getCell(8).getStringCellValue().equals("1");
		
		//groupName = groupName.replaceAll("\r", "");
		//groupName = groupName.replaceAll("\n", "\n     *   ");

		ArrayList<ParamNode> arrParams = new ArrayList<XlsxToGmtoolCodes.ParamNode>();

		for(int n=0; n<MAX_PARAM_COUNT; n++)
		{
			int i=n+OFFSET;
			if(curRow.getCell(i) == null)
				continue;
			
			//curRow.getCell(i).setCellType(XSSFCell.CELL_TYPE_STRING);
			String s = curRow.getCell(i).getStringCellValue();
			if(s.equals(""))
				continue;
			String[] ss = s.split("\\|");
			XlsxToGmtoolCodes tmp = new XlsxToGmtoolCodes();
			ParamNode node = tmp.new ParamNode();
			node.desc = ss[0];
			node.type = ss[1];
			node.name = ss[2];
			arrParams.add(node);
		}
		
		String ret = "\n            ";
		String strParams = "";
		{
			for(ParamNode node : arrParams)
			{
				strParams += String.format("%s=%s, ", node.name, node.desc);
			}
		}
		ret += String.format("// 功能:%s 参数: %s ", groupName+"/"+simpleDesc, strParams);
		ret += "\n            case \""+caseName+"\": {";
		
		for(ParamNode node : arrParams)
		{
			ret += String.format(" %s %s=params.get%s(\"%s\");", ConvertToPrimitiveType(node.type, LANG.JAVA), node.name, ConvertToGetPrimitiveJsonType(node.type), node.name);
		}
		ret += "JSONObject ret=new JSONObject(); ret.put(\"result\", \"ok\"); ret.put(\"err\", \"\");"; 
		if(isSingleServerCmd)
			ret += String.format("return DoCmd.%s(gm, serversessionid, ret", funcName);
		else
			ret += String.format("return DoCmd.%s(gm, ret", funcName);
		if(arrParams.size() > 0)
			ret += ",";
		
		for(int i=0; i<arrParams.size(); i++)
		{
			ParamNode node = arrParams.get(i);
			ret += node.name;
			if(i != arrParams.size()-1)
			{
				ret += ',';
			}
		}
		ret += "); }\n";
		return ret;
	}
	
	private static String CreateCsharpCodes(XSSFRow curRow)
	{
		int id = Integer.parseInt(curRow.getCell(0).getStringCellValue());
		String funcName = curRow.getCell(1).getStringCellValue();
		String simpleDesc = curRow.getCell(2)!=null ? curRow.getCell(2).getStringCellValue() : "";
		String groupName = curRow.getCell(6)!=null ? curRow.getCell(6).getStringCellValue() : "";
		String caseName = curRow.getCell(OFFSET+MAX_PARAM_COUNT)!=null? curRow.getCell(OFFSET+MAX_PARAM_COUNT).getStringCellValue() : "";
		
		//groupName = groupName.replaceAll("\r", "");
		//groupName = groupName.replaceAll("\n", "\n     *   ");

		ArrayList<ParamNode> arrParams = new ArrayList<XlsxToGmtoolCodes.ParamNode>();

		for(int n=0; n<MAX_PARAM_COUNT; n++)
		{
			int i=n+OFFSET;
			if(curRow.getCell(i) == null)
				continue;
			
			//curRow.getCell(i).setCellType(XSSFCell.CELL_TYPE_STRING);
			String s = curRow.getCell(i).getStringCellValue();
			if(s.equals(""))
				continue;
			String[] ss = s.split("\\|");
			XlsxToGmtoolCodes tmp = new XlsxToGmtoolCodes();
			ParamNode node = tmp.new ParamNode();
			node.desc = ss[0];
			node.type = ss[1];
			node.name = ss[2];
			arrParams.add(node);
		}
		
		String ret = "\n            ";
		String strParams = "";
		{
			for(ParamNode node : arrParams)
			{
				strParams += String.format("%s=%s, ", node.name, node.desc);
			}
		}
		ret += String.format("// 功能:%s 参数: %s ", groupName+"/"+simpleDesc, strParams);
		ret += "\n    public static JsonData Request_"+funcName+"(";
		
		for(int i=0; i<arrParams.size(); i++)
		{
			ParamNode node = arrParams.get(i);
			ret += String.format("%s %s", ConvertToPrimitiveType(node.type, LANG.CSHARP), node.name);
			if(i != arrParams.size()-1)
			{
				ret += ',';
			}
		}
		
		ret += ") { JsonData d = new JsonData(); ";
		ret += String.format("d[\"cmd\"]=\"%s\"; ", caseName);
		
		if(arrParams.size() > 0)
		{
			ret += "JsonData para=new JsonData();";
			for(int i=0; i<arrParams.size(); i++)
			{
				ParamNode node = arrParams.get(i);
				if(node.type.equals("json") || node.type.equals("string"))
					ret += String.format("if(%s!=null){para[\"%s\"]=%s;}", node.name, node.name, node.name);
				else
					ret += String.format(" para[\"%s\"]=%s;", node.name, node.name);
			}
			ret += "d[\"para\"] = para; ";
		}
		
		
		ret += "return d; }\n";
		return ret;
	}
	
	private static void createJava(XSSFSheet xssfSheet, String javaPath) 
	{
		String javaName = xssfSheet.getSheetName();
		javaName = "ServerLog";
		String fileName = javaPath + "DispatchCmds" + ".java";

		StringBuffer sb = new StringBuffer();
		sb.append("package com.huanmedia.xserver.gmtool.request;\n");
		sb.append("import net.sf.json.JSONObject;\n\n");
		sb.append("import com.huanmedia.xserver.gmtool.gmsession.GMAccount;");
		sb.append("//本文件由工具生成，不用手改\n");
		sb.append("public class DispatchCmds\n{\n");
		sb.append("\tpublic static JSONObject DoCmd(GMAccount gm, int serversessionid, String cmdxxx, JSONObject params)\n    {\n        switch (cmdxxx)\n\t\t{\n");
		
		int lastRowNumber = GetSheetLastRowNum(xssfSheet);
		int lastCellNumber = GetSheetLastCellNum(xssfSheet);
		
		for (int i = 0; i <= lastRowNumber; i++) 
		{
			if(i <= 2)
				continue;
			String codes = CreateJavaCodes(xssfSheet.getRow(i));
			sb.append(codes);
		}
		sb.append("            default: return null;\n        }\n    }\n}");
		writeFile(sb.toString(), fileName, false);
	}
	
	private static void createcsharp(XSSFSheet xssfSheet, String javaPath) 
	{
		String javaName = xssfSheet.getSheetName();
		javaName = "ServerLog";
		String fileName = javaPath + "AssembleRequestJson" + ".cs";

		StringBuffer sb = new StringBuffer();
		sb.append("using System.Collections;\n");
		sb.append("using System.Collections.Generic;\n");
		sb.append("using LitJson;\n");
		sb.append("//本文件由工具生成，不用手改\n\n");
		sb.append("public class AssembleRequestJson\n{\n");
		
		int lastRowNumber = GetSheetLastRowNum(xssfSheet);
		int lastCellNumber = GetSheetLastCellNum(xssfSheet);
		
		for (int i = 0; i <= lastRowNumber; i++) 
		{
			if(i <= 2)
				continue;
			String codes = CreateCsharpCodes(xssfSheet.getRow(i));
			sb.append(codes);
		}
		sb.append("\n}");
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