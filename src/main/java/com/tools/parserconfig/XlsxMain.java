package com.tools.parserconfig;

import javafx.util.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.xml.crypto.dsig.keyinfo.KeyValue;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

public class XlsxMain
{

    static String strPackageName = "aaa";
    static ArrayList<String> listAvgScripts = new ArrayList<>();

    public static void main(String[] args) throws IOException
    {
        if (args.length > 0)
        {
            strPackageName = args[0];
        }
        try
        {
            File xml = new File("Localize/");
            File cshap = new File("csharp/");
            File csharpHF = new File("csharp/hf/");
            File java = new File("java/");

            if (true)
            {
                File[] files = xml.listFiles();
                if(files == null) {
                    files = new File[0];
                }
                for (File f : files)
                {
                    if (!f.isDirectory())
                        f.delete();
                }
            }

            {
                File file = new File("excel/excel");
                if (file.isDirectory())
                {
                    File[] flist = file.listFiles();
                    for (File fi : flist)
                    {
                        if (!CheckIfValidExcel(fi))
                            continue;

                        readXlsx(fi, "Localize/", "csharp/", "java/", true, true, true, false);
                    }
                }
            }

        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    private static boolean CheckIfValidExcel(File fi)
    {
        boolean isValid = true;
        if (!fi.getName().endsWith(".xlsx"))
        {
            isValid = false;
        }
        else if (fi.isDirectory())
        {
            System.err.println("不解析文件夹");
            isValid = false;
        }
        else if (!fi.getName().endsWith(".xlsx"))
        {
            System.err.println("不解析非excel文件");
            isValid = false;
        }
        else if (fi.getName().contains("~$"))
        {
            System.err.println("不解析excel缓存: " + fi.getName());
            isValid = false;
        }
        return isValid;
    }

    public static void parser(String src, String xmlPath, String csPath,
                              String javaPath)
    {
        try
        {
            File file = new File(src);
            if (file.isDirectory())
            {
                File[] files = file.listFiles();
                for (File fi : files)
                {
                    if (fi.getName().endsWith(".xlsx"))
                        System.out.println("读取配置表：" + fi.getName());
                    readXlsx(fi, xmlPath, csPath, javaPath, true, true, true, false);
                }
            }


        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    public static void readXlsx(File file, String xmlPath, String csPath,
                                String javaPath, boolean bCreateXml, boolean bCreateJava, boolean bCreateCSharp, boolean bAVGScript) throws Exception
    {
        System.out.println("读取配置表：" + file.getName());

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(file.getPath());
        // 循环工作表Sheet
        for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++)
        {
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
            if (xssfSheet == null)
            {
                continue;
            }
            String fileNameString = xssfSheet.getSheetName();
            String str = fileNameString.substring(0, 1);

            if (!str.matches("^[A-Za-z]+$"))
            {
                System.err.println("非英文,拒绝解析文件！" + fileNameString);
                continue;
            }
            if (str.contains("Sheet"))
            {
                System.err.println("拒绝解析未改名Sheet！" + fileNameString);
                continue;
            }
            System.err.println("解析文件 " + fileNameString);
            if (bCreateXml)
                createXml(xmlPath, xssfSheet);
        }
    }

    private static void createCSharp(XSSFSheet xssfSheet, String cshrpPath)
    {
        String javaName = xssfSheet.getSheetName();
        if (hasPostfix(javaName))
            return;

        String head = javaName.substring(0, 3);
        if (head.equals("HF_"))
        {
            javaName = javaName.substring(3, javaName.length());
            javaName = getFirstWordUp(javaName) + "ConfigModel";
            javaName = "HF_" + javaName;
        }
        else
            javaName = getFirstWordUp(javaName) + "ConfigModel";


        String fileName = cshrpPath + javaName + ".cs";
        boolean bHF = false;
        if (javaName.contains("HF_"))
        {
            bHF = true;
            fileName = cshrpPath + "hf/" + javaName + ".cs";
        }


        StringBuffer sb = new StringBuffer();
        String[] zhuShi = null;
        String[] fields = null;
        String[] type = null;
        int len = 0;
        int lastRowNumber = GetSheetLastRowNum(xssfSheet);
        int lastCellNumber = GetSheetLastCellNum(xssfSheet);
        for (int rowNum = 0; rowNum <= lastRowNumber; rowNum++)
        {
            XSSFRow xssfRow = xssfSheet.getRow(rowNum);
            len = lastCellNumber;
            if (rowNum > 2)
                break;
            if (rowNum == 0)
            {
                if (zhuShi == null)
                {
                    zhuShi = new String[len];
                    fields = new String[len];
                    type = new String[len];
                }
            }
            for (int cellNum = 0; cellNum < lastCellNumber; cellNum++)
            {
                XSSFCell xssfCell = xssfRow.getCell(cellNum);
                if (xssfCell == null)
                {
                    if (rowNum == 2)
                    {
//						type[cellNum] = "string";
                    }
                    continue;
                }
                if (rowNum == 0)
                {
                    zhuShi[cellNum] = xssfCell.getStringCellValue();
                }
                else if (rowNum == 1)
                {
                    fields[cellNum] = xssfCell.getStringCellValue();
                }
                else if (rowNum == 2)
                {
                    String val = xssfCell.getStringCellValue();
                    if (val.equals("int"))
                    {
                        type[cellNum] = "int";
                    }
                    else if (val.equals("string"))
                    {
                        type[cellNum] = "string";
                    }
                    else if (val.equals("float"))
                    {
                        type[cellNum] = "float";
                    }

                    else
                    {
                        System.err.println("配置表错误！字段类型错误：" + val);
                    }
                }
            }
        }

        if (bHF)
            sb.append("namespace hf\n{");
        String hfprefix = bHF ? "HF" : "";
        sb.append("public class " + javaName + " :" + hfprefix + "ConfigModel {\n");
        for (int i = 1; i < len; i++)
        {
            if (fields[i].equals("season_sign"))
                continue;

            if (i >= fields.length)
                break;
            if (fields[i].equals("null"))
            {
                continue;
            }
            if (type[i] == null)
            {
                System.err.println("配置表错误！字段类型为空，第" + i + "列");
                break;
            }
            sb.append("\tpublic " + type[i] + " " + fields[i] + ";//"
                    + zhuShi[i] + "\n");
        }
        sb.append("}");
        if (bHF)
            sb.append("}");
        writeFile(sb.toString(), fileName, true);
    }

    private static void createJava(XSSFSheet xssfSheet, String javaPath)
    {
        String javaName = xssfSheet.getSheetName();
        if (hasPostfix(javaName))
            return;

        String head = javaName.substring(0, 3);
        if (head.equals("HF_"))
            javaName = javaName.substring(3, javaName.length());
        javaName = getFirstWordUp(javaName) + "ConfigModel";
        String fileName = javaPath + javaName + ".java";

        StringBuffer sb = new StringBuffer();
        String[] zhuShi = null;
        String[] fields = null;
        String[] type = null;

        int lastRowNumber = GetSheetLastRowNum(xssfSheet);
        int lastCellNumber = GetSheetLastCellNum(xssfSheet);
        int len = lastCellNumber;

        for (int rowNum = 0; rowNum <= lastRowNumber; rowNum++)
        {
            XSSFRow xssfRow = xssfSheet.getRow(rowNum);

            if (rowNum > 2)
                break;
            if (rowNum == 0)
            {
                if (zhuShi == null)
                {
                    zhuShi = new String[len];
                    fields = new String[len];
                    type = new String[len];
                }
            }
            for (int cellNum = 0; cellNum < lastCellNumber; cellNum++)
            {
                XSSFCell xssfCell = xssfRow.getCell(cellNum);
                if (xssfCell == null)
                    continue;
                String val = xssfCell.getStringCellValue();

                if (rowNum == 0)
                    zhuShi[cellNum] = val;
                else if (rowNum == 1)
                    fields[cellNum] = val;
                else if (rowNum == 2)
                {
                    if (val.equals("int"))
                    {
                        type[cellNum] = "int";
                    }
                    else if (val.equals("string"))
                    {
                        type[cellNum] = "String";
                    }
                    else if (val.equals("float"))
                    {
                        type[cellNum] = "float";
                    }
                }
            }
        }


        sb.append("package com." + strPackageName + ".config.configModel;\n");
        sb.append("import com." + strPackageName + ".config.ConfigModel;\n");
        sb.append("public class " + javaName + " extends ConfigModel {\n");
        for (int i = 1; i < len; i++)
        {
            if (fields[i].equals("season_sign"))
                continue;

            if (i >= type.length)
                break;
            if (type[i] == null)
            {
                break;
            }
            if (fields[i].equals("null"))
            {
                break;
            }
            sb.append("\tpublic " + type[i] + " " + fields[i] + ";//"
                    + zhuShi[i] + "\n");
        }
        sb.append("}");
        writeFile(sb.toString(), fileName, false);
    }

    private static boolean hasPostfix(String javaName)
    {
        if (javaName.length() >= 3
                && javaName.charAt(javaName.length() - 3) == '_'
                && javaName.charAt(javaName.length() - 2) == 'S')
            return true;
        return false;
    }

    public static int GetSheetLastRowNum(XSSFSheet xssfSheet)//最后一行行号内部索引  0-xxx
    {
        int lastRowNum = -1;
        int rr = xssfSheet.getLastRowNum();

        for (int rowNum = 0; rowNum <= xssfSheet.getLastRowNum(); rowNum++)
        {
            XSSFRow curRow = xssfSheet.getRow(rowNum);
            if (curRow == null)
                break;
            XSSFCell cell0 = curRow.getCell(0);
            if (cell0 == null)
                break;
            cell0.setCellType(Cell.CELL_TYPE_STRING);
            if (cell0.getStringCellValue().equals(""))
                break;
            lastRowNum = rowNum;
        }
        return lastRowNum;
    }

    public static int GetSheetLastCellNum(XSSFSheet xssfSheet)//列的条目数
    {
        int lastCellNum = 0;

        if (xssfSheet.getLastRowNum() >= 0)
        {
            XSSFRow curRow = xssfSheet.getRow(0);
            int cc = curRow.getLastCellNum();
            for (int cellNum = 0; cellNum < curRow.getLastCellNum(); cellNum++)
            {
                XSSFCell c = curRow.getCell(cellNum);
                if (c == null)
                    break;
                c.setCellType(Cell.CELL_TYPE_STRING);
                if (c.getStringCellValue().equals(""))
                {
                    break;
                }
                lastCellNum++;
            }
        }
        return lastCellNum;
    }


    // private static boolean isNum(String val) {
    // String reg = "\\d+\\.{0,1}\\d*";
    // // String reg = "^[-+]?(([0-9]+)([.]([0-9]+))?|([.]([0-9]+))?)$";
    // return val.matches(reg);
    // }

    public static void createXml(String path, XSSFSheet xssfSheet)
    {
        String sheetName = xssfSheet.getSheetName();
        sheetName = path + getFirstWordUp(sheetName);
        StringBuffer sb = new StringBuffer();

        int lastRowNum = GetSheetLastRowNum(xssfSheet);
        int lastCellNum = GetSheetLastCellNum(xssfSheet);
        String[] fields = null;
        String[] type = null;

        if (lastCellNum != 3)
            return;

        HashMap<String, ArrayList<Pair<String, String>>> mapTextFile = new HashMap<>();

        for (int rowNum = 1; rowNum <= lastRowNum; rowNum++)
        {
            XSSFRow xssfRow = xssfSheet.getRow(rowNum);
            if (xssfRow == null)//todo 断行问题
                continue;

            XSSFCell xKey = xssfRow.getCell(0);
            XSSFCell xValue = xssfRow.getCell(1);
            XSSFCell xFile = xssfRow.getCell(2);

            if (xKey != null) xKey.setCellType(Cell.CELL_TYPE_STRING);
            if (xValue != null) xValue.setCellType(Cell.CELL_TYPE_STRING);
            if (xFile != null) xFile.setCellType(Cell.CELL_TYPE_STRING);

            String key = xKey != null ? xKey.getStringCellValue() : "";
            String value = xValue != null ? xValue.getStringCellValue() : "";
            String file = xFile != null ? xFile.getStringCellValue() : "";

            try
            {
                if (!key.equals(""))
                {
                    int k = 999;
                    k = Integer.parseInt(key);
                    if (k > 0 && k < 100)
                        key = String.format("%03d", k);
                }
            }
            catch (Exception e)
            {

            }

            if (file.endsWith(".csv"))
                file = file.replace(".csv", "");

            if (file.equals("") || key.equals(""))
                continue;

            if (!mapTextFile.containsKey(file))
                mapTextFile.put(file, new ArrayList<Pair<String, String>>());

            mapTextFile.get(file).add(new Pair<>(key, value));
        }

        String folderPath = sheetName;
        File f = new File(folderPath);
        if (!f.exists())
            f.mkdir();

        for (String file : mapTextFile.keySet())
        {
            ArrayList<Pair<String, String>> arr = mapTextFile.get(file);
            String s = "";
            for (Pair p : arr)
            {
                s += p.getKey();
                s += "\t";
                s += p.getValue();
                s += "\n";
            }
            System.err.println(file);
            //writeFile(s, file + ".txt", false);
            writeFile(s, folderPath + "/" + file + ".txt", false);
        }


    }

    private static int getSeasonCellNum(XSSFSheet xssfSheet, int lastRowNum, int lastCellNum)
    {
        //找到赛季字段列的num
        int seasonCellNum = -1;
        for (int rowNum = 0; rowNum <= lastRowNum; rowNum++)
        {
            if (rowNum == 1)
            {
                XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                if (xssfRow == null)
                    break;
                for (int cellNum = 0; cellNum < lastCellNum; cellNum++)
                {
                    XSSFCell xssfCell = xssfRow.getCell(cellNum);
                    if (xssfCell == null)
                        continue;
                    xssfCell.setCellType(Cell.CELL_TYPE_STRING);
                    String val = xssfCell.getStringCellValue();
                    if (val.equals("season_sign"))
                    {
                        seasonCellNum = cellNum;
                        break;
                    }
                }
                break;
            }
        }
        return seasonCellNum;
    }


    public static void writeFile(String data, String filePath, boolean flag)
    {
        try
        {
            File file = new File(filePath);

            if (file.exists() && !file.isDirectory())
                file.delete();
            file.createNewFile();

            final byte[] bom = new byte[]{(byte) 0xEF, (byte) 0xBB,
                    (byte) 0xBF};

            FileOutputStream fos = new FileOutputStream(file, true);

            if (flag)
                fos.write(bom);
            OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
            osw.write(data);
            osw.flush();
            osw.close();

            //System.out.println("创建" + file.getName() + "完成！文件位置："+ file.getAbsolutePath());

        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
    }

    public static void writeHFCsharpFile(String data, String filePath, boolean flag)
    {
        try
        {
            File file = new File(filePath);

            if (!file.exists())
            {
                file.createNewFile();
            }

            final byte[] bom = new byte[]{(byte) 0xEF, (byte) 0xBB,
                    (byte) 0xBF};

            FileOutputStream fos = new FileOutputStream(file, true);
            if (flag)
                fos.write(bom);
            OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
            osw.write(data);
            osw.flush();
            osw.close();

            //System.out.println("创建" + file.getName() + "完成！文件位置："+ file.getAbsolutePath());

        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
    }

    /**
     * 首字母大写
     */
    public static String getFirstWordUp(String str)
    {
        return str.replaceFirst(str.substring(0, 1), str.substring(0, 1)
                .toUpperCase());
    }

}