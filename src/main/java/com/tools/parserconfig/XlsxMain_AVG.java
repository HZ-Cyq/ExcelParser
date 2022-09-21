package com.tools.parserconfig;

import cn.hutool.core.io.FileUtil;
import com.mobi.log.GameLog;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.tools.JavaCompiler;
import javax.tools.ToolProvider;
import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLClassLoader;
import java.sql.SQLOutput;
import java.util.*;
import java.util.Map.Entry;
import java.util.function.IntFunction;
import java.util.stream.Collectors;

public class XlsxMain_AVG {

    static String strPackageName = "mobi";
    static String strExName = "Ex";
    static ArrayList<String> listAvgScripts = new ArrayList<>();
    static ArrayList<String> listAvgScriptsDesc = new ArrayList<>();
    static int rowOffset = 0;
    static String rowIgnoreMask = "说明";

    /**
     * key:表名 如 "Character"
     * value:路径 如 "excel/010_英雄表.xlsx"
     */
    static Map<String, String> sheetNames = new HashMap<>();
    static ArrayList<String> sheetExNames = new ArrayList<>();   // 需要在代码生成器之外单独处理的表格
    static boolean onlyAvg = false;
    static String avgScriptPath = "C:\\Users\\DELL\\Documents\\MOBI\\01_source\\mobi_client\\mobi_client\\mobi_config";
    static String avgScriptMacroFileDirctory = "excel/AVGScripts/Macro/";
    static Macro _currentMacro = null;
    static int startSign = 0;
    static int endSign = 0;

    static boolean check = true;

    ///宏：用于存储指令宏
    public static class Macro {
        String name = "";
        ArrayList<String> macroContentList = new ArrayList<String>();

        public int GetRowCounts() {
            return macroContentList.size();
        }

        public void SetName(String name) {
            this.name = name;
        }

        public String GetName() {
            return name;
        }

        public void AddMacroRow(String macroRow) {
            macroContentList.add(macroRow);
        }

        public String GetMacroRow(int index) {
            return macroContentList.get(index);
        }
    }


    public static void main(String[] args) throws IOException {
        //args = new String["aaa","","","AllAvg"];
        for (int i = 0; i < args.length; i++) {
            System.out.println("传入参数[" + i + "] : " + args[i]);
        }

        if (args.length > 0) {
            strPackageName = args[0];
        }
        if (args.length > 2) {
            strPackageName = args[0];
            onlyAvg = args[1].equals("onlyAvg");
            avgScriptPath = args[2];
        }
        String gentype = ""; // AllAvg:导出所有的AVG    AllConfig:导出所有的Config
        if (args.length > 3) {
            gentype = args[3];
        }

        if (args.length > 4) {
            check = "check".equals(args[4]);
        }

        if (gentype.equals("AllAvg") || gentype.equals("AllConfig")) {
            System.out.println("只导出" + gentype);
            GenAllType(gentype);
            return;
        }

        try {
            if (!onlyAvg) {
                File xml = new File("xml/");
                File cshap = new File("csharp/");
                File csharpHF = new File("csharp/hf/");
                File java = new File("java/");

                if (true) {
                    File[] files = xml.listFiles();
                    for (File f : files) {
                        if (!f.isDirectory())
                            f.delete();
                    }
                    files = cshap.listFiles();
                    for (File f : files) {
                        if (!f.isDirectory())
                            f.delete();
                    }
                    files = csharpHF.listFiles();
                    for (File f : files) {
                        f.delete();
                    }

                    files = java.listFiles();
                    for (File f : files) {
                        f.delete();
                    }
                }

                {
                    File file = new File("excel/");
                    if (file.isDirectory()) {
                        File[] flist = file.listFiles();
                        for (File fi : flist) {
                            if (!CheckIfValidExcel(fi))
                                continue;

                            readXlsx(fi, "xml/", "csharp/", "java/", true, true, true, false, true);
                        }
                    }
                }

                //AVG脚本
                {
                    System.out.println(" ");
                    System.out.println(" ");
                    System.out.println(" ");
                    System.err.println("--开始解析AVG脚本--");
                    File macroFile = new File(avgScriptMacroFileDirctory);
                    if (macroFile.isDirectory()) {
                        File[] flist = macroFile.listFiles();
                        for (File fi : flist) {
                            if (!CheckIfValidExcel(fi))
                                continue;
                            LoadMacroXlsx(fi);
                        }
                    }

                    File file = new File("excel/AVGScripts/");
                    if (file.isDirectory()) {
                        ArrayList<File> flist = traverFile(file);
                        for (File fi : flist) {
                            if (!CheckIfValidExcel(fi))
                                continue;
                            if (fi.getPath().contains("macro") || fi.getPath().contains("Macro"))
                                continue;
                            System.out.println(fi.getPath());
                            readXlsx(fi, "xml/AVGScripts/", "csharp/", "java/", true, false, false, true, false);

                        }
                    }

                    StringBuffer sb = new StringBuffer();
                    sb.append("<root>\n");
                    for (int i = 0; i < listAvgScripts.size(); i++) {
                        sb.append(String.format("<AVGScriptList id=\"%d\" scriptName=\"%s\" description=\"%s\" " +
                                        "fileName=\"%s\"/>\n ", i, listAvgScripts.get(i), listAvgScriptsDesc.get(i),
                                sheetNames.get(listAvgScripts.get(i))));
                    }
                    sb.append("</root>");
                    writeFile(sb.toString(), "xml/AVGScriptList.xml", true);
                }
            } else {
                System.out.println(" ");
                System.out.println(" ");
                System.out.println(" ");
                System.err.println("--开始解析AVG脚本--");
                File file = new File(avgScriptPath);
                if (!CheckIfValidExcel(file))
                    return;
                File macroFile = new File(avgScriptMacroFileDirctory);
                if (macroFile.isDirectory()) {
                    File[] flist = macroFile.listFiles();
                    for (File fi : flist) {
                        if (!CheckIfValidExcel(fi))
                            continue;
                        LoadMacroXlsx(fi);
                    }
                }
                readXlsx(file, "xml/AVGScripts/", "csharp/", "java/", true, false, false, true, false);
            }

            if(check) {
                check();
            }
            System.out.println(" ");
            System.out.println(" ");


        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void check() throws MalformedURLException {

        // 之前的步骤，已经生成好了ConfigManager，xml和modelConfig
        // 在不改动原有逻辑的情况下，把这些文件复制一份到指定的文件夹中。
        String projectRootPath = System.getProperty("user.dir");
        String separator = File.separator;
        String javaRootPath = String.join(separator,projectRootPath, projectRootPath, "src", "main" ,"java");

        // 复制ConfigManager
        String configManagerSrcPath = String.join(separator, projectRootPath, "ConfigManager.java");
        String configManagerDescPath = String.join(separator,projectRootPath, "excel_check", "temporary");
        FileUtil.copy(configManagerSrcPath, configManagerDescPath, true);

        // 复制modelConfig
        String modelConfigSrcPath = String.join(separator, projectRootPath, "java");
        String modelConfigDescPath = String.join(separator,projectRootPath, "excel_check", "temporary");
        FileUtil.copy(modelConfigSrcPath, modelConfigDescPath, true);

        // 复制xml文件
        String xmlSrcPath = String.join(separator, projectRootPath, "xml");
        String xmlDescPath = String.join(separator, projectRootPath, "data");
        File srcFile = new File(xmlSrcPath);
        File descFile = new File(xmlDescPath);
        FileUtil.copyFilesFromDir(srcFile, descFile, true);

        // 找到excel_check文件夹下所有的java文件
        File excelCheckFile = new File(String.join(separator, projectRootPath, "excel_check"));
        FileFilter acceptJavaFile = pathname -> pathname.getName().endsWith(".java");
        List<File> javaFiles = FileUtil.loopFiles(excelCheckFile, acceptJavaFile);
        String[] needCompileFile = javaFiles.stream().map(javaFile -> FileUtil.subPath(projectRootPath, javaFile.getAbsolutePath())).toArray(String[]::new);
        String[] arguments = new String[needCompileFile.length + 2];
        arguments[0] = "-d";
        arguments[1] = "excel_check/class";
        System.arraycopy(needCompileFile, 0, arguments, 2, needCompileFile.length);
        JavaCompiler compiler = ToolProvider.getSystemJavaCompiler();
        // 动态编译ConfigManager
        int run = compiler.run(null, null, null, arguments);
        System.out.println("compiler result： " + run);

        // 执行ConfigManager的Init方法
        try {
            URL url = new URL("file://" + String.join(separator, excelCheckFile.getPath(), "class") + separator);
            URLClassLoader urlClassLoader = URLClassLoader.newInstance(new URL[]{url});
            Class<?> configManagerClass = urlClassLoader.loadClass("com.mobi.config.ConfigManager");
            Method init = configManagerClass.getDeclaredMethod("Init");
            init.invoke(null);
        } catch (MalformedURLException | ClassNotFoundException | InvocationTargetException | NoSuchMethodException |
                 IllegalAccessException e) {
            GameLog.LogError(e.getMessage(), e);
            throw new RuntimeException(e);
        }
//        File configManagerFile = new File(configManagerDescPath);
//        int run = javac.run(null, null, null, "-g", "-verbose", configManagerDescPath);
//        URLClassLoader classLoader = (URLClassLoader) ClassLoader.getSystemClassLoader();
//        try {
//            Method add = URLClassLoader.class.getDeclaredMethod("addURL", new Class[]{URL.class});
//            add.setAccessible(true);
//            add.invoke(classLoader, new Object[]{packageFile.toURI().toURL()});
//            Class c = classLoader.loadClass("ConfigManager");
//            Object o = c.newInstance();
//            Method m = c.getDeclaredMethod("Init");
//            m.invoke(o,null);
//        } catch (NoSuchMethodException | MalformedURLException | InvocationTargetException | IllegalAccessException |
//                 ClassNotFoundException | InstantiationException e) {
//            System.out.println(e);
//            throw new RuntimeException(e);
//        }
//
        // 根据限定语句检查参数正确性
    }
    private static void GenAllType(String type) {
        // AllAvg:导出所有的AVG    AllConfig:导出所有的Config

        try {
            File xml = new File("xml/");
            File cshap = new File("csharp/");
            File csharpHF = new File("csharp/hf/");
            File java = new File("java/");

            if (true) {
                File[] files = xml.listFiles();
                for (File f : files) {
                    if (!f.isDirectory())
                        f.delete();
                }
                files = cshap.listFiles();
                for (File f : files) {
                    if (!f.isDirectory())
                        f.delete();
                }
                files = csharpHF.listFiles();
                for (File f : files) {
                    f.delete();
                }

                files = java.listFiles();
                for (File f : files) {
                    f.delete();
                }
            }


            if (type.equals("AllAvg")) {
                System.err.println("--开始解析AVG脚本--");
                File macroFile = new File(avgScriptMacroFileDirctory);
                if (macroFile.isDirectory()) {
                    File[] flist = macroFile.listFiles();
                    for (File fi : flist) {
                        if (!CheckIfValidExcel(fi))
                            continue;
                        LoadMacroXlsx(fi);
                    }
                }

                File file = new File("excel/AVGScripts/");
                if (file.isDirectory()) {
                    ArrayList<File> flist = traverFile(file);
                    for (File fi : flist) {
                        if (!CheckIfValidExcel(fi))
                            continue;
                        if (fi.getPath().contains("macro") || fi.getPath().contains("Macro"))
                            continue;
                        System.out.println(fi.getPath());
                        readXlsx(fi, "xml/AVGScripts/", "csharp/", "java/", true, false, false, true, false);
                    }
                }
            } else if (type.equals("AllConfig")) {
                {
                    File file = new File("excel/");
                    if (file.isDirectory()) {
                        File[] flist = file.listFiles();
                        for (File fi : flist) {
                            if (!CheckIfValidExcel(fi))
                                continue;

                            readXlsx(fi, "xml/", "csharp/", "java/", true, true, true, false, true);
                        }
                    }
                }
            }

            StringBuffer sb = new StringBuffer();
            sb.append("<root>\n");
            for (int i = 0; i < listAvgScripts.size(); i++) {
                sb.append(String.format("<AVGScriptList id=\"%d\" scriptName=\"%s\" description=\"%s\" " +
                                "fileName=\"%s\"/>\n ", i, listAvgScripts.get(i), listAvgScriptsDesc.get(i),
                        sheetNames.get(listAvgScripts.get(i))));
            }
            sb.append("</root>");
            writeFile(sb.toString(), "xml/AVGScriptList.xml", true);


        } catch (Exception e){
            e.printStackTrace();
        }
    }

    private static boolean CheckIfValidExcel(File fi) {
        boolean isValid = true;
        if (!fi.getName().endsWith(".xlsx") && !fi.getName().endsWith(".xlsm")) {
            isValid = false;
        }
        if (fi.getName().startsWith("____"))
            isValid = false;
        else if (fi.isDirectory()) {
            //System.err.println("不解析文件夹");
            isValid = false;
        } else if (!fi.getName().endsWith(".xlsx") && !fi.getName().endsWith(".xlsm")) {
            //System.err.println("不解析非excel文件");
            isValid = false;
        } else if (fi.getName().contains("~$")) {
            //System.err.println("不解析excel缓存: " + fi.getName());
            isValid = false;
        }
        return isValid;
    }

    /*
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
                        System.out.println("读表：" + fi.getName());
                    readXlsx(fi, xmlPath, csPath, javaPath, true, true, true, false);
                    System.out.println(" ");
                }
            }


        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
    */

    public static void readXlsx(File file, String xmlPath, String csPath,
                                String javaPath, boolean bCreateXml, boolean bCreateJava, boolean bCreateCSharp,
                                boolean bAVGScript, boolean bJavaConfigManager) throws Exception {
        System.out.println("读表：" + file.getName());
        Map<String, ArrayList<XSSFSheet>> configMap = new HashMap<>();

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(file.getPath());
        // 循环工作表Sheet
        for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
            if (xssfSheet == null) {
                continue;
            }
            String fileNameString = GetSheetNameWithOutDesc(xssfSheet);

            if (!fileNameString.matches("^[A-Z_1234567890a-z\\(\\)]+$")) {
                System.err.println("    x  " + fileNameString);
                continue;
            }
            if (fileNameString.contains("Sheet")) {
                //System.err.println("拒绝解析未改名Sheet！" + fileNameString);
                continue;
            }

            fileNameString = ChangeFirstCharacterToUpperCase(fileNameString);

            if (sheetNames == null) sheetNames = new HashMap<>();
            if (sheetNames.containsKey(fileNameString)) {
                System.err.println("------------> sheet名重复 :" + fileNameString + " in (" + file.getPath() + ") in (" + sheetNames.get(fileNameString) + ")");
            } else {
                sheetNames.put(fileNameString, file.getPath());
            }
            boolean isExSheet = CheckSheetNameContainsEx(xssfSheet);
            if (sheetExNames == null) sheetExNames = new ArrayList<>();
            if (isExSheet && !sheetExNames.contains(fileNameString)) {
                sheetExNames.add(fileNameString);
            }

            if (bAVGScript == false) {
                String mainName = GetMainSheetName(xssfSheet);
                mainName = ChangeFirstCharacterToUpperCase(mainName);
                if (configMap.containsKey(mainName)) {
                    configMap.get(mainName).add(xssfSheet);
                } else {
                    configMap.put(mainName, new ArrayList<XSSFSheet>());
                    configMap.get(mainName).add(xssfSheet);
                }
                continue;
            } else {
                System.err.println("  解析 " + fileNameString);
                CheckRowOffset(xssfSheet);

                //if (bCreateXml)
                //    createXml_AVG(xmlPath, xssfSheet);
                //if (bCreateJava)
                //    createJava(xssfSheet, javaPath);
                //if (bCreateCSharp)
                //    createCSharp(xssfSheet, csPath);

                if (bCreateXml)
                    createXml_AVG(xmlPath, xssfSheet);

                listAvgScripts.add(ChangeFirstCharacterToUpperCase(GetSheetNameWithOutDesc(xssfSheet)));
                listAvgScriptsDesc.add(GetSheetDesc(xssfSheet));
            }
        }

        for (String name : configMap.keySet()) {
            ArrayList<XSSFSheet> sheets = configMap.get(name);
            StringBuilder sb = new StringBuilder();
            sb.append("  解析 " + name + " [ ");
            for (int i = 0; i < sheets.size(); i++) {
                XSSFSheet sheet = sheets.get(i);
                sb.append(sheet.getSheetName());
                if (i != sheets.size() - 1) sb.append(" , ");
            }
            sb.append(" ]");

            System.err.println(sb);
            if (bCreateXml)
                createXml(xmlPath, name, sheets);
            if (bCreateJava)
                createJava(sheets.get(0), javaPath);
            if (isSheetNameCreateCSharp(sheets.get(0)))
                createCSharp(sheets.get(0), csPath);
        }

        if (bJavaConfigManager) {
            CreateJavaConfigManager("./");
        }
        System.out.println(" ");

    }

    //private static String ChangeFirstCharacterToUpperCase(String fileNameString)
    //{
    //    fileNameString = fileNameString.substring(0,1).toUpperCase() + fileNameString.substring(1);
    //    return fileNameString;
    //}

    private static String GetSheetNameWithOutDesc(XSSFSheet xssfSheet) {
        String fileNameString = xssfSheet.getSheetName();
        String[] arr = fileNameString.split("\\|");
        fileNameString = arr[0];
        return fileNameString;
    }

    private static boolean CheckSheetNameContainsEx(XSSFSheet xssfSheet) {
        String fileNameString = xssfSheet.getSheetName();
        String[] arr = fileNameString.split("\\|");
        if (arr.length > 1) {
            fileNameString = arr[1];
            if (fileNameString.contains("++")) {
                return true;
            }
        }
        return false;
    }

    //这个字段是否需要生成C#代码
    private static boolean isFieldCreateCSharp(String stringCellValue) {
        String[] arr = stringCellValue.split("\\|");
        if (arr.length > 1) {
            if (arr[arr.length - 1].contains("-c")) {
                return false;
            }
        }
        return true;
    }

    //这个页签是否生成C#代码
    private static boolean isSheetNameCreateCSharp(XSSFSheet xssfSheet) {
        String fileNameString = xssfSheet.getSheetName();
        String[] arr = fileNameString.split("\\|");
        if (arr.length > 1) {
            String clientOrServer = arr[arr.length - 1];
            if (clientOrServer.contains("-c")) {
                return false;
            }
        }
        return true;
    }

    private static String GetMainSheetName(XSSFSheet sheet) {
        String name = sheet.getSheetName();
        String[] arr = name.split("\\|");
        name = arr[0];
        arr = name.split("\\(");
        name = arr[0];
        return name;
    }

    private static String GetSheetDesc(XSSFSheet xssfSheet) {
        String Desc = xssfSheet.getSheetName();
        String[] arr = Desc.split("\\|");
        if (arr.length > 1) {
            return arr[1];
        } else
            return "";
    }


    private static void createCSharp(XSSFSheet xssfSheet, String cshrpPath) {
        String javaName = GetSheetNameWithOutDesc(xssfSheet);
        if (hasPostfix(javaName))
            return;

        String head = javaName.substring(0, 3);
        if (head.equals("HF_")) {
            javaName = javaName.substring(3, javaName.length());
            javaName = ChangeFirstCharacterToUpperCase(javaName) + "ConfigModel";
            javaName = "HF_" + javaName;
        } else
            javaName = ChangeFirstCharacterToUpperCase(javaName) + "ConfigModel";


        String fileName = cshrpPath + javaName + ".cs";
        boolean bHF = false;
        if (javaName.contains("HF_")) {
            bHF = true;
            fileName = cshrpPath + "hf/" + javaName + ".cs";
        }


        StringBuffer sb = new StringBuffer();
        String[] comment = null;
        String[] fields = null;
        String[] type = null;
        int len = 0;
        int lastRowNumber = GetSheetLastRowNum(xssfSheet);
        int lastCellNumber = GetSheetLastCellNum(xssfSheet);
        for (int rowNum = 0; rowNum <= lastRowNumber; rowNum++) {
            XSSFRow xssfRow = GetRow(xssfSheet, rowNum);
            len = lastCellNumber;
            if (rowNum > 2)
                break;
            if (rowNum == 0) {
                if (comment == null) {
                    comment = new String[len];
                    fields = new String[len];
                    type = new String[len];
                }
            }
            for (int cellNum = 0; cellNum < lastCellNumber; cellNum++) {
                XSSFCell xssfCell = xssfRow.getCell(cellNum);
                if (xssfCell == null) {
                    if (rowNum == 2) {
//						type[cellNum] = "string";
                    }
                    continue;
                }
                if (rowNum == 0) {
                    String s = xssfCell.getStringCellValue();
                    s = s.replaceAll("\r", "");
                    s = s.replaceAll("\n", "");
                    comment[cellNum] = s;
                } else if (rowNum == 1) {
                    fields[cellNum] = xssfCell.getStringCellValue();
                } else if (rowNum == 2) {
                    String val = xssfCell.getStringCellValue();
                    if (val.equals("int")) {
                        type[cellNum] = "int";
                    } else if (val.equals("string")) {
                        type[cellNum] = "string";
                    } else if (val.equals("float")) {
                        type[cellNum] = "float";
                    } else if (val.equals("astring")) {
                        type[cellNum] = "System.Collections.Generic.List<string>";
                    } else if (val.equals("aint")) {
                        type[cellNum] = "System.Collections.Generic.List<int>";
                    } else if (val.equals("afloat")) {
                        type[cellNum] = "System.Collections.Generic.List<float>";
                    } else if (val.equals("aastring")) {
                        type[cellNum] = "System.Collections.Generic.List<System.Collections.Generic.List<string>>";
                    } else if (val.equals("aaint")) {
                        type[cellNum] = "System.Collections.Generic.List<System.Collections.Generic.List<int>>";
                    } else if (val.equals("aafloat")) {
                        type[cellNum] = "System.Collections.Generic.List<System.Collections.Generic.List<float>>";
                    } else if (val.equals("desc")) {
                        type[cellNum] = "desc";
                    } else {
                        System.err.println("配置表错误！a字段类型错误：" + val);
                    }
                }
            }
        }

        if (bHF)
            sb.append("namespace hf\n{");
        String hfprefix = bHF ? "HF" : "";
        sb.append("public class " + javaName + " :" + hfprefix + "ConfigModel {\n");
        boolean isFirstField = true;
        for (int i = 0; i < len; i++) {

            if (i >= fields.length || i >= type.length)
                break;
            if (type[i] == null) {
                System.err.println("配置表错误！字段类型为空，第" + i + "列");
                break;
            }

            if (!isFieldCreateCSharp(fields[i]))
                continue;
            if (fields[i].equals("season_sign"))
                continue;
            if (fields[i].equals("null"))
                continue;
            if (type[i].equals("desc"))
                continue;
            if (isFirstField) {
                isFirstField = false;
                continue;
            }

            sb.append("\tpublic " + type[i] + " " + fields[i] + ";//"
                    + comment[i] + "\n");
        }
        sb.append("}");
        if (bHF)
            sb.append("}");
        writeFile(sb.toString(), fileName, true);
    }


    private static void createJava(XSSFSheet xssfSheet, String javaPath) {
        String javaName = GetSheetNameWithOutDesc(xssfSheet);
        if (hasPostfix(javaName))
            return;

        String head = javaName.substring(0, 3);
        if (head.equals("HF_"))
            javaName = javaName.substring(3, javaName.length());
        javaName = ChangeFirstCharacterToUpperCase(javaName) + "ConfigModel";
        String fileName = javaPath + javaName + ".java";

        StringBuffer sb = new StringBuffer();
        String[] comment = null;
        String[] fields = null;
        String[] type = null;

        int lastRowNumber = GetSheetLastRowNum(xssfSheet);
        int lastCellNumber = GetSheetLastCellNum(xssfSheet);
        int len = lastCellNumber;

        for (int rowNum = 0; rowNum <= lastRowNumber; rowNum++) {
            XSSFRow xssfRow = GetRow(xssfSheet, rowNum);

            if (rowNum > 2)
                break;
            if (rowNum == 0) {
                if (comment == null) {
                    comment = new String[len];
                    fields = new String[len];
                    type = new String[len];
                }
            }
            for (int cellNum = 0; cellNum < lastCellNumber; cellNum++) {
                XSSFCell xssfCell = xssfRow.getCell(cellNum);
                if (xssfCell == null)
                    continue;
                String val = xssfCell.getStringCellValue();

                if (rowNum == 0) {
                    val = val.replaceAll("\r", "");
                    val = val.replaceAll("\n", "");
                    comment[cellNum] = val;
                } else if (rowNum == 1) {
                    fields[cellNum] = getOriginField(val);
                } else if (rowNum == 2) {
                    if (val.equals("int")) {
                        type[cellNum] = "int";
                    } else if (val.equals("string")) {
                        type[cellNum] = "String";
                    } else if (val.equals("float")) {
                        type[cellNum] = "float";
                    } else if (val.equals("desc")) {
                        type[cellNum] = "desc";
                    } else if (val.equals("aint")) {
                        type[cellNum] = "ArrInt";
                    } else if (val.equals("afloat")) {
                        type[cellNum] = "ArrFloat";
                    } else if (val.equals("astring")) {
                        type[cellNum] = "ArrString";
                    } else if (val.equals("aaint")) {
                        type[cellNum] = "ArrArrInt";
                    } else if (val.equals("aafloat")) {
                        type[cellNum] = "ArrArrFloat";
                    } else if (val.equals("aastring")) {
                        type[cellNum] = "ArrArrString";
                    }

                }
            }
        }


        sb.append("package com." + strPackageName + ".config.configModel;\n");
        sb.append("import com." + strPackageName + ".config.ConfigModel;\n");
        sb.append("import com." + strPackageName + ".config.wrappedArrayList.*;\n");
        sb.append("public class " + javaName + " extends ConfigModel {\n");

        boolean isFirstField = true;
        for (int i = 0; i < len; i++) {
            if (i >= fields.length || i >= type.length || type[i] == null || fields[i].equals("null"))
                break;

            if (fields[i].equals("season_sign"))
                continue;

            if (type[i].equals("desc"))
                continue;

            if (isFirstField) {
                isFirstField = false;
                continue;
            }

            sb.append("\tpublic " + type[i] + " " + fields[i] + ";//"
                    + comment[i] + "\n");
        /*
            sb.append("\tprivate " + type[i] + " " + fields[i] + ";//"
                    + comment[i] + "\n");
        }

        sb.append("\n");

        isFirstField = true;
        for (int i = 0; i < len; i++)
        {
            if (i >= fields.length || i >= type.length || type[i] == null || fields[i].equals("null"))
                break;

            if (fields[i].equals("season_sign"))
                continue;

            if (type[i].equals("desc"))
                continue;

            if(fields[i].length() == 0)
                continue;

            if(isFirstField){
                isFirstField =false;
                continue;
            }
            String indexKey = "" + fields[i].charAt(0);
            String indexPostfix = "";
            if(fields[i].length() >= 1) {
                indexPostfix = fields[i].substring(1);
            }
            sb.append("\tpublic " + type[i] + " get" + indexKey.toUpperCase() + indexPostfix + "() { return " +
            fields[i] + "; }\n");*/
        }
        sb.append("}");
        writeFile(sb.toString(), fileName, false);
    }


    private static boolean hasPostfix(String javaName) {
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

        for (int rowNum = 0; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
            XSSFRow curRow = GetRow(xssfSheet, rowNum);
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

        if (xssfSheet.getLastRowNum() >= 0) {
            XSSFRow curRow = GetRow(xssfSheet, 0);
            int cc = curRow.getLastCellNum();
            for (int cellNum = 0; cellNum < curRow.getLastCellNum(); cellNum++) {
                XSSFCell c = curRow.getCell(cellNum);
                if (c == null)
                    break;
                c.setCellType(Cell.CELL_TYPE_STRING);
                if (c.getStringCellValue().equals("")) {
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
    static int idOffset = 0;

    public static void createXml_AVG(String path, XSSFSheet xssfSheet) {
        String sheetName = GetSheetNameWithOutDesc(xssfSheet);
        String head = sheetName.substring(0, 3);
        if (head.equals("HF_"))
            sheetName = sheetName.substring(3, sheetName.length());
        sheetName = path + ChangeFirstCharacterToUpperCase(sheetName);
        StringBuffer sb = new StringBuffer();
        StringBuffer macroSB = new StringBuffer();
        boolean isMacro = false;
        idOffset = 0;


        int lastRowNum = GetSheetLastRowNum(xssfSheet);
        int lastCellNum = GetSheetLastCellNum(xssfSheet);
        String[] fields = null;
        String[] type = null;

        //初始化/填满 fields , type 数组
        for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
            XSSFRow xssfRow = GetRow(xssfSheet, rowNum);
            if (xssfRow == null)
                continue;
            if (rowNum == 0)
                continue;

            if (rowNum == 1) {
                fields = new String[lastCellNum];
                for (int cellNum = 0; cellNum < lastCellNum; cellNum++) {
                    XSSFCell xssfCell = xssfRow.getCell(cellNum);
                    if (xssfCell == null)
                        continue;
                    xssfCell.setCellType(Cell.CELL_TYPE_STRING);
                    String val = xssfCell.getStringCellValue();
                    if (val.equals(""))
                        System.err.println("配置表错误！！！！" + sheetName);
                    else
                        fields[cellNum] = val;
                }
            } else if (rowNum == 2) {
                type = new String[lastCellNum];
                for (int cellNum = 0; cellNum < lastCellNum; cellNum++) {
                    XSSFCell xssfCell = xssfRow.getCell(cellNum);
                    if (xssfCell == null)
                        continue;
                    xssfCell.setCellType(Cell.CELL_TYPE_STRING);
                    String val = xssfCell.getStringCellValue();
                    type[cellNum] = val;
                }
            } else if (rowNum == 3)
                break;
        }

        //找到赛季字段列的num
        int seasonCellNum = getSeasonCellNum(xssfSheet, lastRowNum, lastCellNum);

        //用赛季列的num做key , 构造map , value为所有的行
        Map<String, ArrayList<Integer>> RowNumMapBySeason = new HashMap<String, ArrayList<Integer>>();
        if (seasonCellNum < 0) {
            //如果没有赛季列 ,  则默认为1赛季构造map
            for (int rowNum = 3; rowNum <= lastRowNum; rowNum++) {
                if (!RowNumMapBySeason.containsKey("1"))
                    RowNumMapBySeason.put("1", new ArrayList<Integer>());
                RowNumMapBySeason.get("1").add(rowNum);
            }
        } else {
            for (int rowNum = 3; rowNum <= lastRowNum; rowNum++) {
                XSSFRow xssfRow = GetRow(xssfSheet, rowNum);
                if (xssfRow == null)
                    continue;
                XSSFCell xssfCell = xssfRow.getCell(seasonCellNum);
                if (xssfCell == null)
                    continue;

                xssfCell.setCellType(Cell.CELL_TYPE_STRING);
                for (char c : xssfCell.getStringCellValue().toCharArray()) {
                    if (!RowNumMapBySeason.containsKey(String.valueOf(c)))
                        RowNumMapBySeason.put(String.valueOf(c), new ArrayList<Integer>());
                    RowNumMapBySeason.get(String.valueOf(c)).add(rowNum);
                }
            }
        }

        //创建所有赛季的xml
        for (Entry<String, ArrayList<Integer>> kv : RowNumMapBySeason.entrySet()) {
            String postFix = kv.getKey();
            sb.setLength(0);
            sb.append("<root>\n");
            for (Integer rowNum : kv.getValue()) {
                XSSFRow xssfRow = GetRow(xssfSheet, rowNum);
                if (xssfRow == null)
                    continue;
                String originalRow = "";

//                idOffset = rowNum - 4;
                for (int c = 0; c < lastCellNum; c++) {
                    XSSFCell cell = xssfRow.getCell(c);
                    if (cell == null) continue;
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    if (fields[c].equals("id")) {
                        for (int i = 0; i < lastCellNum; i++) {
                            if (!fields[i].equals("originalRow"))
                                originalRow = cell.getStringCellValue();
                        }
                    }
                    if ("action".equals(fields[c])) {
                        XSSFCell action_cell = xssfRow.getCell(c);
                        if (action_cell != null) {
                            String action_value = action_cell.getStringCellValue();
                            if (action_value.equals("指令宏")) {

                                for (int i = 0; i < lastCellNum; i++) {
                                    XSSFCell n_cell = xssfRow.getCell(i);
                                    if (n_cell == null) continue;
//                            n_cell.setCellType(Cell.CELL_TYPE_STRING);
                                    if (fields[i].equals("data")) {
                                        XSSFCell data_cell = xssfRow.getCell(i);
                                        if (data_cell == null) continue;
                                        String value = data_cell.getStringCellValue();
                                        if (value.startsWith("macro_")) {
                                            Macro macro = macroMap.get(value);
                                            for (int j = 0; j < macro.GetRowCounts(); j++) {
                                                String macroRow = macro.GetMacroRow(j);
//                                if (i != 0)
                                                macroSB.append(ChangeMacroId(originalRow, idOffset + j + rowNum - 3,
                                                        GetSheetNameWithOutDesc(xssfSheet), macroRow));
                                            }
                                            idOffset = idOffset + macro.GetRowCounts() - 1;
                                            isMacro = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //文本宏
//                    cell.setCellValue(ReplaceSimpleMacroWithInstruction(cell));
                }


                if (!isMacro) {
                    sb.append("<" + GetSheetNameWithOutDesc(xssfSheet));
                    for (int cellNum = 0; cellNum < lastCellNum; cellNum++) {
                        if (cellNum == seasonCellNum)
                            continue;
                        XSSFCell xssfCell = xssfRow.getCell(cellNum);
                        if (xssfCell == null)
                            continue;
                        xssfCell.setCellType(Cell.CELL_TYPE_STRING);

                        if (fields.length <= cellNum) {
//						System.err.println(GetSheetNameWithOutDesc(xssfSheet) + "有空列。。."
//								+ (cellNum + 1));
                            continue;
                        }
                        if ("null".equals(fields[cellNum]))
                            continue;
                        if ("desc".equals(type[cellNum])) //注释列
                            continue;

                        String str = xssfCell.getStringCellValue();
                        if (type[cellNum].equals("int")) {
                            if (str.equals(""))
                                str = "0";
                            else {
                                double d = Double.parseDouble(str);
                                Integer integer = (int) (d + 0.5);
                            }
                        } else if (type[cellNum].equals("float")) {
                            if (str.equals(""))
                                str = "0";
                            else
                                str = (new Float(Float.parseFloat(str))).toString();
                        }

                        //转换< > & "
                        str = ReplaceString(str);

                        if (fields[cellNum].equals("id")) {
                            originalRow = str;
                            str = String.valueOf(idOffset + Integer.parseInt(str));
                            boolean needAddRowCounst = true;
                            for (int i = 0; i < lastCellNum; i++) {
                                if (fields[i].equals("originalRow"))
                                    needAddRowCounst = false;
                            }
                            if (needAddRowCounst)
                                sb.append(" " + fields[cellNum] + "=\"" + str + "\"");
                            sb.append(" originalRow" + "=\"" + originalRow + "\"");
                        } else
                            sb.append(" " + fields[cellNum] + "=\"" + str + "\"");
                    }
                    sb.append("/>\n");
                } else if (isMacro) {
                    sb.append(macroSB);

                    macroSB.delete(0, macroSB.length());
                    isMacro = false;
                }
            }
            sb.append("</root>");


            if (postFix.equals("1"))
                writeFile(sb.toString(), sheetName + ".xml", false);
            else
                writeFile(sb.toString(), sheetName + "_S" + postFix + ".xml", false);
        }
    }

    public static void createXml(String path, String mainName, ArrayList<XSSFSheet> sheets) {

        String xmlPath = path + mainName;
        StringBuffer mainSB = new StringBuffer();
        StringBuffer subSB = new StringBuffer();
        int fieldCount = 0;
        String[] fields = null;
        String[] types = null;
        // init fields and types by first sheet
        {
            XSSFSheet sheet = sheets.get(0);
            String sheetName = sheet.getSheetName();
            CheckRowOffset(sheet);
            fieldCount = GetSheetLastCellNum(sheet);
            fields = new String[fieldCount];
            types = new String[fieldCount];
            XSSFRow fieldRow = GetRow(sheet, 1);
            XSSFRow typeRow = GetRow(sheet, 2);
            for (int c = 0; c < fieldCount; c++) {
                XSSFCell fieldCell = fieldRow.getCell(c);
                if (fieldCell != null) {
                    fieldCell.setCellType(Cell.CELL_TYPE_STRING);
                    String value = getOriginField(fieldCell.getStringCellValue());
                    if (value.equals("")) System.err.println("配置表错误！！！！" + sheetName);
                    else fields[c] = value;
                } else {
                    System.err.println("配置表错误！！！！" + sheetName);
                }

                XSSFCell typeCell = typeRow.getCell(c);
                if (typeCell != null) {
                    typeCell.setCellType(Cell.CELL_TYPE_STRING);
                    String value = typeCell.getStringCellValue();
                    types[c] = value;
                } else {
                    System.err.println("配置表错误！！！！" + sheetName);
                }
            }
        }

        ///写入脚本到文件
        mainSB.append("<root>\n");
        for (XSSFSheet sheet : sheets) {
            CheckRowOffset(sheet);
            subSB.setLength(0);
            String sheetName = sheet.getSheetName();
            int rowCount = GetSheetLastRowNum(sheet);
            int colCount = GetSheetLastCellNum(sheet);

            // check field and type
            {
                String errMsg = "--------------- 配置表错误！！！！字段数量不一致 ---------------";
                if (colCount != fieldCount) {
                    System.err.println(errMsg + sheetName);
                    continue;
                }
                boolean fieldError = false;
                XSSFRow fieldRow = GetRow(sheet, 1);
                XSSFRow typeRow = GetRow(sheet, 2);
                for (int c = 0; c < colCount; c++) {
                    errMsg = "--------------- 配置表错误！！！！第" + (c + 1) + "列字段不一致 ---------------";
                    XSSFCell fieldCell = fieldRow.getCell(c);
                    if (fieldCell == null) {
                        System.err.println(errMsg + sheetName);
                        fieldError = true;
                        break;
                    }
                    fieldCell.setCellType(Cell.CELL_TYPE_STRING);
                    String field = getOriginField(fieldCell.getStringCellValue());
                    if (!field.equals(fields[c])) {
                        System.err.println(errMsg + sheetName);
                        fieldError = true;
                        break;
                    }
                    XSSFCell typeCell = typeRow.getCell(c);
                    if (typeCell == null) {
                        System.err.println(errMsg + sheetName);
                        fieldError = true;
                        break;
                    }
                    typeCell.setCellType(Cell.CELL_TYPE_STRING);
                    String type = typeCell.getStringCellValue();
                    if (!type.equals(types[c])) {
                        System.err.println(errMsg + sheetName);
                        fieldError = true;
                        break;
                    }
                }
                if (fieldError) {
                    continue;
                }
            }

            // 从第四行开始解析
            for (int r = 4; r <= rowCount; r++) {
                XSSFRow row = GetRow(sheet, r);
                if (row == null) continue;
                subSB.append("<" + mainName);
                for (int c = 0; c < colCount; c++) {
                    XSSFCell cell = row.getCell(c);
                    if (cell == null) continue;
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    if ("null".equals(fields[c]))
                        continue;
                    if ("desc".equals(types[c])) //注释列
                        continue;

                    String value = cell.getStringCellValue();
                    if (types[c].equals("int")) {
                        if (value.equals(""))
                            value = "0";
                        else {
                            double d = Double.parseDouble(value);
                            Integer integer = (int) (d + 0.5);
                        }
                    } else if (types[c].equals("float")) {
                        if (value.equals(""))
                            value = "0";
                        else
                            value = (new Float(Float.parseFloat(value))).toString();
                    }

                    //转换< > & "
                    value = ReplaceString(value);

                    subSB.append(" " + fields[c] + "=\"" + value + "\"");
                }
                subSB.append("/>\n");
            }
            mainSB.append(subSB);
        }
        String limitStat = generateLimitStat(mainName, fields, GetRow(sheets.get(0), 3));
        mainSB.append(limitStat);
        mainSB.append("</root>");

        writeFile(mainSB.toString(), xmlPath + ".xml", false);
    }

    private static String getOriginField(String field) {
        if (field.contains("|")) {
            field = field.substring(0, field.indexOf("|"));
        }
        return field;
    }

    public static String generateLimitStat(String mainName, String[] fields, XSSFRow limitStatRow) {
        StringBuffer stringBuffer = new StringBuffer();
        stringBuffer.append("<").append(mainName).append("LimitStat");
        for (int i = 0; i < fields.length; i++) {
            stringBuffer.append(" ").append(fields[i]).append("=\"").append(limitStatRow.getCell(i)).append("\"");
        }
        stringBuffer.append("/>\n");
        return stringBuffer.toString();
    }

    private static int getSeasonCellNum(XSSFSheet xssfSheet, int lastRowNum, int lastCellNum) {
        //找到赛季字段列的num
        int seasonCellNum = -1;
        for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
            if (rowNum == 1) {
                XSSFRow xssfRow = GetRow(xssfSheet, rowNum);
                if (xssfRow == null)
                    break;
                for (int cellNum = 0; cellNum < lastCellNum; cellNum++) {
                    XSSFCell xssfCell = xssfRow.getCell(cellNum);
                    if (xssfCell == null)
                        continue;
                    xssfCell.setCellType(Cell.CELL_TYPE_STRING);
                    String val = xssfCell.getStringCellValue();
                    if (val.equals("season_sign")) {
                        seasonCellNum = cellNum;
                        break;
                    }
                }
                break;
            }
        }
        return seasonCellNum;
    }


    public static void writeFile(String data, String filePath, boolean flag) {
        try {
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

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void writeHFCsharpFile(String data, String filePath, boolean flag) {
        try {
            File file = new File(filePath);

            if (!file.exists()) {
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

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 首字母大写
     */
    public static String ChangeFirstCharacterToUpperCase(String str) {
        return str.replaceFirst(str.substring(0, 1), str.substring(0, 1)
                .toUpperCase());
    }

    // 检查配置表的行偏移
    private static void CheckRowOffset(XSSFSheet xssfSheet) {
        String sheetName = GetSheetNameWithOutDesc(xssfSheet);
//        System.out.println(" 配置表：" + sheetName + "  开始检查行偏移。");
        rowOffset = 0;
        int rowIndex = 0;
        int maxRow = xssfSheet.getLastRowNum();
        if (maxRow <= 0) {
            System.err.println(" 配置表：" + sheetName + "  行偏移配置错误");
            return;
        }
        while (rowIndex < maxRow) {
            XSSFRow row = xssfSheet.getRow(rowIndex);
            if (row == null) break;
            XSSFCell cell = row.getCell(0);
            if (cell == null) break;
            cell.setCellType(Cell.CELL_TYPE_STRING);
            String val = cell.getStringCellValue();
            if (val.equals("")) {
                System.err.println(" 配置表：" + sheetName + "  行偏移配置错误");
            } else if (val.equals(rowIgnoreMask)) {
                // ignore row
            } else { //val 应为 “ID”
                rowOffset = rowIndex;
//                System.out.println(" 配置表：" + sheetName + "  行偏移为：" + rowOffset);
                return;
            }
            rowIndex++;
        }
    }

    private static XSSFRow GetRow(XSSFSheet xssfSheet, int rowNum) {
        XSSFRow xssfRow = xssfSheet.getRow(rowNum + rowOffset);
        return xssfRow;
    }

    private static void CreateJavaConfigManager(String filePath) {
        String str = "";
        str += "//该文件由导表工具生成，不用手动修改\n";
        str += "//该文件由导表工具生成，不用手动修改\n";
        str += String.format("package com.%s.config;\n", strPackageName);
        str += String.format("import com.%s.config.configModel.*;\n", strPackageName);
        str += String.format("import com.%s.config.modelParser.*;\n", strPackageName);
        str += "import java.util.ArrayList;\n" +
                "import java.util.List;\n";
        str += "\n";
        str += "public class ConfigManager\n" +
                "{\n";
        str += "\tpublic static final List<String> refreshConfigList = new ArrayList<>();\n";
        for (String fileName : sheetNames.keySet()) {
            if (fileName.contains("(") || fileName.contains(")"))
                continue;
            String fileNameEx = "";
            if (sheetExNames.contains(fileName)) {
                fileNameEx += strExName;
            }
            str += String.format("\tpublic static Config<%sConfigModel%s> %sConfigModels = new " +
                    "Config<%sConfigModel%s>();\n", fileName, fileNameEx, fileName, fileName, fileNameEx);
        }
        str += "\n";

        str += "\tpublic static boolean Init()\n" +
                "\t{\n" +
                "\t\tboolean isOk = true;\n";

        for (String fileName : sheetNames.keySet()) {
            if (fileName.contains("(") || fileName.contains(")"))
                continue;
            String fileNameEx = "";
            if (sheetExNames.contains(fileName)) {
                fileNameEx += strExName;
            }
            str += String.format("\t\tisOk = isOk && %sConfigModels.Init(%sConfigModel%s.class);\n", fileName,
                    fileName, fileNameEx);
        }

        str += "\t\tExBuild();\n" +
                "\t\tSystem.out.println(\"hello world\");\n" +
                "\t\treturn isOk;\n" +
                "\t}\n" +
                "\n";
        str += "\tpublic static boolean ReInit()\n" +
                "\t{\n" +
                "\t\tboolean isOk = true;\n";
        for (String fileName : sheetNames.keySet()) {
            if (fileName.contains("(") || fileName.contains(")"))
                continue;
            String fileNameEx = "";
            if (sheetExNames.contains(fileName)) {
                fileNameEx += strExName;
            }
            str += String.format("\t\tisOk = isOk && %sConfigModels.ReInit(%sConfigModel%s.class);\n", fileName,
                    fileName, fileNameEx);
        }

        str += "\t\tExBuild();\n" +
                "\t\treturn isOk;\n" +
                "\t}\n" +
                "\n";

        str += "\tprivate static void ExBuild()\n" +
                "\t{\n";

        for (String fileName : sheetNames.keySet()) {
            if (fileName.contains("(") || fileName.contains(")"))
                continue;
            if (!sheetExNames.contains(fileName))
                continue;
            String fileNameEx = strExName;
            str += String.format("\t\tfor(%sConfigModel%s c : %sConfigModels.dict.values()) c.build();\n", fileName,
                    fileNameEx, fileName);
        }

        str += "\n";


        for (String fileName : sheetNames.keySet()) {
            if (fileName.contains("(") || fileName.contains(")"))
                continue;
            if (!sheetExNames.contains(fileName))
                continue;
            String fileNameEx = strExName;
            str += String.format("\t\tfor(%sConfigModel%s c : %sConfigModels.dict.values()) c.check();\n", fileName,
                    fileNameEx, fileName);
        }

        str += "\t}\n" +
                "\n" +
                "\tpublic static void Clear()\n" +
                "\t{\n";

        for (String fileName : sheetNames.keySet()) {
            if (fileName.contains("(") || fileName.contains(")"))
                continue;
            str += String.format("\t\t%sConfigModels.Clear();\n", fileName);
        }

        str += "\t}\n" +
                "}\n";


        String javaName = "ConfigManager";
        String fileName = filePath + javaName + ".java";
        writeFile(str, fileName, false);

    }


    ///avg指令宏相关
    public static HashMap<String, String> _simpleMacroMap = new HashMap<>();

    public static HashMap<String, Macro> macroMap = new HashMap<>();

    private static void LoadMacroXlsx(File file) {
        System.out.println(" ");
        System.out.println("读宏命令表：" + file.getName());
        Map<String, ArrayList<XSSFSheet>> configMap = new HashMap<>();

        XSSFWorkbook xssfWorkbook = null;
        try {
            xssfWorkbook = new XSSFWorkbook(file.getPath());
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 循环工作表Sheet
        for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
            if (xssfSheet == null) {
                continue;
            }
            String fileNameString = GetSheetNameWithOutDesc(xssfSheet);

            if (!fileNameString.matches("^[A-Z_1234567890a-z\\(\\)]+$")) {
                System.err.println("    x  " + fileNameString);
                continue;
            }
            if (fileNameString.contains("Sheet")) {
                //System.err.println("拒绝解析未改名Sheet！" + fileNameString);
                continue;
            }

            if (sheetNames == null) sheetNames = new HashMap<>();
            if (sheetNames.containsKey(fileNameString)) {
                System.err.println("------------> sheet名重复 :" + fileNameString + " in (" + file.getPath() + ") in (" + sheetNames.get(fileNameString) + ")");
            } else {
                sheetNames.put(fileNameString, file.getPath());
            }

            String mainName = GetMainSheetName(xssfSheet);
            if (configMap.containsKey(mainName)) {
                configMap.get(mainName).add(xssfSheet);
            } else {
                configMap.put(mainName, new ArrayList<XSSFSheet>());
                configMap.get(mainName).add(xssfSheet);
            }
            continue;

        }


        for (String name : configMap.keySet()) {
            ArrayList<XSSFSheet> sheets = configMap.get(name);
            StringBuilder sb = new StringBuilder();
            sb.append("  解析 " + name + " [ ");
            for (int i = 0; i < sheets.size(); i++) {
                XSSFSheet sheet = sheets.get(i);
                sb.append(sheet.getSheetName());
                if (i != sheets.size() - 1) sb.append(" , ");
            }
            sb.append(" ]");

            System.err.println(sb);
            readMacroXlsx(name, sheets);
        }

    }

    private static void readMacroXlsx(String mainName, ArrayList<XSSFSheet> sheets) {
        StringBuffer mainSB = new StringBuffer();
        StringBuffer subSB = new StringBuffer();
        StringBuffer macroSB = new StringBuffer();
        StringBuffer macroSubSB = new StringBuffer();
        int fieldCount = 0;
        String[] fields = null;
        String[] types = null;
        // init fields and types by first sheet
        {
            XSSFSheet sheet = sheets.get(0);
            String sheetName = sheet.getSheetName();
            CheckRowOffset(sheet);
            fieldCount = GetSheetLastCellNum(sheet);
            fields = new String[fieldCount];
            types = new String[fieldCount];
            XSSFRow fieldRow = GetRow(sheet, 1);
            XSSFRow typeRow = GetRow(sheet, 2);
            for (int c = 0; c < fieldCount; c++) {
                XSSFCell fieldCell = fieldRow.getCell(c);
                if (fieldCell != null) {
                    fieldCell.setCellType(Cell.CELL_TYPE_STRING);
                    String value = fieldCell.getStringCellValue();
                    if (value.equals("")) System.err.println("配置表错误！！！！" + sheetName);
                    else fields[c] = value;
                } else {
                    System.err.println("配置表错误！！！！" + sheetName);
                }

                XSSFCell typeCell = typeRow.getCell(c);
                if (typeCell != null) {
                    typeCell.setCellType(Cell.CELL_TYPE_STRING);
                    String value = typeCell.getStringCellValue();
                    types[c] = value;
                } else {
                    System.err.println("配置表错误！！！！" + sheetName);
                }
            }
        }

        ///写入脚本到文件


//        mainSB.append("<root>\n");
        for (XSSFSheet sheet : sheets) {
            CheckRowOffset(sheet);
            subSB.setLength(0);
            String sheetName = sheet.getSheetName();
            int rowCount = GetSheetLastRowNum(sheet);
            int colCount = GetSheetLastCellNum(sheet);

            // check field and type
            {
                String errMsg = "--------------- 配置表错误！！！！字段数量不一致 ---------------";
                if (colCount != fieldCount) {
                    System.err.println(errMsg + sheetName);
                    continue;
                }
                boolean fieldError = false;
                XSSFRow fieldRow = GetRow(sheet, 1);
                XSSFRow typeRow = GetRow(sheet, 2);

                for (int c = 0; c < colCount; c++) {
                    errMsg = "--------------- 配置表错误！！！！第" + (c + 1) + "列字段不一致 ---------------";
                    XSSFCell fieldCell = fieldRow.getCell(c);
                    if (fieldCell == null) {
                        System.err.println(errMsg + sheetName);
                        fieldError = true;
                        break;
                    }
                    fieldCell.setCellType(Cell.CELL_TYPE_STRING);
                    String field = fieldCell.getStringCellValue();
                    if (!field.equals(fields[c])) {
                        System.err.println(errMsg + sheetName);
                        fieldError = true;
                        break;
                    }
                    XSSFCell typeCell = typeRow.getCell(c);
                    if (typeCell == null) {
                        System.err.println(errMsg + sheetName);
                        fieldError = true;
                        break;
                    }
                    typeCell.setCellType(Cell.CELL_TYPE_STRING);
                    String type = typeCell.getStringCellValue();
                    if (!type.equals(types[c])) {
                        System.err.println(errMsg + sheetName);
                        fieldError = true;
                        break;
                    }
                }
                if (fieldError) {
                    continue;
                }
            }

            PreLoadSimpleMacro(fields, sheet);
            startSign = 0;
            endSign = rowCount;

            for (int r = 3; r <= rowCount; r++) {
                XSSFRow row = GetRow(sheet, r);
                if (row == null) continue;

                if (macroSubSB.length() > 0)
                    macroSubSB.delete(0, macroSubSB.length());
                subSB.append("<" + mainName);
                macroSubSB.append("<" + mainName);
                for (int c = 0; c < colCount; c++) {
                    XSSFCell cell = row.getCell(c);
                    if (cell == null)
                        continue;
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    if ("null".equals(fields[c]))
                        continue;
                    if ("desc".equals(types[c])) //注释列
                        continue;

                    String value = cell.getStringCellValue();
                    if (types[c].equals("int")) {
                        if (value.equals(""))
                            value = "0";
                        else {
                            double d = Double.parseDouble(value);
                            Integer integer = (int) (d + 0.5);
                        }
                    } else if (types[c].equals("float")) {
                        if (value.equals(""))
                            value = "0";
                        else
                            value = (new Float(Float.parseFloat(value))).toString();
                    }

//                    转换< > & "
                    if (fields[c].equals("param"))
                        value = ReplaceSimpleMacroWithInstruction(cell);
                    value = ReplaceString(value);
                    subSB.append(" " + fields[c] + "=\"" + value + "\"");
                    macroSubSB.append(" " + fields[c] + "=\"" + value + "\"");
                }
                subSB.append("/>\n");
                macroSubSB.append("/>\n");

                PreLoadMacro(macroSubSB, fields, sheetName, rowCount, colCount, r, row);

            }
//            mainSB.append(subSB);
        }
//        mainSB.append("</root>");

//        writeFile(mainSB.toString(),xmlPath+".xml",false);

    }

    private static boolean PreLoadMacro(StringBuffer macroSubSB, String[] fields, String sheetName, int rowCount,
                                        int colCount, int r, XSSFRow row) {

        String macroName = "";
        for (int i = 0; i < colCount; i++) {
            XSSFCell cell = row.getCell(i);
            if (cell == null) continue;
            cell.setCellType(Cell.CELL_TYPE_STRING);
            if ("null".equals(fields[i]))
                continue;
            if ("groupSign".equals(fields[i])) {
                String value = cell.getStringCellValue();

                if (value.equals("宏开始")) {
//                            System.out.println("宏开始");
                    startSign = r;
                    _currentMacro = new Macro();
                    for (int j = 0; j < colCount; j++) {
                        if ("data".equals(fields[j])) {
                            XSSFCell cell_data = row.getCell(j);
                            String value_data = cell_data.getStringCellValue();
                            _currentMacro.name = value_data;
                        }
                    }
                } else if (value.equals("宏结束")) {
//                            System.out.println("宏结束");
                    endSign = r;
                }
            }
        }


        if (startSign > endSign) {
//                    System.out.println("--------------- 宏命令配置表错误！！！！第"+(r -3 )+"行开始命令后没有结束命令 ---------------" +
//                    sheetName);
//                    System.out.println("startSign   :   = " + startSign);
//                    System.out.println("endSign   :   = " + endSign);
        } else {
            if (r > startSign && r < endSign && _currentMacro != null) {
                _currentMacro.AddMacroRow(macroSubSB.toString());
            }

            if (r == endSign && startSign != endSign && _currentMacro != null) {
                if (macroMap.containsKey(_currentMacro.name)) {
                    System.err.println("--------------- 宏命令配置表错误！！！！" + macroName + "宏命名重复---------------" + sheetName);
                    return true;
                }
                macroMap.put(_currentMacro.name, _currentMacro);
                endSign = rowCount;
                startSign = r;
                _currentMacro = null;
            }
        }
        return false;
    }

    private static void PreLoadSimpleMacro(String[] fields, XSSFSheet sheet) {
        String macroName = "";
        String simpleMacroName = "";
        String simpleMacroContent = "";
        int rowCount = GetSheetLastRowNum(sheet);
        int colCount = GetSheetLastCellNum(sheet);
        String sheetName = sheet.getSheetName();

        //预读文本宏
        for (int r = 3; r <= rowCount; r++) {
            XSSFRow row = GetRow(sheet, r);
            if (row == null) continue;

//            if (macroSubSB.length()>1)
//                macroSubSB.delete(0, macroSubSB.length()-1);
//            subSB.append("<" + mainName);
//            macroSubSB.append("<" + mainName);

            for (int i = 0; i < colCount; i++) {
                XSSFCell cell = row.getCell(i);
                if (cell == null) continue;
                cell.setCellType(Cell.CELL_TYPE_STRING);
                if ("null".equals(fields[i]))
                    continue;
                if ("groupSign".equals(fields[i])) {
                    String value = cell.getStringCellValue();
                    if (value.equals("文本宏")) {
                        for (int j = 0; j < colCount; j++) {
                            if ("data".equals(fields[j])) {
                                XSSFCell cell_data = row.getCell(j);
                                String value_data = cell_data.getStringCellValue();
                                simpleMacroName = value_data;
                            }
                            if ("param".equals(fields[j])) {
                                XSSFCell cell_data = row.getCell(j);
                                String value_data = cell_data.getStringCellValue();
                                simpleMacroContent = value_data;
                            }
                        }
                        if (_simpleMacroMap.containsKey(simpleMacroName))
                            System.err.println("--------------- 宏命令配置表错误！！！！" + simpleMacroName +
                                    "宏命名重复---------------" + sheetName);

                        _simpleMacroMap.put(simpleMacroName, simpleMacroContent);
                    }
                }
            }
        }
    }

    private static String ReplaceSimpleMacroWithInstruction(XSSFCell cell) {
        String value = "";
        StringBuffer finalValueSB = new StringBuffer();
        String param_value = cell.getStringCellValue();
        String[] stringArray = param_value.split("\\{macro=");
        finalValueSB.append(stringArray[0]);
        if (stringArray.length <= 1) ;
        else {
            for (int i = 1; i < stringArray.length; i++) {
                String[] ss = stringArray[i].split("}");
                String textMacroName = ss[0];
                String macro = _simpleMacroMap.get(textMacroName);
//                                macro = ReplaceString(macro);
                finalValueSB.append(macro);
                if (ss.length >= 2)
                    finalValueSB.append(ss[1]);
            }
        }
        value = finalValueSB.toString();
//        value = ReplaceString(value);
        return value;
    }

    private static String ReplaceString(String macro) {
        macro = macro.replace("&", "&amp;");
        macro = macro.replace("<", "&lt;");
        macro = macro.replace(">", "&gt;");
        macro = macro.replace("\"", "&quot;");
        return macro;
    }

    private static String ChangeMacroId(String originalRow, int id, String mainName, String macro) {
        StringBuffer sb = new StringBuffer();
//        String[] sourceString = macro.split("\n");
        String[] ss = macro.split(" ");
        for (int i = 0; i < ss.length; i++) {
            String s = ss[i];
            if (s.startsWith("id=")) {
                String[] newStringArray = s.split("\"");
                s = newStringArray[0] + "\"" + id + "\"" + " " + "originalRow=\"" + originalRow + "\"";
            }
            if (s.startsWith("<macro")) {
                s = "<" + mainName + "";
            }
            if (i != 0)
                sb.append(" " + s);
            else
                sb.append(s);
        }
        return sb.toString();
    }

    private static ArrayList<File> traverFile(File file) {
        LinkedList<File> list = new LinkedList<>();
        ArrayList<File> ret = new ArrayList<>();
        list.push(file);
        while (!list.isEmpty()) {
            File f = list.pop();
            if (!f.isDirectory())
                ret.add(f);
            File[] files = f.listFiles();
            if (files != null) {
                for (File fil : files) {
                    list.push(fil);
                }
            }
        }

        return ret;
    }
}