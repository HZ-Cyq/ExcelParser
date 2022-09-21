package com.mobi.config;

import com.mobi.config.wrappedArrayList.*;
import com.mobi.log.GameLog;
import org.jdom.Attribute;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.input.SAXBuilder;

import java.io.File;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public abstract class ConfigModel {

    public int id;


    //public static void main(String[] args) {
    //	load(DataConfigEx.class);
    //}

    /**
     * 对于非id作为meta主键的扩展,子类可重写此方法
     *
     * @return 主键
     */
    public int getKey() {
        return id;
    }


    @SuppressWarnings({"rawtypes", "unchecked"})
    public static Map load(File file, Class<? extends ConfigModel> className) {

        try {
            Map items2 = new HashMap();

            SAXBuilder builder = new SAXBuilder();

            Document doc;
            GameLog.LogInfo("加载" + file.getPath());

            doc = builder.build(file);

            Element element = doc.getRootElement();
            List<Element> list = element.getChildren();
            Field[] fields = className.getFields();
            for (Element e : list) {
                ConfigModel model = className.newInstance();
                for (Field f : fields) {
                    String fieldName = f.getName();
                    if (fieldName.startsWith("_")) {
                        continue;
                    }
                    Class<?> type = f.getType();
                    Attribute attr = e.getAttribute(fieldName);
                    String val = "";
                    if (attr != null) {
                        val = attr.getValue();
                    }
                    try {
                        String strTypeSimpleName = type.getSimpleName();
                        if (strTypeSimpleName.equals("int")) {
                            f.set(model, val.equals("") ? 0 : getInt(val));
                        } else if (strTypeSimpleName.equals("String")) {
                            f.set(model, val);
                        } else if (strTypeSimpleName.equals("float")) {
                            f.set(model, val.equals("") ? 0f : getFloat(val));
                        } else if (strTypeSimpleName.equals("ArrInt")) {
                            ArrInt ai = new ArrInt();
                            val = val.trim();
                            if (!val.equals("")) {
                                String[] arr = val.split("[|]");
                                for (String s : arr) {
                                    int a;// default 0
                                    try {
                                        a = Integer.parseInt(s);
                                        ai.add(a);
                                    } catch (Exception ex) {
                                        GameLog.LogError(String.format("解析表%s时 AInt数据错误", className));
                                    }
                                }
                            }
                            f.set(model, ai);
                        } else if (strTypeSimpleName.equals("ArrFloat")) {
                            ArrFloat ai = new ArrFloat();
                            val = val.trim();
                            if (!val.equals("")) {
                                String[] arr = val.split("[|]");
                                for (int n = 0; n < arr.length; n++) {
                                    float a = 0f;
                                    try {
                                        a = Float.parseFloat(arr[n]);
                                        ai.add(a);
                                    } catch (Exception ex) {
                                        GameLog.LogError(String.format("解析表%s时 AFloat数据错误", className));
                                    }
                                }
                            }
                            f.set(model, ai);
                        } else if (strTypeSimpleName.equals("ArrString")) {
                            ArrString ai = new ArrString();
                            String[] arr = val.split("[|]");
                            for (int n = 0; n < arr.length; n++) {
                                String s = arr[n];
                                if (s.endsWith("\n"))
                                    s = s.substring(0, s.length() - 1);
                                if (s.endsWith("\r"))
                                    s = s.substring(0, s.length() - 1);
                                ai.add(s);
                            }
                            f.set(model, ai);
                        } else if (strTypeSimpleName.equals("ArrArrInt")) {
                            if (val.equals("")) {
                                f.set(model, new ArrArrInt());
                            } else {
                                val = val.trim();
                                ArrArrInt lli = new ArrArrInt();
                                String[] as = val.split("[|]");
                                for (String _as : as) {
                                    ArrayList<Integer> li = new ArrayList<Integer>();
                                    String[] ls = _as.split(";");
                                    for (String s : ls) {
                                        s.trim();
                                        if (s.equals(""))
                                            li.add(0);
                                        else {
                                            try {
                                                li.add(Integer.parseInt(s));
                                            } catch (Exception ex) {
                                                GameLog.LogError(String.format("解析表%s时 ArrArrInt数据错误", className));
                                            }
                                        }

                                    }
                                    lli.add(li);
                                }
                                f.set(model, lli);
                            }
                        } else if (strTypeSimpleName.equals("ArrArrFloat")) {
                            if (val.equals("")) {
                                f.set(model, new ArrArrFloat());
                            } else {
                                val = val.trim();
                                ArrArrFloat lli = new ArrArrFloat();
                                String[] as = val.split("[|]");
                                for (String _as : as) {
                                    ArrayList<Float> li = new ArrayList<Float>();
                                    String[] ls = _as.split(";");
                                    for (String s : ls) {
                                        s.trim();
                                        if (s.equals(""))
                                            li.add(0f);
                                        else {
                                            try {
                                                li.add(Float.parseFloat(s));
                                            } catch (Exception ex) {
                                                GameLog.LogError(String.format("解析表%s时 ArrArrFloat数据错误", className));
                                            }
                                        }

                                    }
                                    lli.add(li);
                                }
                                f.set(model, lli);
                            }
                        } else if (strTypeSimpleName.equals("ArrArrString")) {
                            if (val.equals("")) {
                                f.set(model, new ArrArrString());
                            } else {
                                val = val.trim();
                                ArrArrString lli = new ArrArrString();
                                String[] as = val.split("[|]");
                                for (String _as : as) {
                                    ArrString li = new ArrString();
                                    String[] ls = _as.split(";");
                                    for (String s : ls) {
                                        if (s.endsWith("\n"))
                                            s = s.substring(0, s.length() - 1);
                                        if (s.endsWith("\r"))
                                            s = s.substring(0, s.length() - 1);
                                        li.add(s);
                                    }
                                    lli.add(li);
                                }
                                f.set(model, lli);
                            }
                        } else {
                            GameLog.LogError("读表(" + className + ")有未知类型: " + strTypeSimpleName);
                        }
                    } catch (Exception e2) {
                        GameLog.LogError("model:" + className + "|field:"
                                + f.getName() + "|val:" + val);
                        e2.printStackTrace();
                    }
                }
                try {
                    if (ConfigParser.class.isAssignableFrom(className)) {
                        Method method = className.getMethod("parser", new Class[0]);
                        method.invoke(model, new Object[]{});
                    }
                } /*catch (ConfigParseException e1) {
					GameLog.LogError(file.getName() + " data is invalid, index = " + e1.getIndexName() + " data = " +
					e1.getDataValue() );
					// TODO 这里是否直接退出，以便发现问题
				} */ catch (Exception ee) {
                    // System.err.println("======" + model.id);
                    GameLog.LogException("{}", ee);
                    // TODO 这里是否直接退出，以便发现问题
                }
                items2.put(model.getKey(), model);
            }
            return items2;
        } catch (Exception e) {
            //System.err.println("[ERROR]"+loadFile);
            GameLog.LogException("[EX]{}", e);
            return null;
        }
    }

    public boolean build() {
        return true;
    }

    public boolean check() {
        return true;
    }

    public static boolean isNum(String val) {
        String reg = "(-|\\+)?\\d+\\.{0,1}\\d*";
        // String reg = "^[-+]?(([0-9]+)([.]([0-9]+))?|([.]([0-9]+))?)$";
        return val.matches(reg);
    }

    public static int getInt(String val) {
        return Integer.parseInt(val);
    }

    public static float getFloat(String val) {
        return Float.parseFloat(val);
    }


    /*public String toString() {
        try {
            Field[] fields = ItemConfigModel.class.getFields();
            String fName = this.getClass().getSimpleName();
            StringBuffer sb = new StringBuffer(fName);
            for (Field f : fields) {
                if (f.getName().equals("path")) {
                    continue;
                }
                sb.append(" | " + f.getName() + ":" + f.get(this));
            }
            return sb.toString();
        } catch (Exception e) {
            return e.getMessage();
        }
    }*/
    public static class ConfigParseException extends Exception {

        private static final long serialVersionUID = 1L;
        private final String indexName;
        private final String dataValue;


        public ConfigParseException(String cName, String value) {
            indexName = cName;
            dataValue = value;
        }

        public String getIndexName() {
            return indexName;
        }

        public String getDataValue() {
            return dataValue;
        }
    }
}
