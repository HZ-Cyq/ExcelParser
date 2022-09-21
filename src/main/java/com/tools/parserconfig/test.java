package com.tools.parserconfig;


import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.LinkedList;

class A {
    public void doThat() { System.out.println("AA");  }
}

class B extends A {
    public void doThat() { /* don't call super.doThat() */ }
}

class C extends B {
    public void doThat() {
        Magic.exec(A.class, this, "doThat");
    }
}


class Magic {
    public static <Type, ChieldType extends Type> void exec(Class<Type> oneSuperType, ChieldType instance,
                                                            String methodOfParentToExec) {
        try {
            Type type = oneSuperType.newInstance();
            shareVars(oneSuperType, instance, type);
            oneSuperType.getMethod(methodOfParentToExec).invoke(type);
            shareVars(oneSuperType, type, instance);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    private static <Type, SourceType extends Type, TargetType extends Type> void shareVars(Class<Type> clazz,
                                                                                           SourceType source, TargetType target) throws IllegalArgumentException, IllegalAccessException {
        Class<?> loop = clazz;
        do {
            for (Field f : loop.getDeclaredFields()) {
                if (!f.isAccessible()) {
                    f.setAccessible(true);
                }
                f.set(target, f.get(source));
            }
            loop = loop.getSuperclass();
        } while (loop != Object.class);
    }
}



public class test
{

    private static ArrayList<File> traverFile(File file)
    {
        LinkedList<File> list = new LinkedList<>();
        ArrayList<File> ret = new ArrayList<>();
        list.push(file);
        while (!list.isEmpty()) {
            File f = list.pop();
            if(!f.isDirectory())
                ret.add(f);
            File[] files = f.listFiles();
            if (files != null) {
                for (File fil : files)
                {
                    list.push(fil);
                }
            }
        }

        return ret;
    }

    public static void main(String[] args) throws IOException, NoSuchMethodException, InvocationTargetException, IllegalAccessException
    {
//        File file = new File("excel/AVGScripts/");
//        ArrayList<File> list = traverFile(file);
//        for(File f : list)
//        {
//            System.out.println(f.getPath());
//        }
        String[] avgArfs = new String[1];
        avgArfs[0] = "mobi";
        XlsxMain_AVG.main(avgArfs);
    }
}
