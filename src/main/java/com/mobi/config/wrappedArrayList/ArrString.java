package com.mobi.config.wrappedArrayList;

import com.mobi.config.ConfigModel;

import java.util.ArrayList;

public class ArrString extends ArrayList<String> {
    // TODO 是否增加抛出异常
    public int getInt(int index) {
        // 目前的异常值设为0
        if (index >= this.size()) {
            return 0;
        }
        String val = this.get(index);
        return val.equals("") ? 0 : ConfigModel.getInt(val);
    }

    // TODO 是否增加抛出异常
    public float getFloat(int index) {
        // 目前的异常值设为0f
        if (index >= this.size()) {
            return 0f;
        }
        String val = this.get(index);
        return val.equals("") ? 0 : ConfigModel.getFloat(val);
    }
}
