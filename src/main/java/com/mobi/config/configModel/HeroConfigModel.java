package com.mobi.config.configModel;
import com.mobi.config.ConfigModel;
import com.mobi.config.wrappedArrayList.*;
public class HeroConfigModel extends ConfigModel {
	public String name;//名字
	public int gender;//性别(1男，2女，3妖怪)
	public int model_id;//战斗模型ID
	public ArrInt skills;//技能组
	public String drawing;//立绘
	public String icon;//头像
}