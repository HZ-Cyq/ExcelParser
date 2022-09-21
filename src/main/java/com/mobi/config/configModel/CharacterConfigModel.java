package com.mobi.config.configModel;
import com.mobi.config.ConfigModel;
import com.mobi.config.wrappedArrayList.*;
public class CharacterConfigModel extends ConfigModel {
	public String name;//卡牌名字
	public int character_type;//类别
	public int gender;//性别(1男，2女，3妖怪)
	public int model_id;//模型ID
	public ArrInt skills;//技能组
	public String drawing;//卡牌立绘
	public String hero_icon;//英雄头像
	public ArrArrInt attribute_base;//基础属性
}