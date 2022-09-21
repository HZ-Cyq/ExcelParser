package com.mobi.config.configModel;
import com.mobi.config.ConfigModel;
import com.mobi.config.wrappedArrayList.*;
public class MonsterGroupConfigModel extends ConfigModel {
	public int group_gid;//怪组
	public String descs;//描述
	public ArrInt monsters;//卡牌名字
	public ArrInt monster_positions;//站位
	public ArrInt monster_levels;//等级
	public ArrInt monster_skill_lvs;//技能等级
	public ArrInt heroes;//主角
	public ArrInt hero_positions;//站位
	public ArrInt hero_skill_lvs;//主角技能等级
}