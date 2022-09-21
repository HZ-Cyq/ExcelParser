public class CharacterConfigModel :ConfigModel {
	public string name;//卡牌名字
	public int character_type;//类别
	public int gender;//性别(1男，2女，3妖怪)
	public int model_id;//模型ID
	public System.Collections.Generic.List<int> skills;//技能组
	public string drawing;//卡牌立绘
	public string hero_icon;//英雄头像
	public System.Collections.Generic.List<System.Collections.Generic.List<int>> attribute_base;//基础属性
}