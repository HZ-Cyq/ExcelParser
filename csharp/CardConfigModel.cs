public class CardConfigModel :ConfigModel {
	public string name;//名字
	public int gender;//性别(1男，2女，3妖怪)
	public int model_id;//战斗模型ID
	public System.Collections.Generic.List<int> skills;//技能组
	public string drawing;//立绘
	public string icon;//头像
	public System.Collections.Generic.List<System.Collections.Generic.List<int>> attribute_base;//基础属性
}