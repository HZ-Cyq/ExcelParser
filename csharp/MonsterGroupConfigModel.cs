public class MonsterGroupConfigModel :ConfigModel {
	public int group_gid;//怪组
	public string descs;//描述
	public System.Collections.Generic.List<int> monsters;//卡牌名字
	public System.Collections.Generic.List<int> monster_positions;//站位
	public System.Collections.Generic.List<int> monster_levels;//等级
	public System.Collections.Generic.List<int> monster_skill_lvs;//技能等级
	public System.Collections.Generic.List<int> heroes;//主角
	public System.Collections.Generic.List<int> hero_positions;//站位
	public System.Collections.Generic.List<int> hero_skill_lvs;//主角技能等级
}