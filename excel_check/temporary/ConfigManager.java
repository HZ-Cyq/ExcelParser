//该文件由导表工具生成，不用手动修改
//该文件由导表工具生成，不用手动修改
package com.mobi.config;
import com.mobi.config.configModel.*;
import com.mobi.config.modelParser.*;
import java.util.ArrayList;
import java.util.List;

public class ConfigManager
{
	public static final List<String> refreshConfigList = new ArrayList<>();
	public static Config<CharacterConfigModel> CharacterConfigModels = new Config<CharacterConfigModel>();
	public static Config<MonsterConfigModel> MonsterConfigModels = new Config<MonsterConfigModel>();
	public static Config<HeroConfigModel> HeroConfigModels = new Config<HeroConfigModel>();
	public static Config<MonsterGroupConfigModel> MonsterGroupConfigModels = new Config<MonsterGroupConfigModel>();
	public static Config<CardConfigModel> CardConfigModels = new Config<CardConfigModel>();

	public static boolean Init()
	{
		boolean isOk = true;
		isOk = isOk && CharacterConfigModels.Init(CharacterConfigModel.class);
		isOk = isOk && MonsterConfigModels.Init(MonsterConfigModel.class);
		isOk = isOk && HeroConfigModels.Init(HeroConfigModel.class);
		isOk = isOk && MonsterGroupConfigModels.Init(MonsterGroupConfigModel.class);
		isOk = isOk && CardConfigModels.Init(CardConfigModel.class);
		ExBuild();
		System.out.println("hello world");
		return isOk;
	}

	public static boolean ReInit()
	{
		boolean isOk = true;
		isOk = isOk && CharacterConfigModels.ReInit(CharacterConfigModel.class);
		isOk = isOk && MonsterConfigModels.ReInit(MonsterConfigModel.class);
		isOk = isOk && HeroConfigModels.ReInit(HeroConfigModel.class);
		isOk = isOk && MonsterGroupConfigModels.ReInit(MonsterGroupConfigModel.class);
		isOk = isOk && CardConfigModels.ReInit(CardConfigModel.class);
		ExBuild();
		return isOk;
	}

	private static void ExBuild()
	{

	}

	public static void Clear()
	{
		CharacterConfigModels.Clear();
		MonsterConfigModels.Clear();
		HeroConfigModels.Clear();
		MonsterGroupConfigModels.Clear();
		CardConfigModels.Clear();
	}
}
