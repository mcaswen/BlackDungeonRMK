// Copyright Epic Games, Inc. All Rights Reserved.

using UnrealBuildTool;

public class BlackDungeonRMK : ModuleRules
{
	public BlackDungeonRMK(ReadOnlyTargetRules Target) : base(Target)
	{
		PCHUsage = PCHUsageMode.UseExplicitOrSharedPCHs;

		PublicDependencyModuleNames.AddRange(new string[] {
			"Core",
			"CoreUObject",
			"Engine",
			"InputCore",
			"EnhancedInput",
			"AIModule",
			"StateTreeModule",
			"GameplayStateTreeModule",
			"UMG"
		});

		PrivateDependencyModuleNames.AddRange(new string[] { });

		PublicIncludePaths.AddRange(new string[] {
			"BlackDungeonRMK",
			"BlackDungeonRMK/Variant_Platforming",
			"BlackDungeonRMK/Variant_Combat",
			"BlackDungeonRMK/Variant_Combat/AI",
			"BlackDungeonRMK/Variant_SideScrolling",
			"BlackDungeonRMK/Variant_SideScrolling/Gameplay",
			"BlackDungeonRMK/Variant_SideScrolling/AI"
		});

		// Uncomment if you are using Slate UI
		// PrivateDependencyModuleNames.AddRange(new string[] { "Slate", "SlateCore" });

		// Uncomment if you are using online features
		// PrivateDependencyModuleNames.Add("OnlineSubsystem");

		// To include OnlineSubsystemSteam, add it to the plugins section in your uproject file with the Enabled attribute set to true
	}
}
