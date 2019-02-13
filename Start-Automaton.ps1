#region Variables

	# Declaring Variables
	$Global:VariableHash = [hashtable]::Synchronized(@{})

	# Setting Variables
	$VariableHash.ScriptDir = $PSScriptRoot
	$VariableHash.ResourcesDir = -join($VariableHash.ScriptDir,"\Resources\Forms\")
	$VariableHash.FormsDir = -join($VariableHash.ScriptDir,"\Resources\Forms\")
	$VariableHash.IconsDir = -join($VariableHash.ScriptDir,"\Resources\Icons\")
	$VariableHash.ImageDir = -join($VariableHash.ScriptDir,"\Resources\Images\")
	$VariableHash.ModuleDir = -join($VariableHash.ScriptDir,"\Resources\Modules\")
	$VariableHash.SubscriptsDir = -join($VariableHash.ScriptDir,"\Resources\Scripts\")
	$VariableHash.LibDir = -join($VariableHash.ScriptDir,"\Resources\Lib\")
	$VariableHash.LogFile = -join($VariableHash.ScriptDir,"\$(get-date -Format "ddmmyy-HHMM") - M365 Inventory Automaton.log")
#endregion Variables

# Importing module
Import-Module "$($VariableHash.ModuleDir)\M365InventoryAutomaton.psm1" -Force -DisableNameChecking



.\resources\scripts\Load-Main.ps1