#region Creating Synchronized collections
	$Global:PrereqHash = [hashtable]::Synchronized(@{})
#endregion Creating Synchronized collections

# Runspace creation
$PrereqHash.host = $Host
$Prereq_Runspace =[runspacefactory]::CreateRunspace()
$Prereq_Runspace.ApartmentState = "STA"
$Prereq_Runspace.ThreadOptions = "ReuseThread"
$Prereq_Runspace.Open()

# Passing variables
$Prereq_Runspace.SessionStateProxy.SetVariable("PrereqHash",$PrereqHash)
$Prereq_Runspace.SessionStateProxy.SetVariable("VariableHash",$VariableHash)
$Prereq_Runspace.SessionStateProxy.SetVariable("Prereq_Runspace",$Prereq_Runspace)

# Create powershell object which will containt the code we're running in the runspace
$psCmdPrereq = [PowerShell]::Create()

# Add runspace to Powershell object
$psCmdPrereq.Runspace = $Prereq_Runspace

# Code to execute in runspace
[Void]$psCmdPrereq.AddScript({

	#region XAML Prep
		# Load Required Assemblies
		Add-Type –assemblyName PresentationFramework # Required to show the GUI
		Add-Type -AssemblyName System.Windows.Forms # Required to use the folder browser dialog
		
		# Loading XAML code	
		[xml]$xaml = Get-Content -Path "$($VariableHash.FormsDir)\PrereqCheck.xaml"
		
		# Loading in to XML Node reader
		$reader = (New-Object System.Xml.XmlNodeReader $xaml)
		
		# Loading XML Node reader in to $Prereq_Runspace window property to launch later
		$PrereqHash.Window = [Windows.Markup.XamlReader]::Load($reader)
	#endregion XAML Prep

	#region Connecting Controls
		# GUI Elements
		$PrereqHash.GUI_TitleBar = $PrereqHash.Window.FindName("TitleBar")
		$synchash.GUI_BTN_Close = $SyncHash.Window.FindName("BTN_Close")
		$synchash.GUI_IMG_Close = $SyncHash.Window.FindName("IMG_Close")	

		# Banner
		$PrereqHash.GUI_Banner = $PrereqHash.Window.FindName("ImageBar")

		# Ouput label
		$PrereqHash.Lbl_Output = $PrereqHash.Window.FindName("LBL_Output")

		# Prerequisite images
		$PrereqHash.IMG_Prereq_Admin = $PrereqHash.Window.FindName("Prereq_IMG_Admin")
		$PrereqHash.IMG_Prereq_Internet = $PrereqHash.Window.FindName("Prereq_IMG_Internet")
		$PrereqHash.IMG_Prereq_SPOMgmtShell = $PrereqHash.Window.FindName("Prereq_IMG_SPOMgmtShell")
		$PrereqHash.IMG_Prereq_AADModule = $PrereqHash.Window.FindName("Prereq_IMG_AADModule")
		$PrereqHash.IMG_Prereq_MSTeamsModule = $PrereqHash.Window.FindName("Prereq_IMG_MSTeams")
		$PrereqHash.IMG_Prereq_IMG_EXOModule = $PrereqHash.Window.FindName("Prereq_IMG_EXOModule")
		$PrereqHash.IMG_Prereq_IMG_WinRMAuth = $PrereqHash.Window.FindName("Prereq_IMG_WinRMAuth")
	#endregion Connecting Controls
	
	#region Configuring Controls
		# GUI elements
		$PrereqHash.GUI_TitleBar.Text = "Automaton Prerequisites check"
		$PrereqHash.GUI_Banner.Source = "$($VariableHash.ImageDir)\LastMile.jpg"

		# Prerequisite elements
		$PrereqHash.IMG_Prereq_Admin.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
		$PrereqHash.IMG_Prereq_Internet.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
		$PrereqHash.IMG_Prereq_SPOMgmtShell.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
		$PrereqHash.IMG_Prereq_AADModule.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
		$PrereqHash.IMG_Prereq_MSTeamsModule.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
		$PrereqHash.IMG_Prereq_IMG_EXOModule.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
		$PrereqHash.IMG_Prereq_IMG_WinRMAuth.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
	#endregion Configuring Controls

	#region Main
		#region Prereq check execution
			$PrereqTest_Runspace =[runspacefactory]::CreateRunspace()
			$PrereqTest_Runspace.ApartmentState = "STA"
			$PrereqTest_Runspace.ThreadOptions = "ReuseThread"
			$PrereqTest_Runspace.Open()

			# Passing variables
			$PrereqTest_Runspace.SessionStateProxy.SetVariable("PrereqHash",$PrereqHash)
			$PrereqTest_Runspace.SessionStateProxy.SetVariable("VariableHash",$VariableHash)
			$PrereqTest_Runspace.SessionStateProxy.SetVariable("Prereq_Runspace",$Prereq_Runspace)
			$PrereqTest_Runspace.SessionStateProxy.SetVariable("PrereqTest_Runspace",$PrereqTest_Runspace)

			$PowerShell = [PowerShell]::Create().AddScript({

				# Importing module
				Import-Module "$($VariableHash.ModuleDir)M365InventoryAutomaton.psm1" -Force -DisableNameChecking
					
				# Prereq testing
				Prereq-Testing

				$VariableHash.PrereqError = $Error
			})

			$PowerShell.Runspace = $PrereqTest_Runspace
			$data = $PowerShell.BeginInvoke()	
		#endregion Prereq check execution

		#region Control Actions
			# Allow window to be dragged around
			$PrereqHash.Window.FindName('Grid').Add_MouseLeftButtonDown({
				$Prereq_Runspace.Window.DragMove()
			})
		#endregion Control Actions
	#endregion Main

	# Show GUI
		$PrereqHash.Window.ShowDialog() | Out-Null
		$VariableHash.PrereqError = $Error

})

# Invoking GUI
$data = $psCmdPrereq.BeginInvoke()