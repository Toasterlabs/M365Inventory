Param(
	[Parameter(mandatory=$false)]
	[SWITCH]$DebugWithGlobalVariables
)

#region Variables
	#region Creating Synchronized collections
		# If debugging is specified, create global parameters
		if($DebugWithGlobalVariables){
			$Global:SyncHash = [hashtable]::Synchronized(@{})
			$Global:VariableHash = [hashtable]::Synchronized(@{})
			$Global:SplashHash = [hashtable]::Synchronized(@{})
		}Else{
			$SyncHash = [hashtable]::Synchronized(@{})
			$VariableHash = [hashtable]::Synchronized(@{})
			$SplashHash = [hashtable]::Synchronized(@{})
		}
	#endregion Creating Synchronized collections

	#region Setting VariableHash
		$VariableHash.ScriptDir = $PSScriptRoot
		$VariableHash.FormsDir = -join($VariableHash.ScriptDir,"\Resources\Forms\")
		$VariableHash.IconsDir = -join($VariableHash.ScriptDir,"\Resources\Icons\")
		$VariableHash.ImageDir = -join($VariableHash.ScriptDir,"\Resources\Images\")
		$VariableHash.ModuleDir = -join($VariableHash.ScriptDir,"\Resources\Modules\")
		$VariableHash.LibDir = -join($VariableHash.ScriptDir,"\Resources\Lib\")
	#endregion Setting VariableHash
#endregion Variables

#region Splash Window
	# Runspace creation
	$SplashHash.host = $Host
	$Splash_Runspace =[runspacefactory]::CreateRunspace()
	$Splash_Runspace.ApartmentState = "STA"
	$Splash_Runspace.ThreadOptions = "ReuseThread"
	$Splash_Runspace.Open()

	# Passing variables
	$Splash_Runspace.SessionStateProxy.SetVariable("SplashHash",$SplashHash)
	$Splash_Runspace.SessionStateProxy.SetVariable("VariableHash",$VariableHash)

	# Create powershell object which will containt the code we're running in the runspace
	$psCmdSplash = [PowerShell]::Create()

	# Add runspace to Powershell object
	$psCmdSplash.Runspace = $Splash_Runspace

		[Void]$psCmdSplash.AddScript({
	
		# Load Required Assemblies
		Add-Type –assemblyName PresentationFramework # Required to show the GUI
			
		# Loading XAML code	
		[xml]$xaml = Get-Content -Path "$($VariableHash.FormsDir)\Splash.xaml"
		
		# Loading in to XML Node reader
		$reader = (New-Object System.Xml.XmlNodeReader $xaml)
		
		# Loading XML Node reader in to $SyncHash window property to launch later
		$SplashHash.Window = [Windows.Markup.XamlReader]::Load($reader)
	
		$SplashHash.GUI_TitleBar = $SplashHash.Window.FindName("TitleBar")
		$SplashHash.GUI_logo = $SplashHash.Window.FindName("Logo")

		$splashHash.GUI_TitleBar.Text = "Loading Automaton..."
		$SplashHash.GUI_Logo.Source = "$($VariableHash.ImageDir)\SogetiLogo.png"

		# Show GUI
		$SplashHash.Window.ShowDialog() | Out-Null
		$VariableHash.Error = $Error
	})

	# Invoking GUI
	$data = $psCmdSplash.BeginInvoke()
#endregion Splash Window

#region Main window
# Runspace creation
	$syncHash.host = $Host
	$Main_Runspace =[runspacefactory]::CreateRunspace()
	$Main_Runspace.ApartmentState = "STA"
	$Main_Runspace.ThreadOptions = "ReuseThread"
	$Main_Runspace.Open()

	# Passing variables
	$Main_Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
	$Main_Runspace.SessionStateProxy.SetVariable("VariableHash",$VariableHash)
	$Main_Runspace.SessionStateProxy.SetVariable("Splashhash",$SplashHash)
	$Main_Runspace.SessionStateProxy.SetVariable("Main_Runspace",$Main_Runspace)

	# Create powershell object which will containt the code we're running in the runspace
	$psCmd = [PowerShell]::Create()

	# Add runspace to Powershell object
	$psCmd.Runspace = $Main_Runspace

	[Void]$psCmd.AddScript({
		#region XAML Prep
			# Load Required Assemblies
			Add-Type –assemblyName PresentationFramework # Required to show the GUI
			Add-Type -AssemblyName System.Windows.Forms # Required to use the folder browser dialog
		
			# Loading XAML code	
			[xml]$xaml = Get-Content -Path "$($VariableHash.FormsDir)\Main.xaml"
		
			# Loading in to XML Node reader
			$reader = (New-Object System.Xml.XmlNodeReader $xaml)
		
			# Loading XML Node reader in to $SyncHash window property to launch later
			$SyncHash.Window = [Windows.Markup.XamlReader]::Load($reader)
		#endregion XAML Prep

		#region Connecting Controls
			#region GUI elements
				$synchash.GUI_BTN_Close = $SyncHash.Window.FindName("BTN_Close")
				$synchash.GUI_IMG_Close = $SyncHash.Window.FindName("IMG_Close")	
				$synchash.GUI_TitleBar = $SyncHash.Window.FindName("TitleBar")
			#endregion GUI elements

			#region menu bar
				#region File
					$synchash.Menu_File_Connect = $SyncHash.Window.FindName("File_Connect")
					$synchash.IMG_Menu_File_Connect = $SyncHash.Window.FindName("File_Connect_IMG")
					$synchash.Menu_File_disConnect = $SyncHash.Window.FindName("File_disconnect")
					$synchash.IMG_Menu_File_disconnect = $SyncHash.Window.FindName("File_disconnect_IMG")
					$synchash.Menu_File_Run = $SyncHash.Window.FindName("File_Run")
					$synchash.IMG_Menu_File_Run = $SyncHash.Window.FindName("File_Run_IMG")
					$synchash.Menu_File_Save = $SyncHash.Window.FindName("File_Save")
					$synchash.IMG_Menu_File_Save = $SyncHash.Window.FindName("File_Save_IMG")
					$synchash.Menu_File_Exit = $SyncHash.Window.FindName("File_Exit")
					$synchash.IMG_Menu_File_Exit = $SyncHash.Window.FindName("File_Exit_IMG")
				#endregion File

				#region help
					$synchash.Menu_Help_About = $SyncHash.Window.FindName("Help_About")
					$synchash.IMG_Menu_Help_About = $SyncHash.Window.FindName("Help_About_IMG")					
				#endregion help			
			#endregion menu bar

			#region Home Tab
				# Wizard image
				$synchash.BTN_Wizard = $SyncHash.Window.FindName("BTN_Wizard")
				$synchash.IMG_Wizard = $SyncHash.Window.FindName("IMG_Wizard")

				# Output/History box
				$synchash.txt_Output = $SyncHash.Window.FindName("TXT_Output")
			
				# Prerequisite images
				$synchash.IMG_Prereq_Admin = $SyncHash.Window.FindName("Prereq_IMG_Admin")
				$synchash.IMG_Prereq_Internet = $SyncHash.Window.FindName("Prereq_IMG_Internet")
				$synchash.IMG_Prereq_SPOMgmtShell = $SyncHash.Window.FindName("Prereq_IMG_SPOMgmtShell")
				$synchash.IMG_Prereq_AADModule = $SyncHash.Window.FindName("Prereq_IMG_AADModule")
				$synchash.IMG_Prereq_MSTeamsModule = $SyncHash.Window.FindName("Prereq_IMG_MSTeams")
				$synchash.Prereq_IMG_CredValidate = $SyncHash.Window.FindName("Prereq_IMG_CredValidate")
				$synchash.Prereq_IMG_Connected = $SyncHash.Window.FindName("Prereq_IMG_Connected")

				# Connections
				$synchash.IMG_Conn_AAD = $SyncHash.Window.FindName("Conn_IMG_AAD")
				$synchash.IMG_Conn_EXO = $SyncHash.Window.FindName("Conn_IMG_EXO")
				$synchash.IMG_Conn_M365 = $SyncHash.Window.FindName("Conn_IMG_M365")
				$synchash.IMG_Conn_SPO = $SyncHash.Window.FindName("Conn_IMG_SPO")

				# Reports
				$synchash.IMG_Report_AllSites = $SyncHash.Window.FindName("IMG_Report_AllSites")
				$synchash.IMG_Report_AllSiteGroups = $SyncHash.Window.FindName("IMG_Report_AllSiteGroups")
				$synchash.IMG_Report_SPOLibrariesgt100 = $SyncHash.Window.FindName("IMG_Report_SPOLibrariesgt100")
				$synchash.IMG_Report_SPOLibrariesgt0 = $SyncHash.Window.FindName("IMG_Report_SPOLibrariesgt0")
				$synchash.IMG_Report_AADUsers = $SyncHash.Window.FindName("IMG_Report_AADUsers")
				$synchash.IMG_Report_Graph = $SyncHash.Window.FindName("IMG_Report_Graph")
				$synchash.IMG_Report_AADDelUser = $SyncHash.Window.FindName("IMG_Report_AADDelUser")
				$synchash.IMG_Report_AADContacts = $SyncHash.Window.FindName("IMG_Report_AADContacts")
				$synchash.IMG_Report_AADGroups = $SyncHash.Window.FindName("IMG_Report_AADGroups")
				$synchash.IMG_Report_AADDomains = $SyncHash.Window.FindName("IMG_Report_AADDomains")
				$synchash.IMG_Report_EXOMailboxes = $SyncHash.Window.FindName("IMG_Report_EXOMailboxes")
				$synchash.IMG_Report_EXOArchives = $SyncHash.Window.FindName("IMG_Report_EXOArchives")
				$synchash.IMG_Report_EXOGroups = $SyncHash.Window.FindName("IMG_Report_EXOGroups")
				$synchash.IMG_Report_VSTS = $SyncHash.Window.FindName("IMG_Report_VSTS")
			#endregion Home Tab

			#region Settings Tab
				# Lastmile
				$synchash.IMG_Settings_LastMile = $SyncHash.Window.FindName("IMG_Settings_LastMile")		

				# Output Directory
				$synchash.IMG_Settings_OutputDir = $SyncHash.Window.FindName("IMG_Settings_OutputDir")
				$synchash.txt_Settings_outputDir = $SyncHash.Window.FindName("txt_Settings_OutputDir")
				$synchash.btn_settings_OutputDir = $SyncHash.Window.FindName("btn_settings_OutputDir")

				# Credentials
				$synchash.IMG_Settings_M365Username = $SyncHash.Window.FindName("IMG_Settings_M365Username")
				$synchash.txt_Settings_M365Username = $SyncHash.Window.FindName("txt_Settings_M365Username")
				$synchash.IMG_Settings_M365Password = $SyncHash.Window.FindName("IMG_Settings_M365Password")
				$synchash.txt_Settings_M365Password = $SyncHash.Window.FindName("txt_Settings_M365Password")
				$synchash.btn_settings_Credentials_Validate = $SyncHash.Window.FindName("btn_settings_Credentials_Validate")

				# Tenant 
				$synchash.IMG_Settings_tenant = $SyncHash.Window.FindName("IMG_Settings_tenant")
				$synchash.txt_Settings_tenant = $SyncHash.Window.FindName("txt_Settings_tenant")

				# Tenant Region
				$synchash.IMG_Settings_tenantRegion = $SyncHash.Window.FindName("IMG_Settings_tenantRegion")
				$synchash.CB_Settings_TenantRegion = $SyncHash.Window.FindName("CB_Settings_TenantRegion")

				# Graph
				$synchash.IMG_Settings_Graph = $SyncHash.Window.FindName("IMG_Settings_Graph")
				$synchash.txt_Settings_GraphAppID = $SyncHash.Window.FindName("txt_Settings_GraphAppID")
				$synchash.IMG_Settings_Redirect = $SyncHash.Window.FindName("IMG_Settings_Redirect")
				$synchash.txt_Settings_GraphRedirectUri = $SyncHash.Window.FindName("txt_Settings_GraphRedirectUri")

			#endregion Settings Tab
		#endregion Connecting Controls

		#region Configuring Elements
			# GUI elements
			$synchash.GUI_IMG_Close.Source = "$($VariableHash.IconsDir)\appbar.Close_white.png"
			$SyncHash.GUI_TitleBar.Text = "M365 Inventory Automaton"
			
			# Menu Elements
			$synchash.IMG_Menu_File_Connect.Source = "$($VariableHash.IconsDir)\appbar.connect.png"
			$synchash.IMG_Menu_File_disconnect.Source = "$($VariableHash.IconsDir)\appbar.disconnect.png"
			$synchash.IMG_Menu_File_Run.Source = "$($VariableHash.IconsDir)\appbar.control.play.ico"
			$synchash.IMG_Menu_File_Save.Source = "$($VariableHash.IconsDir)\appbar.save.ico"
			$synchash.IMG_Menu_File_Exit.Source = "$($VariableHash.IconsDir)\appbar.Close_black.png"
			$synchash.IMG_Menu_Help_About.Source = "$($VariableHash.IconsDir)\appbar.question.png"
			
			# Wizard
			$synchash.IMG_Wizard.Source = "$($VariableHash.ImageDir)\wizard.png"

			# Prerequisite elements
			$synchash.IMG_Prereq_Admin.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Prereq_Internet.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Prereq_SPOMgmtShell.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Prereq_AADModule.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Prereq_MSTeamsModule.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.Prereq_IMG_CredValidate.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.Prereq_IMG_Connected.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"

			# Connections elements
			$synchash.IMG_Conn_AAD.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Conn_EXO.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Conn_M365.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Conn_SPO.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"

			# Settings tab
			$synchash.IMG_Settings_LastMile.Source = "$($VariableHash.ImageDir)\LastMile_flipped.jpg"	
			$synchash.IMG_Settings_OutputDir.Source = "$($VariableHash.IconsDir)\appbar.folder.open.png"
			$synchash.IMG_Settings_M365Username.Source = "$($VariableHash.IconsDir)\appbar.user.tie.png"
			$synchash.IMG_Settings_M365Password.Source = "$($VariableHash.IconsDir)\appbar.interface.password.png"
			$synchash.IMG_Settings_tenant.Source = "$($VariableHash.IconsDir)\appbar.office.365.png"
			$synchash.IMG_Settings_tenantRegion.Source = "$($VariableHash.IconsDir)\appbar.office.365.png"
			$synchash.IMG_Settings_Graph.Source = "$($VariableHash.ImageDir)\graph.png"
			$synchash.txt_Settings_GraphRedirectUri.Source = "$($VariableHash.ImageDir)\graph.png"

			$synchash.CB_Settings_TenantRegion.AddChild("Default")
			$synchash.CB_Settings_TenantRegion.AddChild("Germany")
			$synchash.CB_Settings_TenantRegion.AddChild("China")
			$synchash.CB_Settings_TenantRegion.AddChild("AzurePPE")
			$synchash.CB_Settings_TenantRegion.AddChild("USGovernment")

			# report elements
			$synchash.IMG_Report_AllSites.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_AllSiteGroups.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_SPOLibrariesgt100.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_SPOLibrariesgt0.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			
			$synchash.IMG_Report_Graph.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			
			$synchash.IMG_Report_AADUsers.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_AADDelUser.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_AADContacts.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_AADGroups.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_AADDomains.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			
			$synchash.IMG_Report_EXOMailboxes.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_EXOArchives.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
			$synchash.IMG_Report_EXOGroups.Source = "$($VariableHash.IconsDir)\Check_Waiting.ico"
		#endregion Configuring GUI Elements

		#region Importing Modules
			# Main Module
			Import-Module "$($VariableHash.ModuleDir)M365InventoryAutomaton.psm1" -Force -DisableNameChecking
		#endregion Importing Modules

		#region Control actions
			#region Wizard
				$synchash.BTN_Wizard.Add_Click({
									
					Run-Reports					
					
				})
			#endregion Wizard

			#region Window Controls
				# Allow window to be dragged around
				$Synchash.Window.FindName('Grid').Add_MouseLeftButtonDown({
					$Synchash.Window.DragMove()
				})
			
				# Close Action
				$SyncHash.GUI_BTN_Close.Add_Click({
					Close-Window
				})
			#endregion Window Controls

			#region Menu actions
				# File Connect
				$synchash.Menu_File_Connect.Add_Click({

					If($VariableHash.M365Username){

						# getting tenant name
						$VariableHash.tenantname = $SyncHash.txt_Settings_tenant.text

						# getting tenant region
						$VariableHash.TenantRegion = $SyncHash.CB_Settings_TenantRegion.CurrentItem

						If($VariableHash.TenantRegion -like ""){
							$VariableHash.TenantRegion = "Default"
						}

						# Creating credentials object
						$secpswd = ConvertTo-SecureString $VariableHash.M365password -AsPlainText -Force
						$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $VariableHash.M365Username,$secpswd

						# Calling connection function
						Connect-M365 -Credentials $credentials -Region $VariableHash.TenantRegion -TenantName $VariableHash.tenantname -ExchangeOnline -AzureAD -MSOL -SharePointOnline
						
					}Else{
						$Source = -join($VariableHash.IconsDir,"\appbar.warning.png")
						$Image = New-Object System.Windows.Controls.Image
						$Image.Source = $Source
						$Image.Height = [System.Drawing.Image]::FromFile($Source).Height
						$Image.Width = [System.Drawing.Image]::FromFile($Source).Width
						$Image.Margin = 5
					 
						$TextBlock = New-Object System.Windows.Controls.TextBlock
						$TextBlock.Text = "Please validate credentials before starting inventory!"
						$TextBlock.Padding = 10
						$TextBlock.FontFamily = "Verdana"
						$TextBlock.FontSize = 16
						$TextBlock.TextWrapping = "Wrap"
						$TextBlock.Width = 350
									
						$StackPanel = New-Object System.Windows.Controls.StackPanel
						$StackPanel.Orientation = "Horizontal"
						$StackPanel.Width = 400
						$StackPanel.AddChild($Image)
						$StackPanel.AddChild($TextBlock)
					
						Invoke-WPFMessageBox -Content $StackPanel -Title "WARNING!" -TitleBackground "Orange" -TitleTextForeground "Black" -TitleFontSize "20" -ButtonType OK -WindowHost $SyncHash.Window
					}

				})
			
				# File Run
				$synchash.Menu_File_Run.Add_Click({
					Run-Reports					
				})

				# File Run
				$synchash.Menu_File_Save.Add_Click({

				})
				# File Exit
				$synchash.Menu_File_Exit.Add_Click({
					close-window
				})
		
			#endregion Menu actions

			#region settings tab
				# Output Directory Browse Button
				$synchash.btn_settings_OutputDir.Add_Click({

					# Location dialog for selecting folder
					$FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
					$null = $FolderDialog.ShowDialog()

					# Setting the outputpath
					$synchash.txt_Settings_outputDir.Text = $FolderDialog.SelectedPath
					$VariableHash.OutputPath = $FolderDialog.SelectedPath
				})

				# Credentials validate button
				$SyncHash.btn_settings_Credentials_Validate.Add_Click({

					$M365username = $synchash.txt_Settings_M365Username.Text
					$M365password = $synchash.txt_Settings_M365Password.Password

					Validate-Credentials -UsrNm $M365username -Passwd $M365password

					 If($Global:ValidCreds -eq $false){

						$Source = -join($VariableHash.IconsDir,"\appbar.warning.png")
						$Image = New-Object System.Windows.Controls.Image
						$Image.Source = $Source
						$Image.Height = [System.Drawing.Image]::FromFile($Source).Height
						$Image.Width = [System.Drawing.Image]::FromFile($Source).Width
						$Image.Margin = 5
					 
						$TextBlock = New-Object System.Windows.Controls.TextBlock
						$TextBlock.Text = "Incorrect credentials!"
						$TextBlock.Padding = 10
						$TextBlock.FontFamily = "Verdana"
						$TextBlock.FontSize = 16
						$TextBlock.TextWrapping = "Wrap"
						$TextBlock.Width = 350
									
						$StackPanel = New-Object System.Windows.Controls.StackPanel
						$StackPanel.Orientation = "Horizontal"
						$StackPanel.Width = 400
						$StackPanel.AddChild($Image)
						$StackPanel.AddChild($TextBlock)
					
						Invoke-WPFMessageBox -Content $StackPanel -Title "WARNING!" -TitleBackground "Orange" -TitleTextForeground "Black" -TitleFontSize "20" -ButtonType OK -WindowHost $SyncHash.Window
						$synchash.Prereq_IMG_CredValidate.Source = "$($VariableHash.IconsDir)\Check_Red.ico"
					 }
					
					If($Global:ValidCreds -eq $true){

						$Source = -join($VariableHash.IconsDir,"\appbar.Check.png")
						$Image = New-Object System.Windows.Controls.Image
						$Image.Source = $Source
						$Image.Height = [System.Drawing.Image]::FromFile($Source).Height
						$Image.Width = [System.Drawing.Image]::FromFile($Source).Width
						$Image.Margin = 5
					 
						$TextBlock = New-Object System.Windows.Controls.TextBlock
						$TextBlock.Text = "Credentials validated!"
						$TextBlock.Padding = 10
						$TextBlock.FontFamily = "Verdana"
						$TextBlock.FontSize = 16
						$TextBlock.TextWrapping = "Wrap"
						$TextBlock.Width = 380
									
						$StackPanel = New-Object System.Windows.Controls.StackPanel
						$StackPanel.Orientation = "Horizontal"
						$StackPanel.Width = 400
						$StackPanel.AddChild($Image)
						$StackPanel.AddChild($TextBlock)

						$VariableHash.M365username = $synchash.txt_Settings_M365Username.Text
						$VariableHash.M365password = $synchash.txt_Settings_M365Password.Password
											
						Invoke-WPFMessageBox -Content $StackPanel -Title "Success!" -TitleBackground "Green" -TitleTextForeground "Black" -TitleFontSize "20" -ButtonType OK -WindowHost $SyncHash.Window
						$synchash.Prereq_IMG_CredValidate.Source = "$($VariableHash.IconsDir)\Check_Green.ico"
					 }
				})

			#endregion settings tab	
		#endregion Control actions

		#region Body
			# Loaded GUI
			$Timestamp = (get-date -Format HH:mm:ss)
			Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - User interface loaded!"

			#region Prerequisite testing
				$Prereq_Runspace =[runspacefactory]::CreateRunspace()
				$Prereq_Runspace.ApartmentState = "STA"
				$Prereq_Runspace.ThreadOptions = "ReuseThread"
				$Prereq_Runspace.Open()

				# Passing variables
				$Prereq_Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
				$Prereq_Runspace.SessionStateProxy.SetVariable("VariableHash",$VariableHash)

				$PowerShell = [PowerShell]::Create().AddScript({

					# Importing module
					Import-Module "$($VariableHash.ModuleDir)M365InventoryAutomaton.psm1" -Force -DisableNameChecking
					
					#region IsAdmin testing
						$Timestamp = (get-date -Format HH:mm:ss)
						Update-control -Synchash $synchash -control txt_output -property Text -Append -value "[$timestamp] - Testing user context..."
						# Retrieving Windows Security Principal
						$ThisPrincipal = new-object System.Security.principal.windowsprincipal( [System.Security.Principal.WindowsIdentity]::GetCurrent())
						
						# Checking if the user in in the Administrator Role
						$IsAdmin = $ThisPrincipal.IsInRole("Administrators")

						If($IsAdmin){
							Update-control -Synchash $synchash -control IMG_Prereq_Admin -property Source -value "$($VariableHash.IconsDir)\check_Green.ico"
						}Else{
							Update-control -Synchash $synchash -control IMG_Prereq_Admin -property Source -value "$($VariableHash.IconsDir)\check_Red.ico"
						}
					#endregion IsAdmin testing

					#region Connectivity Check
						$Timestamp = (get-date -Format HH:mm:ss)
						Update-control -Synchash $synchash -control txt_output -property Text -Append -value "[$timestamp] - Testing internet connectivity state..."

						# getting internet connectivity state
						$HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)

						If($HasInternetAccess){
							Update-control -Synchash $synchash -control IMG_Prereq_Internet -property Source -value "$($VariableHash.IconsDir)\check_Green.ico"
						}Else{
							Update-control -Synchash $synchash -control IMG_Prereq_Internet -property Source -value "$($VariableHash.IconsDir)\check_Red.ico"
						}
					#endregion Connectivity Check

					#region Sharepoint Online module
						$Timestamp = (get-date -Format HH:mm:ss)
						Update-control -Synchash $synchash -control txt_output -property Text -Append -value "[$timestamp] - Testing for SharePoint Online Powershell Module..."

						If((Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\Microsoft.Online.SharePoint.PowerShell\' -Recurse -Filter "Microsoft.Online.SharePoint.PowerShell.psd1")){
							Update-control -Synchash $synchash -control IMG_Prereq_SPOMgmtShell -property Source -value "$($VariableHash.IconsDir)\check_Green.ico"
						}Else{
							Update-control -Synchash $synchash -control IMG_Prereq_SPOMgmtShell -property Source -value "$($VariableHash.IconsDir)\check_Red.ico"
						}
					#endregion Sharepoint Online module

					#region Azure Online module
						$Timestamp = (get-date -Format HH:mm:ss)
						Update-control -Synchash $synchash -control txt_output -property Text -Append -value "[$timestamp] - Testing for Azure Active Directory Module..."

						If((Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\AzureAD\' -Recurse -Filter "AzureAD.psd1")){
							Update-control -Synchash $synchash -control IMG_Prereq_AADModule -property Source -value "$($VariableHash.IconsDir)\check_Green.ico"
						}Else{
							Update-control -Synchash $synchash -control IMG_Prereq_AADModule -property Source -value "$($VariableHash.IconsDir)\check_Red.ico"
						}
					#endregion Azure Online module

					#region MSTeams Online module
						$Timestamp = (get-date -Format HH:mm:ss)
						Update-control -Synchash $synchash -control txt_output -property Text -Append -value "[$timestamp] - Testing for Microsoft Teams Module..."

						If((get-childitem -path 'C:\Program Files\WindowsPowerShell\Modules\MicrosoftTeams' -Recurse -Filter "MicrosoftTeams.psd1")){
							Update-control -Synchash $synchash -control IMG_Prereq_MSTeamsModule -property Source -value "$($VariableHash.IconsDir)\check_Green.ico"
						}Else{
							Update-control -Synchash $synchash -control IMG_Prereq_MSTeamsModule -property Source -value "$($VariableHash.IconsDir)\check_Red.ico"
						}
					#endregion MSTeams Online module
				})

				$PowerShell.Runspace = $Prereq_Runspace
				$data = $PowerShell.BeginInvoke()	

			#endregion Prerequisite testing
		#endregion Body

		# Show GUI
		$Splashhash.window.Dispatcher.Invoke("Normal",[action]{ $Splashhash.window.close();$Splash_Runspace.close();$Splash_Runspace.dispose()})
		$SyncHash.Window.ShowDialog() | Out-Null
		$VariableHash.Error = $Error
	
	})

	# Invoking GUI
	$data = $psCmd.BeginInvoke()
#endregion Main window