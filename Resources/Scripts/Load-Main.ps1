
#region Variables
	#region Creating Synchronized collections
			$Global:SyncHash = [hashtable]::Synchronized(@{})
	#endregion Creating Synchronized collections


#endregion Variables

#region Main window
# Runspace creation
	$syncHash.host = $Host
	$Main_Runspace =[runspacefactory]::CreateRunspace()
	$syncHash.Runspace = $Main_Runspace
	$Main_Runspace.ApartmentState = "STA"
	$Main_Runspace.ThreadOptions = "ReuseThread"
	$Main_Runspace.Open()

	# Passing variables
	$Main_Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
	$Main_Runspace.SessionStateProxy.SetVariable("VariableHash",$VariableHash)
	$Main_Runspace.SessionStateProxy.SetVariable("Main_Runspace",$Main_Runspace)

	# Create powershell object which will containt the code we're running in the runspace
	$psCmd = [PowerShell]::Create()

	# Add runspace to Powershell object
	$psCmd.Runspace = $Main_Runspace

	[Void]$psCmd.AddScript({
		#region XAML Prep
			# Load Required Assemblies
			Add-Type –assemblyName PresentationFramework # Required to show the GUI
			Add-Type –assemblyName PresentationCore # Required for MahApps.Metro
			Add-Type –assemblyName WindowsBase # Required for MahApps.Metro
			Add-Type –assemblyName System.Drawing # Required for MahApps.Metro
			Add-Type -AssemblyName System.Windows.Forms # Required to use the folder browser dialog
			[System.Reflection.Assembly]::LoadFrom("$($VariableHash.LibDir)\ControlzEx.dll") | out-null
			[System.Reflection.Assembly]::LoadFrom("$($VariableHash.LibDir)\MahApps.Metro.dll") | out-null
			[System.Reflection.Assembly]::LoadFrom("$($VariableHash.LibDir)\MahApps.Metro.IconPacks.dll") | out-null
		
			# Loading XAML code	
			[xml]$xaml = Get-Content -Path "$($VariableHash.FormsDir)\MainWindow.xaml"
		
			# Loading in to XML Node reader
			$reader = (New-Object System.Xml.XmlNodeReader $xaml)
		
			# Loading XML Node reader in to $SyncHash window property to launch later
			$SyncHash.Window = [Windows.Markup.XamlReader]::Load($reader)
		#endregion XAML Prep

		# Importing module
		Import-Module "$($VariableHash.ModuleDir)\M365InventoryAutomaton.psm1" -Force -DisableNameChecking

		#region Connecting elements
			$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach {
				$Synchash.Add($_.Name,$SyncHash.Window.FindName($_.Name))
			}
		<# Trying somethign else
			#region TitleBar
				# Titlebar textbxox
				$synchash.MainWindow = $SyncHash.Window.FindName("MainWindow")
				$synchash.TitleBar = $SyncHash.Window.FindName("TitleBar")
				
				# TitleBar text box
				$synchash.IMG_TitleBar_Icon = $SyncHash.Window.FindName("IMG_TitleBar_Icon")
			
				# Blog
				$synchash.BTN_Blog = $SyncHash.Window.FindName("BTN_Blog")
				$synchash.IMG_TitleBar_Blog = $SyncHash.Window.FindName("IMG_TitleBar_Blog")
				
				# Github
				$synchash.BTN_GitHub = $SyncHash.Window.FindName("BTN_GitHub")
				$synchash.IMG_TitleBar_GitHub = $SyncHash.Window.FindName("IMG_TitleBar_GitHub")
		
				# Close
				$synchash.BTN_Close = $SyncHash.Window.FindName("BTN_Close")
				$synchash.IMG_TitleBar_Close = $SyncHash.Window.FindName("IMG_TitleBar_Close")
			#endregion TitleBar

			#region Home Tab
				$synchash.TabItem_Home = $SyncHash.Window.FindName("TabItem_Home")
				$synchash.IMG_Home_User = $SyncHash.Window.FindName("IMG_Home_User")
				$synchash.txt_Home_Username = $SyncHash.Window.FindName("txt_Home_Username")
				$synchash.IMG_Home_Password = $SyncHash.Window.FindName("IMG_Home_Password")
				$synchash.txt_Home_Password = $SyncHash.Window.FindName("txt_Home_Password")
				$synchash.IMG_Home_OutputDir = $SyncHash.Window.FindName("IMG_Home_OutputDir")
				$synchash.txt_Home_OutPutDir = $SyncHash.Window.FindName("txt_Home_OutPutDir")
				$synchash.IMG_Home_Tenant = $SyncHash.Window.FindName("IMG_Home_Tenant")
				$synchash.txt_Home_Tenant = $SyncHash.Window.FindName("txt_Home_Tenant")
				$synchash.IMG_Home_region = $SyncHash.Window.FindName("IMG_Home_region")
				$synchash.DD_Home_Region = $SyncHash.Window.FindName("DD_Home_Region")
				$synchash.IMG_Home_GraphAPPID = $SyncHash.Window.FindName("IMG_Home_GraphAPPID")
				$synchash.txt_Home_GraphAppID = $SyncHash.Window.FindName("txt_Home_GraphAppID")
				$synchash.chk_MFA = $SyncHash.Window.FindName("chk_MFA")
				$synchash.chk_Output = $SyncHash.Window.FindName("chk_Output")
		
				$synchash.BTN_Home_GO = $SyncHash.Window.FindName("BTN_Home_GO")
				$synchash.BTN_Home_Browse = $SyncHash.Window.FindName("BTN_Home_Browse")
				$synchash.BTN_Home_Validate = $SyncHash.Window.FindName("BTN_Home_Validate")

				$synchash.TXT_Output = $SyncHash.Window.FindName("TXT_Output")

				# Connections
				$synchash.IMG_Conn_AAD = $SyncHash.Window.FindName("IMG_Conn_AAD")
				$synchash.IMG_Conn_EXO = $SyncHash.Window.FindName("IMG_Conn_EXO")
				$synchash.IMG_Conn_MSOL = $SyncHash.Window.FindName("IMG_Conn_MSOL")
				$synchash.IMG_Conn_SPO = $SyncHash.Window.FindName("IMG_Conn_SPO")

				# Reports
				$synchash.IMG_Report_AllSites = $SyncHash.Window.FindName("IMG_Report_AllSites")
				$synchash.IMG_Report_AllSiteGroups = $SyncHash.Window.FindName("IMG_Report_AllSiteGroups")
				$synchash.IMG_Report_SPOLibraries = $SyncHash.Window.FindName("IMG_Report_Libraries")
				$synchash.IMG_Report_AADUsers = $SyncHash.Window.FindName("IMG_Report_AADUsers")
				$synchash.IMG_Report_Graph = $SyncHash.Window.FindName("IMG_Report_Graph")
				$synchash.IMG_Report_AADDelUser = $SyncHash.Window.FindName("IMG_Report_AADDelusers")
				$synchash.IMG_Report_AADContacts = $SyncHash.Window.FindName("IMG_Report_AADContacts")
				$synchash.IMG_Report_AADGroups = $SyncHash.Window.FindName("IMG_Report_AADGroups")
				$synchash.IMG_Report_AADDomains = $SyncHash.Window.FindName("IMG_Report_AADDomains")
				$synchash.IMG_Report_EXOMailboxes = $SyncHash.Window.FindName("IMG_Report_EXOMailboxes")
				$synchash.IMG_Report_EXOArchives = $SyncHash.Window.FindName("IMG_Report_EXOArchives")
				$synchash.IMG_Report_EXOGroups = $SyncHash.Window.FindName("IMG_Report_EXOGroups")
				$synchash.IMG_Report_SPOShared = $SyncHash.Window.FindName("IMG_Report_SPOShared")
				$synchash.IMG_Report_ExternalUsers = $SyncHash.Window.FindName("IMG_Report_ExternalUsers")
				$synchash.IMG_Report_Flow = $SyncHash.Window.FindName("IMG_Report_Flow")
				$synchash.IMG_Report_PowerApps = $SyncHash.Window.FindName("IMG_Report_PowerApps")
				$synchash.IMG_Report_AzureVM = $SyncHash.Window.FindName("IMG_Report_AzureVM")
				$synchash.IMG_Report_LicenseUsage = $SyncHash.Window.FindName("IMG_Report_LicenseUsage")
				$synchash.IMG_Report_PowerBIWorkspaces = $SyncHash.Window.FindName("IMG_Report_PowerBIWorkspaces")
				$synchash.IMG_Report_PowerBIDashboards = $SyncHash.Window.FindName("IMG_Report_PowerBIDashboards")
				$synchash.IMG_Report_PowerBIReports = $SyncHash.Window.FindName("IMG_Report_PowerBIReports")
				$synchash.IMG_Report_PowerBIDatasources = $SyncHash.Window.FindName("IMG_Report_PowerBIDatasources")
		
				# Exchange Online - Mailboxes
				$synchash.EXOSubTab_Recipients = $SyncHash.Window.FindName("EXOSubTab_Recipients")
				$synchash.DataGrid_EXOMailboxes = $SyncHash.Window.FindName("DataGrid_EXOMailboxes")				
			#endregion Home Tab
		#>
		#endregion Connecting elements

		#region Configuring elements
			$SyncHash.MainWindow.Title = "M365 Inventory Automaton"
			#region TitleBar
				# Titlebar textbxox
				$synchash.TitleBar.Text= "M365 Inventory Automaton"
				
				# TitleBar text box
				$synchash.IMG_TitleBar_Icon.Source = "$($VariableHash.IconsDir)\Toaster.ico"
			
				# Blog
				$synchash.BTN_Blog.Source = "$($VariableHash.IconsDir)\appbar.browser.png"
								
				# Github
				$synchash.IMG_TitleBar_GitHub.Source = "$($VariableHash.IconsDir)\appbar.social.github.octocat.png"
		
				# Close
				$synchash.IMG_TitleBar_Close.Source = "$($VariableHash.IconsDir)\appbar.Close_black.png"
			#endregion TitleBar

			#region Home Tab
				$synchash.IMG_Home_User = $SyncHash.Window.FindName("IMG_Home_User")
				$synchash.IMG_Home_Password = $SyncHash.Window.FindName("IMG_Home_Password")
				$synchash.IMG_Home_OutputDir = $SyncHash.Window.FindName("IMG_Home_OutputDir")
				$synchash.IMG_Home_Tenant = $SyncHash.Window.FindName("IMG_Home_Tenant")
				$synchash.IMG_Home_region = $SyncHash.Window.FindName("IMG_Home_region")
				$synchash.IMG_Home_GraphAPPID = $SyncHash.Window.FindName("IMG_Home_GraphAPPID")
				
				# Connections
				$synchash.IMG_Conn_AAD.Source = "$($VariableHash.IconsDir)\Light.Red.ico"
				$synchash.IMG_Conn_EXO.Source = "$($VariableHash.IconsDir)\Light.Red.ico"
				$synchash.IMG_Conn_MSOL.Source = "$($VariableHash.IconsDir)\Light.Red.ico"
				$synchash.IMG_Conn_SPO.Source = "$($VariableHash.IconsDir)\Light.Red.ico"

				# Reports
				$synchash.IMG_Report_AllSites.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_AllSiteGroups.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_SPOLibraries.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_AADUsers.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_Graph.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_AADDelUser.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_AADContacts.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_AADGroups.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_AADDomains.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_EXOMailboxes.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_EXOArchives.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				$synchash.IMG_Report_EXOGroups.Source = "$($VariableHash.IconsDir)\Light.Orange.ico"
				
			#endregion Home Tab

			#region Drop Down adding items
				$synchash.DD_Home_Region.AddChild("Default")
				$synchash.DD_Home_Region.AddChild("Germany")
				$synchash.DD_Home_Region.AddChild("China")
				$synchash.DD_Home_Region.AddChild("AzurePPE")
				$synchash.DD_Home_Region.AddChild("USGovernment")
			#endregion Drop Down adding items
		#endregion Configuring elements

		#region Event handling
			
			#region Allowing for the window to be dragged
				$Synchash.Window.FindName('Grid').Add_MouseLeftButtonDown({
					$Synchash.Window.DragMove()
				})
			#endregion Allowing for the window to be dragged

			#region Close the window
				$SyncHash.BTN_Close.Add_Click({
					$SyncHash.Window.Close()

					# Cleaning up validation runspace
					$Validation_Runspace.close()
					$Validation_Runspace.Dispose()

					# Cleaning up reporting runspace
					$Report_Runspace.close()
					$Report_Runspace.Dispose()

					# Cleaning up main runspace
					$Main_Runspace.close()
					$Main_Runspace.Dispose()

					# Invoking garbage collection
					[gc]::Collect()
					[gc]::WaitForPendingFinalizers() 
				})
			#endregion Close the window

			#region Blog
				$Synchash.BTN_Blog.Add_Click({[system.Diagnostics.Process]::Start('http://geekswithBlogs.net/marcde')})
			#endregion Blog

			#region GitHub
				$Synchash.BTN_GitHub.Add_Click({[system.Diagnostics.Process]::Start('http://github.com/Toasterlabs')})
			#endregion Github

			#region Validate Credentials
			$synchash.BTN_Home_Validate.Add_click({
				# Reporting Event
				$message = "Starting credential validation (This may take a few moments to complete...)"
				Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
				Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

				$M365username = $synchash.txt_Home_Username.Text
				$M365password = $synchash.txt_Home_Password.Password
	
					
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
					
						# Reporting Event
						$message = "Supplied credentials are incorrect!"
						Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
						Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

						Invoke-WPFMessageBox -Content $StackPanel -Title "WARNING!" -TitleBackground "Orange" -TitleTextForeground "Black" -TitleFontSize "20" -ButtonType OK -WindowHost $SyncHash.Window
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

						$VariableHash.M365username = $synchash.txt_Home_Username.Text
						$VariableHash.M365password = $synchash.txt_Home_Password.Password
						$synchash.BTN_Home_GO.IsEnabled = $true
				
						# Reporting Event
						$message = "Credentials validated!"
						Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
						Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile						
					
						Invoke-WPFMessageBox -Content $StackPanel -Title "Success!" -TitleBackground "Green" -TitleTextForeground "Black" -TitleFontSize "20" -ButtonType OK -WindowHost $SyncHash.Window
					}

			})
			#endregion Validate Credentials

			#region browse for folder
			$SyncHash.BTN_Home_Browse.Add_click({
				# Reporting Event
				$message = "Selecting output path"
				Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
				Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

				# Location dialog for selecting folder
				$FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
				$null = $FolderDialog.ShowDialog()

				# Setting the outputpath
				$synchash.txt_Home_OutPutDir.Text = $FolderDialog.SelectedPath
				$VariableHash.OutputPath = $FolderDialog.SelectedPath
			})
			#endregion browse for folder

			#region report collection
			$SyncHash.BTN_Home_go.Add_click({
				# Reporting Event
				$message = "Starting information gathering"
				Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
				Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

				Run-reports
			})
			#endregion report collection

		#endregion Event handling

		# Reporting event
		$message = "Tool load complete!"
		Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
		Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

		# Show GUI
		$SyncHash.Window.ShowDialog() | Out-Null
		$VariableHash.Error = $Error
	
	})

	# Invoking GUI
	$data = $psCmd.BeginInvoke()
#endregion Main window