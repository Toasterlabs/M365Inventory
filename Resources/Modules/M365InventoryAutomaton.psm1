Function Invoke-WPFMessageBox {

    <#
		Function by SMS Agent
		https://smsagent.blog/2017/08/24/a-customisable-wpf-messagebox-for-powershell/
	#>
    
    # Define Parameters
    [CmdletBinding()]
    Param
    (
        # The popup Content
        [Parameter(Mandatory=$True,Position=0)]
        [Object]$Content,

        # The window title
        [Parameter(Mandatory=$false,Position=1)]
        [string]$Title,

        # The buttons to add
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
        [array]$ButtonType = 'OK',

        # The buttons to add
        [Parameter(Mandatory=$false,Position=3)]
        [array]$CustomButtons,

        # Content font size
        [Parameter(Mandatory=$false,Position=4)]
        [int]$ContentFontSize = 14,

        # Title font size
        [Parameter(Mandatory=$false,Position=5)]
        [int]$TitleFontSize = 14,

        # BorderThickness
        [Parameter(Mandatory=$false,Position=6)]
        [int]$BorderThickness = 0,

        # CornerRadius
        [Parameter(Mandatory=$false,Position=7)]
        [int]$CornerRadius = 8,

        # ShadowDepth
        [Parameter(Mandatory=$false,Position=8)]
        [int]$ShadowDepth = 3,

        # BlurRadius
        [Parameter(Mandatory=$false,Position=9)]
        [int]$BlurRadius = 20,

        # WindowHost
        [Parameter(Mandatory=$false,Position=10)]
        [object]$WindowHost,

        # Timeout in seconds,
        [Parameter(Mandatory=$false,Position=11)]
        [int]$Timeout,

        # Code for Window Loaded event,
        [Parameter(Mandatory=$false,Position=12)]
        [scriptblock]$OnLoaded,

        # Code for Window Closed event,
        [Parameter(Mandatory=$false,Position=13)]
        [scriptblock]$OnClosed

    )

    # Dynamically Populated parameters
    DynamicParam {
        
        # Add assemblies for use in PS Console 
        Add-Type -AssemblyName System.Drawing, PresentationCore
        
        # ContentBackground
        $ContentBackground = 'ContentBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
        

        # FontFamily
        $FontFamily = 'FontFamily'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)  
        $arrSet = [System.Drawing.FontFamily]::Families | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
        $PSBoundParameters.FontFamily = "Segui"

        # TitleFontWeight
        $TitleFontWeight = 'TitleFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

        # ContentFontWeight
        $ContentFontWeight = 'ContentFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
        

        # ContentTextForeground
        $ContentTextForeground = 'ContentTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

        # TitleTextForeground
        $TitleTextForeground = 'TitleTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

        # BorderBrush
        $BorderBrush = 'BorderBrush'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.BorderBrush = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


        # TitleBackground
        $TitleBackground = 'TitleBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

        # ButtonTextForeground
        $ButtonTextForeground = 'ButtonTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ButtonTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

        # Sound
        $Sound = 'Sound'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        #$ParameterAttribute.Position = 14
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav','')
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    Begin {
        Add-Type -AssemblyName PresentationFramework
    }
    
    Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid >
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >                   
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

    # Load the window from XAML
    $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

    # Custom function to add a button
    Function Add-Button {
        Param($Content)
        $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
        $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
        $ButtonText.Text = "$Content"
        $Button.Content = $ButtonText
        $Button.Add_MouseEnter({
            $This.Content.FontSize = "17"
        })
        $Button.Add_MouseLeave({
            $This.Content.FontSize = "16"
        })
        $Button.Add_Click({
			New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -scope Global -Force
            $Window.Close()
        })
        $Window.FindName('ButtonHost').AddChild($Button)
    }

    # Add buttons
    If ($ButtonType -eq "OK")
    {
        Add-Button -Content "OK"
    }

    If ($ButtonType -eq "OK-Cancel")
    {
        Add-Button -Content "OK"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Abort-Retry-Ignore")
    {
        Add-Button -Content "Abort"
        Add-Button -Content "Retry"
        Add-Button -Content "Ignore"
    }

    If ($ButtonType -eq "Yes-No-Cancel")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Yes-No")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
    }

    If ($ButtonType -eq "Retry-Cancel")
    {
        Add-Button -Content "Retry"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Cancel-TryAgain-Continue")
    {
        Add-Button -Content "Cancel"
        Add-Button -Content "TryAgain"
        Add-Button -Content "Continue"
    }

    If ($ButtonType -eq "None" -and $CustomButtons)
    {
        Foreach ($CustomButton in $CustomButtons)
        {
            Add-Button -Content "$CustomButton"
        }
    }

    # Remove the title bar if no title is provided
    If ($Title -eq "")
    {
        $TitleBar = $Window.FindName('TitleBar')
        $Window.FindName('StackPanel').Children.Remove($TitleBar)
    }

    # Add the Content
    If ($Content -is [String])
    {
        # Replace double quotes with single to avoid quote issues in strings
        If ($Content -match '"')
        {
            $Content = $Content.Replace('"',"'")
        }
        
        # Use a text box for a string value...
        $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
        $Window.FindName('ContentHost').AddChild($ContentTextBox)
    }
    Else
    {
        # ...or add a WPF element as a child
        Try
        {
            $Window.FindName('ContentHost').AddChild($Content) 
        }
        Catch
        {
            $_
        }        
    }

    # Enable window to move when dragged
    $Window.FindName('Grid').Add_MouseLeftButtonDown({
        $Window.DragMove()
    })

    # Activate the window on loading
    If ($OnLoaded)
    {
        $Window.Add_Loaded({
            $This.Activate()
            Invoke-Command $OnLoaded
        })
    }
    Else
    {
        $Window.Add_Loaded({
            $This.Activate()
        })
    }
    

    # Stop the dispatcher timer if exists
    If ($OnClosed)
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
            Invoke-Command $OnClosed
        })
    }
    Else
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
        })
    }
    

    # If a window host is provided assign it as the owner
    If ($WindowHost)
    {
        $Window.Owner = $WindowHost
        $Window.WindowStartupLocation = "CenterOwner"
    }

    # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
    If ($Timeout)
    {
        $Stopwatch = New-object System.Diagnostics.Stopwatch
        $TimerCode = {
            If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
            {
                $Stopwatch.Stop()
                $Window.Close()
            }
        }
        $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
        $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
        $DispatcherTimer.Add_Tick($TimerCode)
        $Stopwatch.Start()
        $DispatcherTimer.Start()
    }

    # Play a sound
    If ($($PSBoundParameters.Sound))
    {
        $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
        $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
        $SoundPlayer.Add_LoadCompleted({
            $This.Play()
            $This.Dispose()
        })
        $SoundPlayer.LoadAsync()
    }

    # Display the window
    $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

    }
}

Function Update-Control {
	Param (
		$syncHash,
        $Control,
        $Property,
        $Value,
        [switch]$AppendContent
	)

	If($AppendContent){
		$syncHash.Window.Dispatcher.Invoke(
			[action]{
				# Updating Control
				$timestamp = (get-date -Format HH:mm:ss)
				$current = $syncHash.$Control.$Property
				$Value = "$current`n$timestamp - $value"
				$syncHash.$Control.$Property = $Value
				$syncHash.$Control.ScrollToEnd()
			},"Normal")
	}Else{
		$syncHash.Window.Dispatcher.Invoke(
			[action]{
				$syncHash.$Control.$Property = $Value
			},"Normal")
	}
}

function Validate-Credentials{
    Param(
    [Parameter(Mandatory=$True)]
    $UsrNm,
    [Parameter(Mandatory=$True)]
    $Passwd
    )

    $secpswd = ConvertTo-SecureString $Passwd -AsPlainText -Force
    $O365crds = new-object -typename System.Management.Automation.PSCredential -argumentlist $UsrNm,$secpswd

    try
    {
		# Reporting Event
		$message = "Attempting connection..."
		Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
		
        $Global:ErrorActionPreference = 'stop'
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365crds -Authentication Basic -AllowRedirection
		$Global:ValidCreds = $true

		# Reporting Event
		$message = "Connection successfull!"
		Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
		    }
    Catch [System.Net.WebException],[System.Exception]
    {
		# Reporting Event
		$message = "Failed to establish connection..."
	    Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	    Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

        $Global:ValidCreds = $false
    }
    Finally
    {
	# Reporting Event
	$message = "Hanging up..."
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
    $Global:ErrorActionPreference = 'continue'
    Remove-PSSession $Session -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
}

Function Launch-Admin{
	Param(
		$ScriptToLaunch
	)

	# Create a new process object that starts PowerShell
   $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
   
   # Specify the current script path and name as a parameter
   $newProcess.Arguments = $myInvocation.MyCommand.Definition;;
   
   # Indicate that the process should be elevated
   $newProcess.Verb = "runas";
   
   # Start the new process
   [System.Diagnostics.Process]::Start($newProcess);
   
   # Exit from the current, unelevated, process
   exit
}

Function Close-Window{
	$SyncHash.Window.Close()

	# Cleaning up reporting runspace
	$Report_Runspace.close()
	$Report_Runspace.Dispose()

	# Cleaning up main runspace
	$Main_Runspace.close()
	$Main_Runspace.Dispose()

	# Invoking garbage collection
	[gc]::Collect()
	[gc]::WaitForPendingFinalizers()    

}

function Connect-M365{
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$False,HelpMessage='Credentials to connect to Microsoft 365')]
		$Credentials,
		[Parameter(Mandatory=$false,HelpMessage='Region to connect to')]
		[ValidateSet('Germany', 'China', 'AzurePPE', 'USGovernment', 'Default')]
		$Region = "Default",
		[SWITCH]$ExchangeOnline,
		[SWITCH]$AzureAD,
		[SWITCH]$MSOL,
		[Parameter(Mandatory = $false, HelpMessage = 'Connect to Sharepoint Online', ParameterSetName = "SharePointOnline")]
		[SWITCH]$SharePointOnline,
		[Parameter(Mandatory = $false, HelpMessage = 'Tenant Name (Contoso, NOT Contoso.onmicrosoft.com)', ParameterSetName = "SharePointOnline")]
		[STRING]$TenantName,
		[Parameter(Mandatory = $false, HelpMessage = 'Skype for Business Online')]
		[SWITCH]$SkypeForBusinessOnline,
		[Parameter(Mandatory = $false, HelpMessage = 'Security and Compliance Center')]
		[SWITCH]$SCC,
		[Parameter(Mandatory = $false, HelpMessage = 'Should we use Multi Factor Authentication?')]
		[SWITCH]$MFA
	)

	# Reporting event
	$message = "Starting connection process..."
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	#region Variables
	## Emtpy Hashtable for region
	$global:M365Services = @{}

	## Defining URI based on region
	# Reporting event
	$message = "Defining connection URI based on location"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	Switch($Region){
			'Germany' {
				$global:M365Services['ConnectionEndpointUri'] = 'https://outlook.office.de/PowerShell-LiveID'
				$global:M365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
				$global:M365Services['AzureADAuthorizationEndpointUri'] = 'https://login.microsoftonline.de/common'
				$global:M365Services['SharePointRegion'] = 'Germany'
				$global:M365Services['AzureEnvironment'] = 'AzureGermanyCloud'
			}
			'China' {
				$global:M365Services['ConnectionEndpointUri'] = 'https://partner.outlook.cn/PowerShell-LiveID'
				$global:M365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
				$global:M365Services['AzureADAuthorizationEndpointUri'] = 'https://login.chinacloudapi.cn/common'
				$global:M365Services['SharePointRegion'] = 'China'
				$global:M365Services['AzureEnvironment'] = 'AzureChinaCloud'
			}
			'AzurePPE' {
				$global:M365Services['ConnectionEndpointUri'] = ''
				$global:M365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
				$global:M365Services['AzureADAuthorizationEndpointUri'] = ''
				$global:M365Services['SharePointRegion'] = ''
				$global:M365Services['AzureEnvironment'] = 'AzurePPE'
			}
			'USGovernment' {
				$global:M365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
				$global:M365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
				$global:M365Services['AzureADAuthorizationEndpointUri'] = 'https://login-us.microsoftonline.com/'
				$global:M365Services['SharePointRegion'] = 'ITAR'
				$global:M365Services['AzureEnvironment'] = 'AzureUSGovernment'
			}
			default {
				$global:M365Services['ConnectionEndpointUri'] = 'https://outlook.office365.com/PowerShell-LiveId'
				$global:M365Services['SCCConnectionEndpointUri'] = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId'
				$global:M365Services['AzureADAuthorizationEndpointUri'] = 'https://login.windows.net/common'
				$global:M365Services['SharePointRegion'] = 'Default'
				$global:M365Services['AzureEnvironment'] = 'AzureCloud'
			}
		}
	#endregion

	# Setting Execution Policy
	# Reporting event
	$message = "Setting remote execution policy"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	Set-ExecutionPolicy RemoteSigned -Force

	#region Connections
		# Connecting to Exchange Online
		If($ExchangeOnline){
			If($MFA){
				Write-Verbose "Connecting to Exchange Online using MFA"
				Connect-EXOPSSession -UserPrincipalName $Credentials.UserName -ConnectionUri $global:M365Services['ConnectionEndpointUri'] -AzureADAuthorizationEndPointUri $global:M365Services['AzureADAuthorizationEndpointUri']
			}Else{
				# Reporting event
				$message = "Connecting to Exchange Onine"
				Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
				Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

				# Creating the session
				$Session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $global:M365Services['ConnectionEndpointUri'] -Credential $Credentials -Authentication Basic -AllowRedirection
				# If the session was created, import it
				If($Session365){
					Try{
						Import-PSSession -Session $Session365 -AllowClobber -DisableNameChecking
						Update-control -Synchash $synchash -control IMG_Conn_EXO -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
					}Catch{
						Update-control -Synchash $synchash -control IMG_Conn_EXO -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
				}Else{
					Update-control -Synchash $synchash -control IMG_Conn_EXO -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
				}
			}
		}
		}
		# Connect to Azure AD
		If($AzureAD){
			# Reporting event
			$message = "Connecting to Azure Active Directory"
			Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
			Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

			Try{
				Connect-AzureAD -Credential $credentials -AzureEnvironmentName $global:M365Services['AzureEnvironment']
				Update-control -Synchash $synchash -control IMG_Conn_AAD -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
			}Catch{
				Update-control -Synchash $synchash -control IMG_Conn_AAD -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
			}
		}

		# Connect to MSOL
		If($MSOL){
			# Reporting event
			$message = "Connecting to MS Online"
			Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
			Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

			Try{
				Connect-MsolService -Credential $credentials -AzureEnvironment $global:M365Services['AzureEnvironment']
				Update-control -Synchash $synchash -control IMG_Conn_MSOL -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
			}Catch{
				Update-control -Synchash $synchash -control IMG_Conn_MSOL -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
			}
			
		}

		# Connect to SharePoint Online
		If($SharePointOnline){
			# Reporting event
			$message = "Connecting to SharePoint Online"
			Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
			Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

			Try{
				Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
				$tenantName = ($TenantName).split(".")[0]
				Connect-SPOService -Url https://$TenantName-admin.sharepoint.com -credential $credentials -Region $global:M365Services['SharePointRegion']
				Update-control -Synchash $synchash -control IMG_Conn_SPO -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
			}Catch{
				Update-control -Synchash $synchash -control IMG_Conn_M365 -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
			}

		}

		# Connect to Skype for Business Online
		If($SkypeForBusinessOnline){
			# Reporting event
			$message = "Connecting to Skype for Business Online"
			Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
			Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
			
			Import-Module SkypeOnlineConnector
			$sfboSession = New-CsOnlineSession -Credential $credentials
			If($sfboSession){Import-PSSession $sfboSession}
		}

		# Connect to security and compliance center
		If($SCC){
			If($MFA){
				Connect-IPPSSession -UserPrincipalName $Credentials.UserName -ConnectionUri $global:M365Services['SCCConnectionEndpointUri'] -AzureADAuthorizationEndPointUri $global:M365Services['AzureADAuthorizationEndpointUri']
			}Else{
				# Reporting event
				$message = "Connecting to Security and Compliance Center"
				Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
				Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

				$SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $global:M365Services['SCCConnectionEndpointUri'] -Credential $credentials -Authentication "Basic" -AllowRedirection
				If($SccSession){Import-PSSession $SccSession -Prefix cc}
			}
		}
	#endregion
	
}

Function Connect-ProvisioningWebServiceAPI{
	<#
		.SYNOPSIS
			Connects to the Office 365 provisioning web service API.

		.DESCRIPTION
			Connects to the Office 365 provisioning web service API.
			
			If a credential is specified, it will be used to establish a connection with the provisioning
			web service API.
			
			If a credential is not specified, an attempt is made to identify an existing connection to
			the provisioning web service API.  If an existing connection is identified, the existing
			connection is used.  If an existing connection is not identified, the user is prompted for
			credentials so that a new connection can be established.

		.PARAMETER Credential
			Specifies the credential to use when connecting to the provisioning web service API
			using Connect-MsolService.

		.EXAMPLE
			PS> ConnectProvisioningWebServiceAPI

		.EXAMPLE
			PS> ConnectProvisioningWebServiceAPI -Credential
			
		.INPUTS
			[System.Management.Automation.PsCredential]

		.OUTPUTS

		.NOTES

	#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $False)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	# if a credential was supplied, assume a new connection is intended and create a new
	# connection using specified credential
	If ($Credential)
	{
		If ((!$Credential) -or (!$Credential.Username) -or ($Credential.Password.Length -eq 0))
		{
			Write-warning -Message ("Invalid credential.  Please verify the credential and try again.")
			Exit
		}
		
		# connect to provisioning web service api
		Write-Verbose -Message "Connecting to the Office 365 provisioning web service API.  Please wait..."
		Connect-MsolService -Credential $Credential
		If($? -eq $False){WriteConsoleMessage -Message "Error while connecting to the Office 365 provisioning web service API.  Quiting..." -MessageType "Error";Exit}
	}
	Else
	{
		Write-Verbose -Message "Attempting to identify an open connection to the Office 365 provisioning web service API.  Please wait..." 
		$getMsolCompanyInformationResults = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
		If (!$getMsolCompanyInformationResults)
		{
			Write-Verbose -Message "Could not identify an open connection to the Office 365 provisioning web service API." 			If (!$Credential)
			{
				$Credential = $Host.UI.PromptForCredential("Enter Credential",
					"Enter the username and password of an Office 365 administrator account.",
					"",
					"userCreds")
			}
			If ((!$Credential) -or (!$Credential.Username) -or ($Credential.Password.Length -eq 0))
			{
				Write-Verbose -Message ("Invalid credential.  Please verify the credential and try again.")
				Exit
			}
			
			# connect to provisioning web service api
			Write-Verbose -Message "Connecting to the Office 365 provisioning web service API.  Please wait..."
			Connect-MsolService -Credential $Credential
			If($? -eq $False){WriteConsoleMessage -Message "Error while connecting to the Office 365 provisioning web service API.  Quiting..." -MessageType "Error";Exit}
			$getMsolCompanyInformationResults = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
			WriteConsoleMessage -Message ("Connected to Office 365 tenant named: `"{0}`"." -f $getMsolCompanyInformationResults.DisplayName) -MessageType "Information"
		}
		Else
		{
			Write-Warning -Message ("Connected to Office 365 tenant named: `"{0}`"." -f $getMsolCompanyInformationResults.DisplayName) 
		}
	}
	If (!$Script:Credential) {$Script:Credential = $Credential}
}

Function Get-SPOInventory{


	# Loading Libraries
	# Reporting Event
	$message = "Loading Sharepoint libraries"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	foreach ($Assembly in (Dir $variableHash.LibDir -Filter *.dll)) {
		Write-verbose "$loading $($Assembly.fullname)"
		[System.Reflection.Assembly]::LoadFrom($Assembly.fullName) | out-null
	}

	# Variables
	$Databases = @()

	# Creating credentials object
	$secpswd = ConvertTo-SecureString $VariableHash.M365password -AsPlainText -Force
	$Creds = new-object -typename System.Management.Automation.PSCredential -argumentlist $VariableHash.M365Username,$secpswd

	# Retrieving all licensed and unlicensed users
    # Reporting Event
	$message = "Retrieving users"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	$Users = Get-MsolUser -All 
	$UnLicensedUsers = Get-MsolUser -UnlicensedUsersOnly 

	# Retrieving all SPO sites, including Personal sites
	# Reporting Event
	$message = "Retrieving all SPO sites"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	$SitesIncludingPersonal = Get-SPOSite -IncludePersonalSite $true -Limit All -Detailed 

	# Adding User to the site collection admins (Fixing an rights error encountered)

	# Reporting event
	$message = "Adding Inventory requests to the site collection admins"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	foreach($site in $Sites){ 
	  Set-SPOUser -Site $site.Url -LoginName $Credentials.UserName -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue 
	} 

	# Reporting event
	$message = "Retrieving SharePoint Online Sites"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	

	foreach($site in $Sites){ 
	  $Users = Get-SPOUser -Site $site.Url -Limit All | Select * -Verbose 
	  foreach($User in $Users) 
	  { 
		$DB = New-Object PSObject 
		Add-Member -input $DB noteproperty 'SiteUrl' $site.Url  
		Add-Member -input $DB noteproperty 'DisplayName' $site.DisplayName 
		Add-Member -input $DB noteproperty 'LoginName' $site.LoginName 
		Add-Member -input $DB noteproperty 'IsSiteAdmin' $site.IsSiteAdmin 
		Add-Member -input $DB noteproperty 'IsGroup' $site.IsGroup 
		$Databases += $DB 
	  } 
	} 
	Update-control -Synchash $synchash -control IMG_Report_AllSites -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

	# Reporting Event
	$message = "Retrieving all site groups"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	$Databases = @()
	foreach($site in $Sites){ 
	  $Groups = Get-SPOSiteGroup -Site $site.Url -Limit 100 | Select * 
	  foreach($Group in $Groups) 
	  { 
		$DB = New-Object PSObject 
		Add-Member -input $DB noteproperty 'SiteUrl' $site.Url  
		Add-Member -input $DB noteproperty 'DisplayName' $site.Title 
		Add-Member -input $DB noteproperty 'LoginName' $site.LoginName 
		Add-Member -input $DB noteproperty 'OwnerLoginName' $site.OwnerLoginName 
		Add-Member -input $DB noteproperty 'OwnerTitle' $site.OwnerTitle 
		$RolesString = "" 
		foreach($Role in $Group.Roles) 
		{ 
		  $RolesString+=$Role 
		  $RolesString+="," 
		} 
		Add-Member -input $DB noteproperty 'Roles' $RolesString 
		$Databases += $DB 
	  } 
	} 
	Update-control -Synchash $synchash -control IMG_Report_AllSiteGroups -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

	$Databases = @()
	foreach($site in $site){ 
	  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($site.Url) 
	  #Authenticate 
	  $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName , $Creds.Password) 
	  $ctx.Credentials = $credentials 
 
	  #Retrieving the users in Site Collection 
	  $ctx.Load($ctx.Web.Webs) 
	  $Lists = $ctx.Web.Lists 
	  $ctx.Load($Lists) 
	  $ctx.ExecuteQuery() 
	  
	foreach($List in $Lists){ 
		if($List.Hidden -eq $false){ 
		  if($List.ItemCount -gt 100) { 
			$DB = New-Object PSObject 
			Add-Member -input $DB noteproperty 'SiteUrl' $site.Url  
			Add-Member -input $DB noteproperty 'Title' $List.Title 
			Add-Member -input $DB noteproperty 'ListType' $List.BaseType 
			Add-Member -input $DB noteproperty 'ItemCount' $List.ItemCount 
			$Databases += $DB 
		  } 
       
		} 
	  } 
 
	  foreach($Web in $ctx.Web.Webs){ 
		$Lists = $Web.Lists 
		$ctx.Load($Lists) 
		$ctx.ExecuteQuery() 
		foreach($List in $Lists){ 
		  if($List.Hidden -eq $false){ 
			if($List.ItemCount -gt 100){ 
			  $DB = New-Object PSObject 
			  Add-Member -input $DB noteproperty 'SiteUrl' $Web.Url  
			  Add-Member -input $DB noteproperty 'Title' $List.Title 
			  Add-Member -input $DB noteproperty 'ListType' $List.BaseType 
			  Add-Member -input $DB noteproperty 'ItemCount' $List.ItemCount 
			  $Databases += $DB 
			} 
		  } 
		} 
	  } 
	} 

	# Reporting Event
	$message = "Retrieving libraries"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	$Databases = @(); 
	foreach($site in $Sites){ 
	  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($site.url) 
	  
	  #Authenticate 
	  $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName , $Creds.Password) 
	  $ctx.Credentials = $credentials 
 
	  #Retrieving the users in Site Collection 
	  $ctx.Load($ctx.Web.Webs) 
	  $Lists = $ctx.Web.Lists 
	  $ctx.Load($Lists) 
	  $ctx.ExecuteQuery() 
	  foreach($List in $Lists) 
	  { 
		if($List.Hidden -eq $false) 
		{ 
		  $ctx.Load($List) 
		  $ctx.ExecuteQuery() 
		  if($List.WorkflowAssociations.Count -gt 0) 
		  { 
			$DB = New-Object PSObject 
			Add-Member -input $DB noteproperty 'SiteUrl' $site.Url  
			Add-Member -input $DB noteproperty 'Title' $List.Title 
			Add-Member -input $DB noteproperty 'ListType' $List.BaseType 
			Add-Member -input $DB noteproperty 'WorkflowsCount' $List.WorkflowAssociations.Count 
			$Databases += $DB 
			Write-Host $List.Title $List.ItemCount 
		  } 
       
		} 
	  } 
 
	  foreach($web in $ctx.Web.Webs) { 
		$Lists = $web.Lists 
		$ctx.Load($Lists) 
		$ctx.ExecuteQuery() 
		foreach($List in $Lists){ 
		  if($List.Hidden -eq $false){ 
			if($List.ItemCount -gt 100){ 
			  $DB = New-Object PSObject 
			  Add-Member -input $DB noteproperty 'SiteUrl' $web.Url  
			  Add-Member -input $DB noteproperty 'Title' $List.Title 
			  Add-Member -input $DB noteproperty 'ListType' $List.BaseType 
			  Add-Member -input $DB noteproperty 'WorkflowsCount' $List.WorkflowAssociations.Count 
			  $Databases += $DB 
			  Write-Host $List.Title $List.ItemCount 
			} 
		  } 
		} 
	  }
	} 
	Update-control -Synchash $synchash -control IMG_Report_Libraries -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

	#region shared files and Sites
		# Reporting Event
		$message = "Retrieving Files and sites viewable by external users"
		Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
		Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

		# Configure Site URL and User 
		$TenantID = (($VariableHash.tenantname).Split("."))[0]
		$siteURL = "https://$tenantID.sharepoint.com"
		
		# client context object and setting the credentials  
		$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
		$Context.Credentials = $Credentials

		# Calling Search API - Create the instance of KeywordQuery and set the properties 
		$keywordQuery = New-Object Microsoft.SharePoint.Client.Search.Query.KeywordQuery($Context)

		# Formulate the query
		$queryText="ViewableByExternalUsers=true"
		$keywordQuery.QueryText = $queryText
		$keywordQuery.TrimDuplicates=$false
		$keywordQuery.SelectProperties.Add("LastModifiedTime")
		$keywordQuery.SelectProperties.Add("ViewsLifeTime")
		$keywordQuery.SelectProperties.Add("ModifiedBy")
		$keywordQuery.SelectProperties.Add("ViewsLifeTimeUniqueUsers")
		$keywordQuery.SelectProperties.Add("Created")
		$keywordQuery.SelectProperties.Add("CreatedBy")
		$keywordQuery.SortList.Add("ViewsLifeTime","Asc")

		#Search API - Create the instance of SearchExecutor and get the result 
		$searchExecutor = New-Object Microsoft.SharePoint.Client.Search.Query.SearchExecutor($Context)
		$results = $searchExecutor.ExecuteQuery($keywordQuery)
		$Context.ExecuteQuery()

		#CSV file location, to store the result 
		$exportlocation = "C:\Temp\ViewableByExternalUsers.csv"
		foreach($result in $results.Value[0].ResultRows){
			
			$SPO_ExternalShared.Add((
				New-Object PSObject -Property @{
					Title               = $result.Title
					Path                = $result.path
					LifeTimeViews       = $result.viewslifetime
					LifeTimeUniqueUsers = $result.ViewsLifeTimeUniqueUsers
					CreatedBy           = $result.CreatedBy
					Created             = $result.Created
					ModifiedBy          = $result.ModifiedBy
					LastModifyTime      = $result.LastModifiedTime
				}
				))						
		}
		Update-control -Synchash $synchash -control IMG_Report_SPOShared -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
	#endregion shared files and Sites
	
	# Updating "AT A GLANCE" 
	Update-control -Synchash $SyncHash -control txtTotalSiteCollections -property text -value $SitesIncludingPersonal.count

	# Output list of externally shared items
	$SPO_ExternalShared | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - Externally Shared.csv" -NoTypeInformation -Force 

	# Output List of sites
	$SitesIncludingPersonal | Select * | Export-Csv -Path "$($VariableHash.OutputPath)\SPOSitesIncludingPersonal.csv"

	# Output list of site users
	$Databases | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - AllSiteUsers.csv" -NoTypeInformation -Force 

	# Output list of site groups
	$Databases | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - AllSiteGroups.csv" -NoTypeInformation -Force 

	# Output libraries
	$Databases | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - Librariesgt100.csv" -NoTypeInformation -Force
	$Databases | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - Libraries.csv" -NoTypeInformation -Force 

}

Function get-AtAGlance{
	# Reporting Event
	$message = "Running At A Glance"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	# Gathering information
	$Tenant = Get-OrganizationConfig | Select-Object -ExpandProperty Name
	$MSOLCompanyInfo = Get-MsolCompanyInformation
	[System.String]$DirSyncEnabled = $MSOLCompanyInfo | Select-Object -ExpandProperty DirectorySynchronizationEnabled
	$DirSyncLastSync = $MSOLCompanyInfo | Select-Object -ExpandProperty LastDirSyncTime
	$PassSyncLastSync = $MSOLCompanyInfo | Select-Object -ExpandProperty LastPasswordSyncTime
	[System.String]$PassSyncEnabled = $MSOLCompanyInfo | Select-Object -ExpandProperty PasswordSynchronizationEnabled
	[System.String]$TenantDisplayName = $MSOLCompanyInfo | Select-Object -ExpandProperty DisplayName
    [System.String]$TenantCountry = $MSOLCompanyInfo | Select-Object -ExpandProperty CountryLetterCode
    [System.String]$TechContact = $MSOLCompanyInfo | Select-Object -ExpandProperty TechnicalNotificationEmails
    [System.String]$TechContactPhone = $MSOLCompanyInfo | Select-Object -ExpandProperty TelephoneNumber
	$MSOLAccountSKU = Get-MsolAccountSku
    $TotalPlans = ($MSOLAccountSKU).count
    $TotalLicenses = $MSOLAccountSKU | Measure-Object ActiveUnits -Sum | Select-Object -ExpandProperty Sum
    $TotalLicensesAssigned = $MSOLAccountSKU | Measure-Object ConsumedUnits -Sum | Select-Object -ExpandProperty Sum
	$FeaturesRelease = Get-OrganizationConfig | Select-Object -ExpandProperty ReleaseTrack

	$Users = Get-MsolUser -All 
	$UnLicensedUsers = $users | where {$_.IsLicensed -eq "False"}
	$syncedUsers = $users | Where-Object {$_.ImmutableId -ne $null}
	$CloudUsers = $users | Where-Object {$_.ImmutableId -eq $null}

	$Contacts = get-msolcontact -All
	$Guest = $users | Where-Object {$_.UserType -eq "Guest"}
	$groups = Get-MsolGroup -All
	$mailboxes = Get-Mailbox -ResultSize Unlimited
	$SharedMailboxes = $mailboxes | where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox'} | Measure-Object
	$RoomMailboxes = $mailboxes | where-Object { $_.RecipientTypeDetails -eq 'RoomMailbox'} | Measure-Object
	$equipmentMailboxes = $mailboxes  | where-Object { $_.RecipientTypeDetails -eq 'EquipmentMailbox'} | Measure-Object
	
	# Accounting for randomness
	if($DirSyncEnabled -eq $true){
		Update-control -Synchash $SyncHash -control txtDirSyncLastSync -property text -value ($DirSyncLastSync).DateTime
	}Else{
		Update-control -Synchash $SyncHash -control txtDirSyncLastSync -property text -value "N/A"
	}

	if($PassSyncEnabled -eq $true){
		Update-control -Synchash $SyncHash -control txtPassSyncLastSync -property text -value ($PassSyncLastSync).DateTime
	}Else{
		Update-control -Synchash $SyncHash -control txtPassSyncLastSync -property text -value "N/A"
	}

	$VariableHash.tenantname = $Tenant

	# Updating "AT A GLANCE"
	Update-control -Synchash $SyncHash -control txtTenant -property text -value $Tenant
	Update-control -Synchash $SyncHash -control txtOrg -property text -value $TenantDisplayName
	Update-control -Synchash $SyncHash -control txtCountry -property text -value $TenantCountry
	Update-control -Synchash $SyncHash -control txtTechnicalContact -property text -value $TechContact
	Update-control -Synchash $SyncHash -control txtContactPhone -property text -value $TechContactPhone
	Update-control -Synchash $SyncHash -control txtTotalPlans -property text -value $TotalPlans
	Update-control -Synchash $SyncHash -control txtTotalLicense -property text -value $TotalLicenses
	Update-control -Synchash $SyncHash -control txtTotalAssignedLicenses -property text -value $TotalLicensesAssigned
	Update-control -Synchash $SyncHash -control txtPassSyncEnabled -property text -value $PassSyncEnabled
	Update-control -Synchash $SyncHash -control txtFeatureRelease -property text -value $FeaturesRelease

	Update-control -Synchash $SyncHash -control txtTotalUsers -property text -value $Users.count
	Update-control -Synchash $SyncHash -control txtTotalUnlicensedUsers -property text -value $UnLicensedUsers.count
	Update-control -Synchash $SyncHash -control txtTotalSyncedUsers -property text -value $syncedUsers.count
	Update-control -Synchash $SyncHash -control txtTotalCloudUsers -property text -value $CloudUsers.count
	Update-control -Synchash $SyncHash -control txtTotalContacts -property text -value $Contacts.count
	Update-control -Synchash $SyncHash -control txtTotalGuests -property text -value $Guest.count
	Update-control -Synchash $SyncHash -control txtTotalTotalGroups -property text -value $groups.count
	Update-control -Synchash $SyncHash -control txtTotalmailboxes -property text -value $mailboxes.count
	Update-control -Synchash $SyncHash -control txtTotalSharedMailboxes -property text -value $SharedMailboxes.count
	Update-control -Synchash $SyncHash -control txtTotalRooms -property text -value $RoomMailboxes.count
	Update-control -Synchash $SyncHash -control txtTotalEquipment -property text -value $equipmentMailboxes.count
	
}

Function Get-AADInventory{

	# Reporting Event
	$message = "Starting Azure Active Directory User report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	# Variables
	$AADUsers_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]    
	$AADUsers_Observable.Clear()
	$Sku = @{
        "NonLicensed"                        = "User is Not Licensed"
		"O365_BUSINESS_ESSENTIALS"		     = "Office 365 Business Essentials"
	    "O365_BUSINESS_PREMIUM"			     = "Office 365 Business Premium"
	    "DESKLESSPACK"					     = "Office 365 (Plan K1)"
	    "DESKLESSWOFFPACK"				     = "Office 365 (Plan K2)"
	    "LITEPACK"						     = "Office 365 (Plan P1)"
	    "EXCHANGESTANDARD"				     = "Office 365 Exchange Online Only"
	    "STANDARDPACK"					     = "Enterprise Plan E1"
	    "STANDARDWOFFPACK"				     = "Office 365 (Plan E2)"
	    "ENTERPRISEPACK"					 = "Enterprise Plan E3"
	    "ENTERPRISEPACKLRG"				     = "Enterprise Plan E3"
	    "ENTERPRISEWITHSCAL"				 = "Enterprise Plan E4"
	    "STANDARDPACK_STUDENT"			     = "Office 365 (Plan A1) for Students"
	    "STANDARDWOFFPACKPACK_STUDENT"	     = "Office 365 (Plan A2) for Students"
	    "ENTERPRISEPACK_STUDENT"			 = "Office 365 (Plan A3) for Students"
	    "ENTERPRISEWITHSCAL_STUDENT"		 = "Office 365 (Plan A4) for Students"
	    "STANDARDPACK_FACULTY"			     = "Office 365 (Plan A1) for Faculty"
	    "STANDARDWOFFPACKPACK_FACULTY"	     = "Office 365 (Plan A2) for Faculty"
	    "ENTERPRISEPACK_FACULTY"			 = "Office 365 (Plan A3) for Faculty"
	    "ENTERPRISEWITHSCAL_FACULTY"		 = "Office 365 (Plan A4) for Faculty"
	    "ENTERPRISEPACK_B_PILOT"			 = "Office 365 (Enterprise Preview)"
	    "STANDARD_B_PILOT"				     = "Office 365 (Small Business Preview)"
	    "VISIOCLIENT"					     = "Visio Pro Online"
	    "POWER_BI_ADDON"					 = "Office 365 Power BI Addon"
	    "POWER_BI_INDIVIDUAL_USE"		     = "Power BI Individual User"
	    "POWER_BI_STANDALONE"			     = "Power BI Stand Alone"
	    "POWER_BI_STANDARD"				     = "Power-BI Standard"
	    "PROJECTESSENTIALS"				     = "Project Lite"
	    "PROJECTCLIENT"					     = "Project Professional"
	    "PROJECTONLINE_PLAN_1"			     = "Project Online"
	    "PROJECTONLINE_PLAN_2"			     = "Project Online and PRO"
	    "ProjectPremium"					 = "Project Online Premium"
	    "ECAL_SERVICES"					     = "ECAL"
	    "EMS"							     = "Enterprise Mobility Suite"
	    "RIGHTSMANAGEMENT_ADHOC"			 = "Windows Azure Rights Management"
	    "MCOMEETADV"						 = "PSTN conferencing"
	    "SHAREPOINTSTORAGE"				     = "SharePoint storage"
	    "PLANNERSTANDALONE"				     = "Planner Standalone"
	    "CRMIUR"							 = "CMRIUR"
	    "BI_AZURE_P1"					     = "Power BI Reporting and Analytics"
	    "INTUNE_A"						     = "Windows Intune Plan A"
	    "PROJECTWORKMANAGEMENT"			     = "Office 365 Planner Preview"
	    "ATP_ENTERPRISE"					 = "Exchange Online Advanced Threat Protection"
	    "EQUIVIO_ANALYTICS"				     = "Office 365 Advanced eDiscovery"
	    "AAD_BASIC"						     = "Azure Active Directory Basic"
	    "RMS_S_ENTERPRISE"				     = "Azure Active Directory Rights Management"
	    "AAD_PREMIUM"					     = "Azure Active Directory Premium"
	    "MFA_PREMIUM"					     = "Azure Multi-Factor Authentication"
	    "STANDARDPACK_GOV"				     = "Microsoft Office 365 (Plan G1) for Government"
	    "STANDARDWOFFPACK_GOV"			     = "Microsoft Office 365 (Plan G2) for Government"
	    "ENTERPRISEPACK_GOV"				 = "Microsoft Office 365 (Plan G3) for Government"
	    "ENTERPRISEWITHSCAL_GOV"			 = "Microsoft Office 365 (Plan G4) for Government"
	    "DESKLESSPACK_GOV"				     = "Microsoft Office 365 (Plan K1) for Government"
	    "ESKLESSWOFFPACK_GOV"			     = "Microsoft Office 365 (Plan K2) for Government"
	    "EXCHANGESTANDARD_GOV"			     = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"
	    "EXCHANGEENTERPRISE_GOV"			 = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"
	    "SHAREPOINTDESKLESS_GOV"			 = "SharePoint Online Kiosk"
	    "EXCHANGE_S_DESKLESS_GOV"		     = "Exchange Kiosk"
	    "RMS_S_ENTERPRISE_GOV"			     = "Windows Azure Active Directory Rights Management"
	    "OFFICESUBSCRIPTION_GOV"			 = "Office ProPlus"
	    "MCOSTANDARD_GOV"				     = "Lync Plan 2G"
	    "SHAREPOINTWAC_GOV"				     = "Office Online for Government"
	    "SHAREPOINTENTERPRISE_GOV"		     = "SharePoint Plan 2G"
	    "EXCHANGE_S_ENTERPRISE_GOV"		     = "Exchange Plan 2G"
	    "EXCHANGE_S_ARCHIVE_ADDON_GOV"	     = "Exchange Online Archiving"
	    "EXCHANGE_S_DESKLESS"			     = "Exchange Online Kiosk"
	    "SHAREPOINTDESKLESS"				 = "SharePoint Online Kiosk"
	    "SHAREPOINTWAC"					     = "Office Online"
	    "YAMMER_ENTERPRISE"				     = "Yammer for the Starship Enterprise"
	    "EXCHANGE_L_STANDARD"			     = "Exchange Online (Plan 1)"
	    "MCOLITE"						     = "Lync Online (Plan 1)"
	    "SHAREPOINTLITE"					 = "SharePoint Online (Plan 1)"
	    "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ" = "Office ProPlus"
	    "EXCHANGE_S_STANDARD_MIDMARKET"	     = "Exchange Online (Plan 1)"
	    "MCOSTANDARD_MIDMARKET"			     = "Lync Online (Plan 1)"
	    "SHAREPOINTENTERPRISE_MIDMARKET"	 = "SharePoint Online (Plan 1)"
	    "OFFICESUBSCRIPTION"				 = "Office ProPlus"
	    "YAMMER_MIDSIZE"					 = "Yammer"
	    "DYN365_ENTERPRISE_PLAN1"		     = "Dynamics 365 Customer Engagement Plan Enterprise Edition"
	    "ENTERPRISEPREMIUM_NOPSTNCONF"	     = "Enterprise E5 (without Audio Conferencing)"
	    "ENTERPRISEPREMIUM"				     = "Enterprise E5 (with Audio Conferencing)"
	    "MCOSTANDARD"					     = "Skype for Business Online Standalone Plan 2"
	    "PROJECT_MADEIRA_PREVIEW_IW_SKU"	 = "Dynamics 365 for Financials for IWs"
	    "STANDARDWOFFPACK_IW_STUDENT"	     = "Office 365 Education for Students"
	    "STANDARDWOFFPACK_IW_FACULTY"	     = "Office 365 Education for Faculty"
	    "EOP_ENTERPRISE_FACULTY"			 = "Exchange Online Protection for Faculty"
	    "EXCHANGESTANDARD_STUDENT"		     = "Exchange Online (Plan 1) for Students"
	    "OFFICESUBSCRIPTION_STUDENT"		 = "Office ProPlus Student Benefit"
	    "STANDARDWOFFPACK_FACULTY"		     = "Office 365 Education E1 for Faculty"
	    "STANDARDWOFFPACK_STUDENT"		     = "Microsoft Office 365 (Plan A2) for Students"
	    "DYN365_FINANCIALS_BUSINESS_SKU"	 = "Dynamics 365 for Financials Business Edition"
	    "DYN365_FINANCIALS_TEAM_MEMBERS_SKU" = "Dynamics 365 for Team Members Business Edition"
	    "FLOW_FREE"						     = "Microsoft Flow Free"
	    "POWER_BI_PRO"					     = "Power BI Pro"
	    "O365_BUSINESS"					     = "Office 365 Business"
	    "DYN365_ENTERPRISE_SALES"		     = "Dynamics Office 365 Enterprise Sales"
	    "RIGHTSMANAGEMENT"				     = "Rights Management"
	    "PROJECTPROFESSIONAL"			     = "Project Professional"
	    "VISIOONLINE_PLAN1"				     = "Visio Online Plan 1"
	    "EXCHANGEENTERPRISE"				 = "Exchange Online Plan 2"
	    "DYN365_ENTERPRISE_P1_IW"		     = "Dynamics 365 P1 Trial for Information Workers"
	    "DYN365_ENTERPRISE_TEAM_MEMBERS"	 = "Dynamics 365 For Team Members Enterprise Edition"
	    "CRMSTANDARD"					     = "Microsoft Dynamics CRM Online Professional"
	    "EXCHANGEARCHIVE_ADDON"			     = "Exchange Online Archiving For Exchange Online"
	    "EXCHANGEDESKLESS"				     = "Exchange Online Kiosk"
	    "SPZA_IW"						     = "App Connect"
	    "WINDOWS_STORE"					     = "Windows Store for Business"
	    "MCOEV"							     = "Microsoft Phone System"
	    "VIDEO_INTEROP"					     = "Polycom Skype Meeting Video Interop for Skype for Business"
	    "SPE_E5"							 = "Microsoft 365 E5"
	    "SPE_E3"							 = "Microsoft 365 E3"
	    "ATA"							     = "Advanced Threat Analytics"
	    "MCOPSTN2"						     = "Domestic and International Calling Plan"
	    "FLOW_P1"						     = "Microsoft Flow Plan 1"
	    "FLOW_P2"						     = "Microsoft Flow Plan 2"
    }
				
	try{
		$users = get-msoluser -all -ea stop | where{$_.UserPrincipalName -notlike "*#ext#*"} | select DisplayName, FirstName, LastName, UserPrincipalName, Title, Department, Office, PhoneNumber, MobilePhone, CloudAnchor, IsLicensed, @{Name="License"; Expression = {$_.licenses.accountskuid}} 
				
		ForEach ($user in $users) { 
			If (-NOT [System.String]::IsNullOrEmpty($user)) { 

				$Licenses = ((Get-MsolUser -UserPrincipalName $User.UserPrincipalName).Licenses).AccountSkuID

				If (($Licenses).Count -gt 1){
					Foreach ($License in $Licenses){
						$LicenseItem = $License -split ":" | Select-Object -Last 1
						$TextLic = $Sku.Item("$LicenseItem")

						$Object02 = $null
						$Object02 = @()
						$Object01 = New-Object PSObject
						$Object01 | Add-Member -MemberType NoteProperty -Name "License" -Value "$TextLic"
						$Object02 += $NewObject01
					}
				}Else{
					$LicenseItem = ((Get-MsolUser -UserPrincipalName $User.UserPrincipalName).Licenses).AccountSkuID -split ":" | Select-Object -Last 1
					$TextLic = $Sku.Item("$LicenseItem")

						$Object02 = $null
						$Object02 = @()
						$Object01 = New-Object PSObject
						$Object01 | Add-Member -MemberType NoteProperty -Name "License" -Value "$TextLic"
						$Object02 += $NewObject01
				}


				$AADUsers_Observable.Add((
				New-Object PSObject -Property @{
					DisplayName = $user.DisplayName
					FirstName = $user.FirstName
					LastName = $user.LastName	
					UserPrincipalName = $user.UserPrincipalName
					Title = $user.Title
					Department = $user.Department
					Office = $user.Office	
					PhoneNumber = $user.PhoneNumber
					MobilePhone = $user.MobilePhone
					CloudAnchor = $user.CloudAnchor
					IsLicensed = $user.IsLicensed
					Licenses = ($Object02 | Out-String).Trim()										
				}
				))						
			}
		}
	}Catch{
		# Nothing here yet
	}
	Update-control -Synchash $synchash -control IMG_Report_AADUsers -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

	$AADUsers_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - User Report.csv" -NoTypeInformation -Force 
}

Function get-O365UsageReports{
	<#
.Synopsis
	Get-O365UsageReports gather all the Office 365 Usage Reports via Graph (Beta endpoint) and generates an Excel document

.DESCRIPTION
	This PowerShell script requires an Azure Application Client ID which has access to the Microsoft Graph's Read all usage reports permissions to pull the Office 365 Usage Report and save to an Excel document.
    $APIVersion defaults to the Microsoft Graft beta version, as the Teams usage reports are not currently in the v1.0 API.
    See: Manage-AzureAppRegistration: http://realtimeuc.com/2017/12/manage-azureappregistration

.NOTES
	NAME:			Get-O365UsageReports.ps1
    VERSION:      	2.0
    AUTHOR:       	Michael LaMontagne 
    LASTEDIT:     	5/24/2018

V 1.0 - Jan 2018 -	Fast Publish.
V 2.0 - Jan 2018 -	Graph Change, no DLL required.

.LINK
   Website: http://realtimeuc.com
   Twitter: http://www.twitter.com/realtimeuc
   LinkedIn: http://www.linkedin.com/in/mlamontagne/

.EXAMPLE
   $Results = .\Get-O365UsageReports.ps1
   
	Description
	-----------
	Prompts for Azure Tenant AD Domain Name (domain.onmicrosoft.com), prompts for Azure Application Client ID, prompts for credentials 
    before connecting to Microsoft Graph to pull the Office 365 Usage Reports for the last 30 days and saving to an Excel document in c:\temp\O365Reports.xlsx.
    Will also return the Usage Reports as a hashtable in $Results.
	
.EXAMPLE
	$cred = get-credential
    $Results = .\Get-O365UsageReports.ps1 -$AzureAppClientId '7d856782-ba2c-XXXX-a39e-778c33e4ecd4' -Credential $cred -period 'd180' -File 'c:\test\o365.xls' 
   
	Description
	-----------
    Connecting to Microsoft Graph to pull the Office 365 Usage Reports for the last 180 days and saving to an Excel document in c:\test\O365.xlsx.
    Will also return the Usage Reports as a hashtable in $Results.

.EXAMPLE
	$cred = get-credential
    $Results = .\Get-O365UsageReports.ps1 -$AzureAppClientId '7d856782-ba2c-XXXX-a39e-778c33e4ecd4' -Credential $cred -NoExcel
   
	Description
	-----------
    Connecting to Microsoft Graph to pull the Office 365 Usage Reports for the last 30 days and return the Usage Reports as a hashtable in $Results. Excel document output disabled.
    
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] 
    [string]$AzureAppClientId,  #Azure Application Client ID with Microsoft Graph - Read all usage reports permissions, Manage-AzureAppRegistration: http://realtimeuc.com/2017/12/manage-azureappregistration
    [Parameter(Mandatory=$true)]
    [Pscredential]$Credential = $(Get-Credential),
    [ValidateSet('D7','D30','D90','D180')] #Reporting Period in Days. Valid entries:
    [string]$Period = 'D30',
    [switch]$NoExcel, #Switch to prevent Excel export
    [string]$File ='c:\temp\O365Reports.xlsx', #Excel file name
    [string]$APIVersion ='beta' #beta or v1.0
)

$Periods = @('D7','D30','D90','D180')
$Period = $Period.ToUpper()

#Raw data arrays
$objectCollection = @{}

#Request Graph API Token and build request header.
$resourceURL = "https://graph.microsoft.com/" #Resource URI to the Microsoft Graph

function Connect-Graph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [pscredential]$Credential,
        
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$ResourceURL = "https://graph.windows.net/",
        
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientID = '1950a258-227b-4e31-a9cf-717495945fc2'
    )
    $tokenArgs = @{
        grant_type = "password"
        resource   = $ResourceURL
        username   = $Credential.Username
        password   = $Credential.GetNetworkCredential().Password
        client_id  = $ClientID # from msonline extended
    }
    try {
        $token = Invoke-RestMethod -Uri https://login.microsoftonline.com/common/oauth2/token -body $tokenArgs -Method POST
        if($token) {
            # note we don't refresh so this token is only good for maybe 1 hour
            $Script:AadToken = "$($token.token_type) $($token.access_token)"
            $Script:AadHeader = @{
                "Authorization" = $Script:AadToken
                "Content-Type" = "application/json"
            }
            $true
        } else {
            $Script:AadToken = $false
            $Script:AadHeader = $false
            $false
        }
    } catch {
        $false
    }
}

function RestMethod {
    Param (
    [parameter(Mandatory=$true)]
    [ValidateSet("GET","POST","PATCH","DELETE", "PUT")]
    [String]$Method,

    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$URI,

    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    $Headers=$Script:AadHeader,

    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [String]$Body
    )

    $RestResults = $null
   try {
        if ($PSBoundParameters.ContainsKey("Body")) {
            $RestResults = Invoke-RestMethod -Method $Method -Uri $URI -Headers $Headers -Body $Body -Verbose
        }
        else {
            $RestResults = Invoke-RestMethod -Method $Method -Uri $URI -Headers $Headers -Verbose
        }
     
    }
    catch {
        $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd()   
        $Message = $(($responseBody -split('"value":"') )[1] -split('"'))[0] 
        Write-error "$Message" 
        return $Message
    }

    return $RestResults
}



#Graph Usage Reports:
    #https://github.com/microsoftgraph/microsoft-graph-docs/blob/master/api-reference/beta/resources/report.md
    #https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/microsoft_teams_device_usage_reports
    #https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/microsoft_teams_user_activity_reports
$O365Reports = @(
    'getEmailActivityUserDetail';
    'getEmailActivityCounts';
    'getEmailActivityUserCounts';
    'getEmailAppUsageUserDetail';
    'getEmailAppUsageAppsUserCounts';
    'getEmailAppUsageUserCounts';
    'getEmailAppUsageVersionsUserCounts';
    'getMailboxUsageDetail';
    'getMailboxUsageMailboxCounts';
    'getMailboxUsageQuotaStatusMailboxCounts';
    'getMailboxUsageStorage';
    'getOffice365ActivationsUserDetail';
    'getOffice365ActivationCounts';
    'getOffice365ActivationsUserCounts';
    'getOffice365ActiveUserDetail';
    'getOffice365ActiveUserCounts';
    'getOffice365ServicesUserCounts';
    'getOffice365GroupsActivityDetail';
    'getOffice365GroupsActivityCounts';
    'getOffice365GroupsActivityGroupCounts';
    'getOffice365GroupsActivityStorage';
    'getOffice365GroupsActivityFileCounts';
    'getOneDriveActivityUserDetail';
    'getOneDriveActivityUserCounts';
    'getOneDriveActivityFileCounts';
    'getOneDriveUsageAccountDetail';
    'getOneDriveUsageAccountCounts';
    'getOneDriveUsageFileCounts';
    'getOneDriveUsageStorage';
    'getSharePointActivityUserDetail';
    'getSharePointActivityFileCounts';
    'getSharePointActivityUserCounts';
    'getSharePointActivityPages';
    'getSharePointSiteUsageDetail';
    'getSharePointSiteUsageFileCounts';
    'getSharePointSiteUsageSiteCounts';
    'getSharePointSiteUsageStorage';
    'getSharePointSiteUsagePages';
    'getSkypeForBusinessActivityUserDetail';
    'getSkypeForBusinessActivityCounts';
    'getSkypeForBusinessActivityUserCounts';
    'getSkypeForBusinessDeviceUsageUserDetail';
    'getSkypeForBusinessDeviceUsageDistributionUserCounts';
    'getSkypeForBusinessDeviceUsageUserCounts';
    'getSkypeForBusinessOrganizerActivityCounts';
    'getSkypeForBusinessOrganizerActivityUserCounts';
    'getSkypeForBusinessOrganizerActivityMinuteCounts';
    'getSkypeForBusinessParticipantActivityCounts';
    'getSkypeForBusinessParticipantActivityUserCounts';
    'getSkypeForBusinessParticipantActivityMinuteCounts';
    'getSkypeForBusinessPeerToPeerActivityCounts';
    'getSkypeForBusinessPeerToPeerActivityUserCounts';
    'getSkypeForBusinessPeerToPeerActivityMinuteCounts';
    'getteamsDeviceUsageUserDetail';
    'getteamsDeviceUsageUserCounts';
    'getteamsDeviceUsagedistributionUserCounts';
    'getteamsUserActivityUserDetail';
    'getteamsUserActivityCounts';
    'getteamsUserActivityUserCounts';
    'getYammerActivityUserDetail';
    'getYammerActivityCounts';
    'getYammerActivityUserCounts';
    'getYammerDeviceUsageUserDetail';
    'getYammerDeviceUsageDistributionUserCounts';
    'getYammerDeviceUsageUserCounts';
    'getYammerGroupsActivityDetail';
    'getYammerGroupsActivityGroupCounts';
    'getYammerGroupsActivityCounts'
)

#Get Graph Token
$connect = Connect-Graph $credential $resourceURL $AzureAppClientId

#Data gathering via Graph
if($connect){   
    foreach ($Report in $O365Reports){
        $Results = $null
        $Request = $null
        $ReportName = $null

        if($Periods -notcontains $Period -or $report -like "getOffice365Activation*"){
            $Request = "https://graph.microsoft.com/$($APIVersion)/reports/$($Report)"  
        }
        else{
            $Request = "https://graph.microsoft.com/$($APIVersion)/reports/$($Report)(period='$($Period)')"   
        }

        $Results = RestMethod -Method "Get" -URI $Request       
    
        #Shorten report name due to Excel limits
        $ReportName = $Report.ToLower()
		$ReportName = $ReportName.Replace("get","")
    
        if($Results){
            $Results = $Results.replace("﻿","") | ConvertFrom-Csv
            if($Results){
                $objectCollection.Add($($ReportName),$Results)    
            }
            else{
               $objectCollection.Add($($ReportName),$objResults)  
            }
        }    
    }
}

return $objectCollection
}

Function Connect-ProvisioningWebServiceAPI{
	<#
		.SYNOPSIS
			Connects to the Office 365 provisioning web service API.

		.DESCRIPTION
			Connects to the Office 365 provisioning web service API.
			
			If a credential is specified, it will be used to establish a connection with the provisioning
			web service API.
			
			If a credential is not specified, an attempt is made to identify an existing connection to
			the provisioning web service API.  If an existing connection is identified, the existing
			connection is used.  If an existing connection is not identified, the user is prompted for
			credentials so that a new connection can be established.

		.PARAMETER Credential
			Specifies the credential to use when connecting to the provisioning web service API
			using Connect-MsolService.

		.EXAMPLE
			PS> ConnectProvisioningWebServiceAPI

		.EXAMPLE
			PS> ConnectProvisioningWebServiceAPI -Credential
			
		.INPUTS
			[System.Management.Automation.PsCredential]

		.OUTPUTS

		.NOTES

	#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $False)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	# if a credential was supplied, assume a new connection is intended and create a new
	# connection using specified credential
	If ($Credential)
	{
		If ((!$Credential) -or (!$Credential.Username) -or ($Credential.Password.Length -eq 0))
		{
			Write-warning -Message ("Invalid credential.  Please verify the credential and try again.")
			Exit
		}
		
		# connect to provisioning web service api
		Write-Verbose -Message "Connecting to the Office 365 provisioning web service API.  Please wait..."
		Connect-MsolService -Credential $Credential
		If($? -eq $False){WriteConsoleMessage -Message "Error while connecting to the Office 365 provisioning web service API.  Quiting..." -MessageType "Error";Exit}
	}
	Else
	{
		Write-Verbose -Message "Attempting to identify an open connection to the Office 365 provisioning web service API.  Please wait..." 
		$getMsolCompanyInformationResults = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
		If (!$getMsolCompanyInformationResults)
		{
			Write-Verbose -Message "Could not identify an open connection to the Office 365 provisioning web service API." 			If (!$Credential)
			{
				$Credential = $Host.UI.PromptForCredential("Enter Credential",
					"Enter the username and password of an Office 365 administrator account.",
					"",
					"userCreds")
			}
			If ((!$Credential) -or (!$Credential.Username) -or ($Credential.Password.Length -eq 0))
			{
				Write-Verbose -Message ("Invalid credential.  Please verify the credential and try again.")
				Exit
			}
			
			# connect to provisioning web service api
			Write-Verbose -Message "Connecting to the Office 365 provisioning web service API.  Please wait..."
			Connect-MsolService -Credential $Credential
			If($? -eq $False){WriteConsoleMessage -Message "Error while connecting to the Office 365 provisioning web service API.  Quiting..." -MessageType "Error";Exit}
			$getMsolCompanyInformationResults = Get-MsolCompanyInformation -ErrorAction SilentlyContinue
			WriteConsoleMessage -Message ("Connected to Office 365 tenant named: `"{0}`"." -f $getMsolCompanyInformationResults.DisplayName) -MessageType "Information"
		}
		Else
		{
			Write-Warning -Message ("Connected to Office 365 tenant named: `"{0}`"." -f $getMsolCompanyInformationResults.DisplayName) 
		}
	}
	If (!$Script:Credential) {$Script:Credential = $Credential}
}

Function get-M365LicenseUsage{

	# Reporting Event
	$message = "Running licensing report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	Connect-ProvisioningWebServiceAPI -Credential $Credential

	# get Office 365 SKU info
	WriteConsoleMessage -Message "Getting SKU information.  Please wait..." -MessageType "Information"
	$getMsolAccountSkuResults = Get-MsolAccountSku

	# iterate through the sku results
	WriteConsoleMessage -Message "Processing SKU results.  Please wait..." -MessageType "Information"
	$arrSkuData = @()
	foreach($sku in $getMsolAccountSkuResults)
	{
		$objSkuData = New-Object PSObject
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "AccountSkuId" -Value $sku.accountskuid
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "ActiveUnits" -Value $sku.activeunits
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "ConsumedUnits" -Value $sku.consumedunits
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "AvailableUnits" -Value ($($sku.activeunits - $sku.consumedunits) | Out-String).Trim()
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "WarningUnits" -Value $sku.warningunits
		Add-Member -InputObject $objSkuData -MemberType NoteProperty -Name "SuspendedUnits" -Value $sku.suspendedunits
		$arrSkuData += $objSkuData
	}
	Update-control -Synchash $synchash -control IMG_Report_LicenseUsage -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
	
	$arrSkuData | Export-Csv -Path "$($VariableHash.OutputPath)\M365 - License Usage Report.csv" -NoTypeInformation -Force 
}

Function get-AzureInventory{

	# Creating credentials object
	$secpswd = ConvertTo-SecureString $VariableHash.M365password -AsPlainText -Force
	$Creds = new-object -typename System.Management.Automation.PSCredential -argumentlist $VariableHash.M365Username,$secpswd

	# Variables
	$objVirtualMachine = @()

	# Login to Azure RM
	Login-AzureRmAccount -Credential $Creds

	# Retrieving subscription list
	$Subscriptions = Get-AzureRmSubscription

	Foreach($subscription in $Subscriptions){
		# Setting subscription
		Select-AzureRmSubscription -Subscription $subscription

		#region Azure VMs
		$AzureVirtualMachines = New-Object System.Collections.ObjectModel.ObservableCollection[object]
		$AzureVirtualMachines.Clear()

		Try{
			# Reporting
			$Timestamp = (get-date -Format HH:mm:ss)
			Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Azure Virtual Machine report" -AppendContent
		
			$AzureVMS = get-AzureVM | Get-AzureVM | Format-List DeploymentName,Name,Label,VM,InstanceStatus,IpAddress,
			InstanceStateDetails,PowerState,InstanceErrorCode,InstanceFaultDomain,InstanceName,InstanceUpgradeDomain,
			InstanceSize,AvailabilitySetName,DNSName,ServiceName,OperationDescription,OperationId,OperationStatus

			Foreach($azureVM in $AzureVMS){
				If(-NOT [System.String]::IsNullOrEmpty($azureVM)){
					$AzureVirtualMachines.Add((
						New-Object PSObject -Property @{
							Name            = $AzureVM.Name
							DNSName		    = $AzureVM.DNSName
							IP              = $AzureVM.IPAddress
							State           = $AzureVM.PowerState
							Label           = $AzureVM.Label
							Size            = $AzureVM.InstanceSize
							OperationStatus = $AzureVM.OperationStatus
							OperationDescr  = $AzureVM.OperationDescription
							AvailabilitySet = $AzureVM.AvailabilitySetName
							Instance        = $AzureVM.InstanceName
							InstanceStatus  = $AzureVM.InstanceStatus
							FaultDomain     = $AzureVM.InstanceFaultDomain
							UpgradeDomain   = $AzureVM.InstanceUpgradeDomain
						}
					))
				}
			}
		}Catch{

		}
		
		Update-control -Synchash $synchash -control IMG_Report_AzureVM -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
		
		$AzureVirtualMachines | Export-Csv -Path "$($VariableHash.OutputPath)\Azure - Virtual Machines.csv" -NoTypeInformation -Force 
		#endregion Azure VMs
	}

}

Function get-AADDeletedUsers{
	$AADDeletedUsers_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$AADDeletedUsers_Observable.Clear()
			
	# Reporting Event
	$message = "Starting Azure Active Directory Deleted users report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	# Body
	try{
		$users = Get-MsolUser -All -ReturnDeletedUsers | select SignInName, UserPrincipalName, DisplayName, SoftDeletionTimestamp, IsLicensed, @{Name="License"; Expression = {$_.licenses.accountskuid}} 
				
		ForEach ($user in $users) { 
			If (-NOT [System.String]::IsNullOrEmpty($user)) {  
				$AADDeletedUsers_Observable.Add((
					New-Object PSObject -Property @{
						SignInName = $user.SignInName
						UserPrincipalName = $user.UserPrincipalName
						DisplayName = $user.DisplayName
						SoftDeletionTimestamp = $user.SoftDeletionTimestamp
						IsLicensed = $user.IsLicensed
						Licenses = $user.License
					}
				))   					
			}
		}
				
		Update-control -Synchash $synchash -control IMG_Report_AADDelusers -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
		
		# export
		$AADDeletedUsers_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - Deleted User Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_AADDelUser -property Source -Value "$($VariableHash.IconsDir)\Light.red.ico"
	}		
}

Function get-AADContacts{
	$AADContacts_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$AADContacts_Observable.Clear()
			
	# Reporting Event
	$message = "Starting Azure Active Directory contacts report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	try{
		$Contacts = Get-Msolcontact -all | select DisplayName, EmailAddress
				
		ForEach ($Contact in $Contacts) { 
			If (-NOT [System.String]::IsNullOrEmpty($Contact)) {  
				$AADContacts_Observable.Add((
					New-Object PSObject -Property @{
						DisplayName = $Contact.DisplayName
						EmailAddress = $Contact.EmailAddress
					}
				))   					
			}
		}
				
		Update-control -Synchash $synchash -control IMG_Report_AADContacts -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
		
		# Export
		$AADContacts_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - Contacts Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_AADContacts -property Source -Value "$($VariableHash.IconsDir)\Light.red.ico"
	}
}

Function get-AADGroups{
	$AADGroups_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$AADGroups_Observable.Clear()
			
	# Reporting Event
	$message = "Running Azure Active Directory Groups report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	try{
		$groups = get-msolGroup | select DisplayName, EmailAddress, GroupType, ValidationStatus
				
		ForEach ($group in $groups) { 
			If (-NOT [System.String]::IsNullOrEmpty($group)) {  
				$AADGroups_Observable.Add((
					New-Object PSObject -Property @{
						GroupType = $group.GroupType
						DisplayName = $group.DisplayName
						MailAddress = $group.EmailAddress
						ValidationStatus = $group.ValidationStatus
					}
				))   					
			}
		}
		
		Update-control -Synchash $synchash -control IMG_Report_AADGroups -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

		# Export
		$AADGroups_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - Groups Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_AADGroups -property Source -Value "$($VariableHash.IconsDir)\Light.red.ico"
	}
}

Function get-AADDomains{
	$AADDomains_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$AADDomains_Observable.Clear()
			
	# Reporting Event
	$message = "Running Azure Active Directory Domains Report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	try{
		$domains = Get-MsolDomain | select Name, Status, Authentications
				
		ForEach ($domain in $Domains) { 
			If (-NOT [System.String]::IsNullOrEmpty($domain)) {  
				$AADDomains_Observable.Add((
					New-Object PSObject -Property @{
						Name = $domain.Name
						Status = $domain.Status
						Authentications = $domain.Authentications
					}
				))   					
			}
		}
				
		Update-control -Synchash $synchash -control IMG_Report_AADDomains -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
		
		# Export
		$AADDomains_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - Domains Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_AADDomains -property Source -Value "$($VariableHash.IconsDir)\Light.red.ico"
	}		
}

Function get-ExchangeMailboxes{
	$ExchangeMailboxes_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$ExchangeMailboxes_Observable.Clear()
			
	# Reporting Event
	$message = "Running Exchange Online mailboxes report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	try{
		$ExchangeMailboxes = Get-Mailbox | sort DisplayName | select DisplayName, Alias, PrimarySMTPAddress, ArchiveStatus, UsageLocation, WhenMailboxCreated, UserPrincipalName, RecipientTypeDetails, AuditEnabled, IsDirSynced, IsShared
		
		ForEach ($ExchangeMailbox in $ExchangeMailboxes) { 
			If (-NOT [System.String]::IsNullOrEmpty($ExchangeMailbox)) { 
				$statistics = Get-MailboxStatistics $ExchangeMailbox.alias -WarningAction:SilentlyContinue| select ItemCount, TotalItemSize, LastLogonTime
				$ExchangeMailboxes_Observable.Add((
					New-Object PSObject -Property @{
						DisplayName = $ExchangeMailbox.DisplayName
						Alias = $ExchangeMailbox.Alias
						PrimarySMTPAddress = $ExchangeMailbox.PrimarySMTPAddress
						ItemCount = $statistics.ItemCount
						TotalItemSize = $statistics.TotalItemSize
						ArchiveStatus = $ExchangeMailbox.ArchiveStatus
						UsageLocation = $ExchangeMailbox.UsageLocation
						WhenMailboxCreated = $ExchangeMailbox.WhenMailboxCreated
						UserPrincipalName = $ExchangeMailbox.UserPrincipalName
						AuditEnabled = $ExchangeMailbox.AuditEnabled
						RecipientTypeDetails = $ExchangeMailbox.RecipientTypeDetails
						IsDirSynced = $ExchangeMailbox.IsDirSynced
						IsShared = $ExchangeMailbox.IsShared
						LastLogonTime = $statistics.LastLogonTime
					}
				))   					
			}
		}
				
		Update-control -Synchash $synchash -control IMG_Report_EXOMailboxes -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
		
		# Export
		Update-control -Synchash $SyncHash -control DataGrid_EXOMailboxes -property ItemsSource -value $ExchangeMailboxes_Observable
		$ExchangeMailboxes_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\EXO - Mailboxes Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_EXOMailboxes -property Source -Value "$($VariableHash.IconsDir)\Light.red.ico"
	}			
}

Function get-ExchangeArchives{
	$ExchangeArchives_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$ExchangeArchives_Observable.Clear()
			
	# Reporting Event
	$message = "Running Exchange Online Archives report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	try{
		$ExchangeArchives = Get-Mailbox -Archive | sort DisplayName | select DisplayName, Alias, PrimarySMTPAddress, ArchiveStatus, UsageLocation, WhenMailboxCreated
				
		ForEach ($ExchangeArchive in $ExchangeArchives) { 
			If (-NOT [System.String]::IsNullOrEmpty($ExchangeArchive)) {
				$statistics = Get-MailboxStatistics $ExchangeArchive.alias -archive -WarningAction:SilentlyContinue| select ItemCount, TotalItemSize, LastLogonTime
				$ExchangeArchives_Observable.Add((
					New-Object PSObject -Property @{
						DisplayName = $ExchangeArchive.DisplayName
						Alias = $ExchangeArchive.Alias
						PrimarySMTPAddress = $ExchangeArchive.PrimarySMTPAddress
						ItemCount = $statistics.ItemCount
						TotalItemSize = $statistics.TotalItemSize
						ArchiveStatus = $ExchangeArchive.ArchiveStatus
						UsageLocation = $ExchangeArchive.UsageLocation
						WhenMailboxCreated = $ExchangeArchive.WhenMailboxCreated
						LastLogonTime = $statistics.LastLogonTime
					}
				))   					
			}
		}

		Update-control -Synchash $synchash -control IMG_Report_EXOArchives -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
		
		# Export
		$ExchangeArchives_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\EXO - Archives Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_EXOArchives -property Source -Value "$($VariableHash.IconsDir)\Light.red.ico"
	}		
}

Function get-ExchangeGroups{
	$ExchangeGroups_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$ExchangeGroups_Observable.Clear()
			
	# Reporting Event
	$message = "Running Exchange Online Groups report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
	try{
		$ExchangeGroups = Get-Group | where{$_.RecipientTypeDetails -ne "RoleGroup"} | sort DisplayName | select DisplayName, RecipientTypeDetails, @{Name="Owner"; Expression = {$_.ManagedBy}}, WindowsEmailAddress
				
		ForEach ($ExchangeGroup in $ExchangeGroups) { 
			If (-NOT [System.String]::IsNullOrEmpty($ExchangeGroup)) {
				$ExchangeGroups_Observable.Add((
					New-Object PSObject -Property @{
						DisplayName = $ExchangeGroup.DisplayName
						RecipientTypeDetails = $ExchangeGroup.RecipientTypeDetails
						Owner = $ExchangeGroup.Owner
						WindowsEmailAddress = $ExchangeGroup.WindowsEmailAddress
					}
				))   					
			}
		}
		Update-control -Synchash $synchash -control IMG_Report_EXOGroups -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

		# Export
		$ExchangeGroups_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\EXO - Groups Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_EXOGroups -property Source -Value "$($VariableHash.IconsDir)\Light.red.ico"	
	}		
}

Function get-ExternalUsers{
	# Reporting Event
	$message = "Running External users report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	$ExternalUsers = Get-MsolUser -All | ? {$_.UserType -eq "Guest"} | Select DisplayName,SignInName | FT

	Foreach($user in $users){
		$ExternalUsers_Observable.Add((
			New-Object PSObject -Property @{
				DisplayName = $user.DisplayName
				SignInName = $user.SignInName
			}
		))   					
	}

	Update-control -Synchash $synchash -control IMG_Report_ExternalUsers -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

	# Export
	$ExternalUsers_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - External users.csv" -NoTypeInformation -Force
}

function Get-Flows{
	# Reporting Event
	$message = "Running Microsoft Flow report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	$Username = $variablehash.M365username
	$password = $variablehash.M365password
	$SecurePass = ConvertTo-SecureString $password -AsPlainText -Force
	clear-variable password
	Add-PowerAppsAccount -Username $Username -Password $SecurePass

	$flows = get-flow | select DisplayName, Enabled

	Foreach ($flow in $flows){
		$Flows_Observable.Add((
			New-Object PSObject -Property @{
				DisplayName = $flow.DisplayName
				Enabled = $flow.Enabled
			}
		))   			
	}

	$AdminFlows = Get-AdminFlow

	Foreach ($adminflow in $AdminFlows){
		$Flows_Observable.Add((
			New-Object PSObject -Property @{
				DisplayName = $adminflow.DisplayName
				Enabled = $adminflow.Enabled
			}
		))   			
	}

	Update-control -Synchash $synchash -control IMG_Report_Flow -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

	# Export
	$ExternalUsers_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\Microsoft Flow.csv" -NoTypeInformation -Force
}

Function get-powerapps{
	# Reporting Event
	$message = "Running Microsoft PowerApps report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	$Username = $variablehash.M365username
	$password = $variablehash.M365password
	$SecurePass = ConvertTo-SecureString $password -AsPlainText -Force
	clear-variable password
	Add-PowerAppsAccount -Username $Username -Password $SecurePass

	$powerapps = get-adminpowerapp

	foreach($app in $powerapps){
		$Apps_Observable.Add((
			New-Object PSObject -Property @{
				DisplayName    = $app.Internal.Properties.DisplayName
				Description    = $app.Internal.Properties.Description
				CreatedBy      = $app.Internal.Properties.CreatedBy.userPrincipalName
				SharedGroups   = $app.Internal.Properties.sharedGroupsCount
				SharedAccounts = $app.Internal.Properties.sharedUsersCount
				FeaturedApp    = $app.Internal.Properties.isFeaturedApp
				ConsentBypass  = $app.Internal.Properties.bypassConsent
				WebLink        = $app.Internal.Properties.appPackageDetails.webPackage.value


			}
		))   			
	}

	Update-control -Synchash $synchash -control IMG_Report_PowerApps -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

	# Export
	$Apps_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\Microsoft Powerapps.csv" -NoTypeInformation -Force
}

Function get-PowerBIWorkspaces{
	# Reporting Event
	$message = "Running Microsoft PowerBI Workspaces report"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	$Username = $variablehash.M365username
	$password = $variablehash.M365password
	$SecurePass = ConvertTo-SecureString $password -AsPlainText -Force
	$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $Username,$SecurePass
	clear-variable password

	# Connecting
	connect-PowerBIServiceAccount -credential $credentials

	#region Retrieving all workspaces
	$workspaces = Get-PowerBIWorkspace -Scope Organization -All

	Foreach($workspace in $workspaces){
		$PowerBI_Workspaces.Add((
			New-Object PSObject -Property @{
				Name              = $workspace.Name
				Type              = $workspace.Type
				State             = $workspace.State
				ReadOnly          = $workspace.IsReadOnly
				Orphaned          = $workspace.IsOrphaned
				DedicatedCapacity = $workspace.IsOnDedicatedCapacity
				Users             = $workspace.Users
			}
		))   	
	}

	Update-control -Synchash $synchash -control IMG_Report_PowerBIWorkspaces -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
	#endregion

	#region dashboards
		# Reporting Event
		$message = "Running Microsoft PowerBI Dashboards report"
		Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
		Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

		$dashboards = Get-PowerBIDashboard -Scope Organization

		foreach($dash in $dashboards){
			$PowerBI_DashBoards.Add((
				New-Object PSObject -Property @{
					Name       = $dash.Name
					ReadOnly   = $dash.IsReadOnly
					ID         = $dash.Id
				}
			))   	
		}
	
		Update-control -Synchash $synchash -control IMG_Report_PowerBIDashboards -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
	#endregion

	#region reports
		# Reporting Event
		$message = "Running Microsoft PowerBI reports report"
		Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
		Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

		$reports = Get-PowerBIDashboard -Scope Organization

		foreach($report in $reports){
			$PowerBI_Reports.Add((
				New-Object PSObject -Property @{
					Name       = $report.Name
					ID         = $report.Id
					DataSetID  = $reports.DataSetID
					WebURL     = $report.WebURL
				}
			))   	
		}
	
		Update-control -Synchash $synchash -control IMG_Report_PowerBIReports -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"
	#endregion

	#region Datasources
		# Reporting Event
		$message = "Running Microsoft PowerBI datasources report"
		Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
		Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

		$Datasets = Get-PowerBIDataset -Scope Organization

		Foreach($Set in $datasets){
			$PowerBI_Datasets.Add((
				New-Object PSObject -Property @{
					Name                           = $Set.Name
					ConfiguredBy                   = $Set.ConfiguredBy
					RetentionPolicy                = $Set.DefaultRetentionPolicy
					AddRowsAPI                     = $Set.AddRowsApiEnabled
					Tables                         = $Set.Tables
					WebURL                         = $Set.WebUrl
					Relationships                  = $Set.Relationships
					Datasources                    = $Set.Datasources
					DefaultMode                    = $Set.DefaultMode
					Refreshable                    = $Set.IsRefreshable
					EffectiveIdentityRequired      = $Set.IsEffectiveIdentityRequired
					EffectiveIdentityRolesRequired = $Set.IsEffectiveIdentityRolesRequired
					OnPremGatewayRequired          = $Set.IsOnPremGatewayRequired
				}
			))   	
		}

		Update-control -Synchash $synchash -control IMG_Report_PowerBIDatasources -property Source -Value "$($VariableHash.IconsDir)\Light.green.ico"

	#endregion

	# Export
	$PowerBI_Workspaces | Export-Csv -Path "$($VariableHash.OutputPath)\PowerBI - Workspaces.csv" -NoTypeInformation -Force
	$PowerBI_DashBoards | Export-Csv -Path "$($VariableHash.OutputPath)\PowerBI - Dashboards.csv" -NoTypeInformation -Force
	$PowerBI_Reports | Export-Csv -Path "$($VariableHash.OutputPath)\PowerBI - Reports.csv" -NoTypeInformation -Force
	$PowerBI_Datasets | Export-Csv -Path "$($VariableHash.OutputPath)\PowerBI - DataSets.csv" -NoTypeInformation -Force
}

Function Run-Reports{
	# Reporting Event
	$message = "Grabbing information"
	Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
	Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile

	# Variables
	$VariableHash.tenantname = $SyncHash.txt_Home_Tenant.text
	$VariableHash.GraphAppID = $syncHash.txt_Home_GraphAppID.Text
	$VariableHash.TenantRegion = $SyncHash.DD_Home_Region.CurrentItem
	$VariableHash.M365username = $synchash.txt_Home_Username.Text
	$VariableHash.M365password = $synchash.txt_Home_Password.Password
	
	If($VariableHash.TenantRegion -like ""){
		$VariableHash.TenantRegion = "Default"
	}

	$Report_Runspace =[runspacefactory]::CreateRunspace()
	$SyncHash.ReportRunspace = $Report_Runspace
	$Report_Runspace.ApartmentState = "STA"
	$Report_Runspace.ThreadOptions = "ReuseThread"
	$Report_Runspace.Open()

	# Passing variables
	$Report_Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
	$Report_Runspace.SessionStateProxy.SetVariable("VariableHash",$VariableHash)
					
	# Create powershell object which will containt the code we're running in the runspace
	$psCmd = [PowerShell]::Create()

	# Add runspace to Powershell object
	$psCmd.Runspace = $Report_Runspace

	[Void]$psCmd.AddScript({
		# Importing module
		Import-Module "$($VariableHash.ModuleDir)M365InventoryAutomaton.psm1" -Force -DisableNameChecking

		If($VariableHash.M365Username){

		
			# Creating credentials object
			$secpswd = ConvertTo-SecureString $VariableHash.M365password -AsPlainText -Force
			$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $VariableHash.M365Username,$secpswd

			# Calling connection function
			Connect-M365 -Credentials $credentials -Region $VariableHash.TenantRegion -TenantName $VariableHash.tenantname -ExchangeOnline -AzureAD -MSOL -SharePointOnline

			# At A Glance
			get-AtAGlance

			# Azure Active Directory User Report
			get-AADInventory

			# SPO site inventories
			get-SPOInventory

			#region Graph reports
			# Reporting Event
			$message = "Running Graph reports"
			Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
			Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
			
			$Graphreports = get-O365UsageReports -AzureAppClientId $variableHash.GraphAppID -Credential $credentials -Period D180 -NoExcel

			Foreach($item in $Graphreports.keys){
				$Graphreports[$item] | Export-Csv -Path "$($VariableHash.OutputPath)\Graph - $item.csv" -NoTypeInformation
			}

			Update-control -Synchash $synchash -control IMG_Report_Graph -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
			#endregion Graph reports

			# Deleted User reports
			get-AADDeletedUsers

			# AAD Contact report
			get-AADContacts

			# AAD Groups
			get-AADGroups

			# AAD Domains
			get-AADDomains

			# Exchange mailboxes
			get-ExchangeMailboxes

			# Exchange Archives
			get-ExchangeArchives

			# Exchange Groups
			get-ExchangeGroups

			# License usage report
			get-M365LicenseUsage

			# Azure Inventory
			get-AzureInventory

			# Flows
			get-flows

			# Ppowerapps
			get-powerapps

			# PowerBI Workpaces
			get-PowerBIWorkspaces

			# Disconnecting exchange (Pesky little session limit...)
			
			# Reporting event
			$message = "Disconnecting Connections"
			Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
			Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
			Remove-PSSession $Session
			Update-control -Synchash $synchash -control IMG_Conn_EXO -property Source -value "$($VariableHash.IconsDir)\Check_Waiting.ico"
						
			# Reporting event
			$message = "Report run completed"
			Update-control -Synchash $SyncHash -control txt_output -property text -value $message -AppendContent
			Invoke-logging -loglevel INFO -message $message -Runlog $VariableHash.logFile
	
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

	# Begin!
	$data = $psCmd.BeginInvoke()
}

Function Prereq-Failures{
	$Source = -join($VariableHash.IconsDir,"\appbar.warning.png")
	$Image = New-Object System.Windows.Controls.Image
	$Image.Source = $Source
	$Image.Height = [System.Drawing.Image]::FromFile($Source).Height
	$Image.Width = [System.Drawing.Image]::FromFile($Source).Width
	$Image.Margin = 5
					 
	$TextBlock = New-Object System.Windows.Controls.TextBlock
	$TextBlock.Text = "Prerequisite failures detected!`nPlease correct these error and relaunch the tooling!"
	$TextBlock.Padding = 10
	$TextBlock.FontFamily = "Verdana"
	$TextBlock.FontSize = 16
	$TextBlock.TextWrapping = "Wrap"
	$TextBlock.Width = 350
	$TextBlock.VerticalContentAlignment = "Center"
									
	$StackPanel = New-Object System.Windows.Controls.StackPanel
	$StackPanel.Orientation = "Horizontal"
	$StackPanel.Width = 400
	$StackPanel.AddChild($Image)
	$StackPanel.AddChild($TextBlock)
					
	Invoke-WPFMessageBox -Content $StackPanel -Title "WARNING!" -TitleBackground "Orange" -TitleTextForeground "Black" -TitleFontSize "20" -ButtonType OK
}

Function Prereq-Testing{
	#region IsAdmin testing
					
	Update-control -Synchash $PrereqHash -control Lbl_Output -property Content -value "Testing user context..."
	# Retrieving Windows Security Principal
	$ThisPrincipal = new-object System.Security.principal.windowsprincipal( [System.Security.Principal.WindowsIdentity]::GetCurrent())
						
	# Checking if the user in in the Administrator Role
	$IsAdmin = $ThisPrincipal.IsInRole("Administrators")

	If($IsAdmin){
		Update-control -Synchash $PrereqHash -control IMG_Prereq_Admin -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
	}Else{
		Update-control -Synchash $PrereqHash -control IMG_Prereq_Admin -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
		$PrereqHash.PrereqFailed = $true
		$PrereqHash.RunAsAdmin = "Failed"
	}
	#endregion IsAdmin testing

	#region Connectivity Check	
	Update-control -Synchash $PrereqHash -control Lbl_Output -property Content -value "Testing internet connectivity state..."
	
	# getting internet connectivity state
	$HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)

	If($HasInternetAccess){
		Update-control -Synchash $PrereqHash -control IMG_Prereq_Internet -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
	}Else{
		Update-control -Synchash $PrereqHash -control IMG_Prereq_Internet -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
		$PrereqHash.PrereqFailed = $true
		$PrereqHash.InternetAccess = "Failed"
	}
	#endregion Connectivity Check

				#region Sharepoint Online module
					
					Update-control -Synchash $PrereqHash -control Lbl_Output -property Content -value "Testing for SharePoint Online Powershell Module..."

					If((Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\Microsoft.Online.SharePoint.PowerShell\' -Recurse -Filter "Microsoft.Online.SharePoint.PowerShell.psd1")){
						Update-control -Synchash $PrereqHash -control IMG_Prereq_SPOMgmtShell -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
					}Else{
						Update-control -Synchash $PrereqHash -control IMG_Prereq_SPOMgmtShell -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
						$PrereqHash.PrereqFailed = $true
						$PrereqHash.SPOModule = "Failed"
					}
				#endregion Sharepoint Online module

				#region Azure Online module
					
					Update-control -Synchash $PrereqHash -control Lbl_Output -property Content -value "Testing for Azure Active Directory Module..."

					If((Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\AzureAD\' -Recurse -Filter "AzureAD.psd1")){
						Update-control -Synchash $PrereqHash -control IMG_Prereq_AADModule -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
					}Else{
						Update-control -Synchash $PrereqHash -control IMG_Prereq_AADModule -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
						$PrereqHash.PrereqFailed = $true
						$PrereqHash.AADModule = "Failed"
					}
				#endregion Azure Online module

				#region MSTeams Online module
					
					Update-control -Synchash $PrereqHash -control Lbl_Output -property Content -value "Testing for Microsoft Teams Module..."

					If((get-childitem -path 'C:\Program Files\WindowsPowerShell\Modules\MicrosoftTeams' -Recurse -Filter "MicrosoftTeams.psd1")){
						Update-control -Synchash $PrereqHash -control IMG_Prereq_MSTeamsModule -property Source -value "$($VariableHash.IconsDir)\Light.green.ico"
					}Else{
						Update-control -Synchash $PrereqHash -control IMG_Prereq_MSTeamsModule -property Source -value "$($VariableHash.IconsDir)\Light.red.ico"
						$PrereqHash.PrereqFailed = $true
						$PrereqHash.TeamsModule = "Failed"
					}
				#endregion MSTeams Online module

				# Warn
				If($PrereqHash.PrereqFailed -like $true){
					Update-control -Synchash $PrereqHash -control Lbl_Output -property Content -value "Prerequisites failure(s) detected!"
					Prereq-Failures
				}Else{
					Update-control -Synchash $PrereqHash -control Lbl_Output -property Content -value "Prerequisites tested successfully"
				}
}

function Invoke-ColorOutput{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False,Position=1,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[alias('message')]
		[alias('msg')]
		[Object]$Object,
        [Parameter(Mandatory=$False,Position=2,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
		[alias('fore')]
		[ConsoleColor] $ForegroundColor,
        [Parameter(Mandatory=$False,Position=3,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
		[ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')]
		[alias('back')]
		[alias('BGR')]
		[ConsoleColor] $BackgroundColor,
        [Switch]$NoNewline
    )    

    # Save previous colors
    $previousForegroundColor = $host.UI.RawUI.ForegroundColor
    $previousBackgroundColor = $host.UI.RawUI.BackgroundColor

    # Set BackgroundColor if available
    if($BackgroundColor -ne $null)
    { 
       $host.UI.RawUI.BackgroundColor = $BackgroundColor
    }

    # Set $ForegroundColor if available
    if($ForegroundColor -ne $null)
    {
        $host.UI.RawUI.ForegroundColor = $ForegroundColor
    }

    # Always write (if we want just a NewLine)
    if($Object -eq $null)
    {
        $Object = ""
    }

    if($NoNewline)
    {
        [Console]::Write($Object)
    }
    else
    {
        Write-Output $Object
    }

    # Restore previous colors
    $host.UI.RawUI.ForegroundColor = $previousForegroundColor
    $host.UI.RawUI.BackgroundColor = $previousBackgroundColor
}

Function Invoke-Logging{
    Param(
        [Parameter(Mandatory=$True)]
        [STRING]$Message,
		[Parameter(Mandatory=$True)]
		[Validateset('TEXT','TITLE','STATUS','AUDIT','INFO','SUCCESS','ALERT','WARNING','ERROR','CRITICAL','VERBOSE','DEBUG')]
		[STRING]$LogLevel,
        [Parameter(Mandatory=$False)]
        [System.IO.FileInfo]$RunLog
	)

		Switch($LogLevel){
			"TEXT"     {
				Invoke-ColorOutput $Message -ForegroundColor White
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"TITLE"    {
				Invoke-ColorOutput $Message -ForegroundColor Green
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"STATUS"   {
				$Message = (get-date -Format HH:mm:ss) + " - [STATUS]: " + $Message
				Invoke-ColorOutput $Message -ForegroundColor Magenta
				Update-Control -syncHash $synchash -Control TXT_output -Property TEXT -Value $Message -AppendContent
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"AUDIT"    {
				$Message = (get-date -Format HH:mm:ss) + " - [AUDIT]: " + $Message 
				Update-Control -syncHash $synchash -Control TXT_output -Property TEXT -Value $Message -AppendContent
				Invoke-ColorOutput $Message -ForegroundColor DarkGrey
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch
				{<#No Error Handling... This is horrible!#>
				}
			}
			"INFO"     {
				$Message = (get-date -Format HH:mm:ss) + " - [INFO]: " + $Message 
				Invoke-ColorOutput $Message -ForegroundColor White
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"SUCCESS"     {
				$Message = (get-date -Format HH:mm:ss) + " - [SUCCESS]: " + $Message 
				Invoke-ColorOutput $Message -ForegroundColor Green
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"ALERT"    {
				$Message = (get-date -Format HH:mm:ss) + " - [ALERT]: " + $Message
				Write-Host $Message -ForegroundColor Yellow
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"NEWLINE"  {Write-Host "";try{Write-Output "`n" | Out-File $RunLog -Append -ErrorAction SilentlyContinue}Catch{}}
			"WARNING"  {
				Write-Warning -Message $Message
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"ERROR"    {
				Write-Error -Message $Message
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"CRITICAL" {
				$Message = (get-date -Format HH:mm:ss) + " - [CRITICAL]: " + $Message
				Invoke-colorOutput $Message -ForegroundColor White -BGR Red
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"VERBOSE" {
				$Message = (get-date -Format HH:mm:ss) + " - [VERBOSE]: " + $Message
				Write-Verbose -Message $Message
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
			"DEBUG" {
				$Message = (get-date -Format HH:mm:ss) + " - [DEBUG]: " + $Message
				Invoke-colorOutput $Message -ForegroundColor White -BGR Cyan
				Try{
					$Message | Out-File $RunLog -Append -ErrorAction SilentlyContinue
				}Catch{
					<#No Error Handling... This is horrible!#>
				}
			}
		}
}