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
		$logFile,
        [switch]$AppendContent,
        [Switch]$Add
	)

    # This is kind of a hack, there may be a better way to do this
    If ($Property -eq "Close") {
		$syncHash.Window.Dispatcher.invoke([action]{$Hash.Window.Close()},"Normal")
        Return
	}

    # This updates the control based on the parameters passed to the function
    $syncHash.$Control.Dispatcher.Invoke([action]{
		# This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
		If ($PSBoundParameters['AppendContent']) {

			# Updating Control
			$current = $syncHash.$Control.$Property
			$Value = "$current`n$value"
			$syncHash.$Control.$Property = $Value
			$syncHash.$Control.ScrollToEnd()
				
			# Logging to file
			$Value | Out-File $LogFile -Append -ErrorAction SilentlyContinue
        }
            
		If ($PSBoundParameters['Add']) {
			$current = $syncHash.$Control.$Property
            $Value = $current + $Value
            $syncHash.$Control.$Property = $Value
		} Else {
			$syncHash.$Control.$Property = $Value
		}
	}, "Normal")
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
        $Global:ErrorActionPreference = 'stop'
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365crds -Authentication Basic -AllowRedirection
		$Global:ValidCreds = $true
    }
    Catch [System.Net.WebException],[System.Exception]
    {
        $Global:ValidCreds = $false
    }
    Finally
    {
    $Global:ErrorActionPreference = 'continue'
    Remove-PSSession $Session -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
}

Function Close-Window{
	$Source = -join($VariableHash.IconsDir,"\appbar.warning.png")
	$Image = New-Object System.Windows.Controls.Image
	$Image.Source = $Source
	$Image.Height = [System.Drawing.Image]::FromFile($Source).Height
	$Image.Width = [System.Drawing.Image]::FromFile($Source).Width
	$Image.Margin = 5
					 
	$TextBlock = New-Object System.Windows.Controls.TextBlock
	$TextBlock.Text = "Are you certain you wish to close the window?"
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
					
	Invoke-WPFMessageBox -Content $StackPanel -Title "WARNING!" -TitleBackground "Orange" -TitleTextForeground "Black" -TitleFontSize "20" -ButtonType OK-Cancel -WindowHost $SyncHash.Window

	If($WPFMessageBoxOutput -eq "OK"){
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

	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Starting connection process" -AppendContent

	#region Variables
	## Emtpy Hashtable for region
	$global:M365Services = @{}

	## Defining URI based on region
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Defining connection URI" -AppendContent
	
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
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Setting remote signed execution policy" -AppendContent
	Set-ExecutionPolicy RemoteSigned -Force

	#region Connections
		# Connecting to Exchange Online
		If($ExchangeOnline){
			If($MFA){
				Write-Verbose "Connecting to Exchange Online using MFA"
				Connect-EXOPSSession -UserPrincipalName $Credentials.UserName -ConnectionUri $global:M365Services['ConnectionEndpointUri'] -AzureADAuthorizationEndPointUri $global:M365Services['AzureADAuthorizationEndpointUri']
			}Else{
				# Reporting
				$Timestamp = (get-date -Format HH:mm:ss)
				Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Connecting to Exchange Online" -AppendContent

				# Creating the session
				$Session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $global:M365Services['ConnectionEndpointUri'] -Credential $Credentials -Authentication Basic -AllowRedirection
				# If the session was created, import it
				If($Session365){
					Try{
						Import-PSSession -Session $Session365 -AllowClobber -DisableNameChecking
						Update-control -Synchash $synchash -control IMG_Conn_EXO -property Source -value "$($VariableHash.IconsDir)\Check_Green.ico"
					}Catch{
						Update-control -Synchash $synchash -control IMG_Conn_EXO -property Source -value "$($VariableHash.IconsDir)\Check_Red.ico"
				}Else{
					Update-control -Synchash $synchash -control IMG_Conn_EXO -property Source -value "$($VariableHash.IconsDir)\Check_Red.ico"
				}
			}
		}
		}
		# Connect to Azure AD
		If($AzureAD){
			# Reporting
			$Timestamp = (get-date -Format HH:mm:ss)
			Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Connecting to Azure Active Directory" -AppendContent

			Try{
				Connect-AzureAD -Credential $credentials -AzureEnvironmentName $global:M365Services['AzureEnvironment']
				Update-control -Synchash $synchash -control IMG_Conn_AAD -property Source -value "$($VariableHash.IconsDir)\Check_Green.ico"
			}Catch{
				Update-control -Synchash $synchash -control IMG_Conn_AAD -property Source -value "$($VariableHash.IconsDir)\Check_Red.ico"
			}
		}

		# Connect to MSOL
		If($MSOL){
			# Reporting
			$Timestamp = (get-date -Format HH:mm:ss)
			Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Connecting to M365" -AppendContent

			Try{
				Connect-MsolService -Credential $credentials -AzureEnvironment $global:M365Services['AzureEnvironment']
				Update-control -Synchash $synchash -control IMG_Conn_M365 -property Source -value "$($VariableHash.IconsDir)\Check_Green.ico"
			}Catch{
				Update-control -Synchash $synchash -control IMG_Conn_M365 -property Source -value "$($VariableHash.IconsDir)\Check_Red.ico"
			}
			
		}

		# Connect to SharePoint Online
		If($SharePointOnline){
			# Reporting
			$Timestamp = (get-date -Format HH:mm:ss)
			Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Connecting to SharePoint Online" -AppendContent

			Try{
				Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
				$tenantName = ($TenantName).split(".")[0]
				Connect-SPOService -Url https://$TenantName-admin.sharepoint.com -credential $credentials -Region $global:M365Services['SharePointRegion']
				Update-control -Synchash $synchash -control IMG_Conn_SPO -property Source -value "$($VariableHash.IconsDir)\Check_Green.ico"
			}Catch{
				Update-control -Synchash $synchash -control IMG_Conn_M365 -property Source -value "$($VariableHash.IconsDir)\Check_Red.ico"
			}

		}

		# Connect to Skype for Business Online
		If($SkypeForBusinessOnline){
			Write-Verbose "Connecting to Skype Online"
			Import-Module SkypeOnlineConnector
			$sfboSession = New-CsOnlineSession -Credential $credentials
			If($sfboSession){Import-PSSession $sfboSession}
		}

		# Connect to security and compliance center
		If($SCC){
			If($MFA){
				Connect-IPPSSession -UserPrincipalName $Credentials.UserName -ConnectionUri $global:M365Services['SCCConnectionEndpointUri'] -AzureADAuthorizationEndPointUri $global:M365Services['AzureADAuthorizationEndpointUri']
			}Else{
				Write-Verbose "Connecting to Security and Compliance Center"
				$SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $global:M365Services['SCCConnectionEndpointUri'] -Credential $credentials -Authentication "Basic" -AllowRedirection
				If($SccSession){Import-PSSession $SccSession -Prefix cc}
			}
		}
	#endregion
	Update-control -Synchash $synchash -control Prereq_IMG_Connected -property Source -value "$($VariableHash.IconsDir)\Check_Green.ico"
}

Function Get-SPOInventory{


	# Loading Libraries
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Loading Sharepoint SDK Assemblies" -AppendContent
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
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Retrieving licensed and unlicensed users" -AppendContent
	$Users = Get-MsolUser -All 
	$UnLicensedUsers = Get-MsolUser -UnlicensedUsersOnly 

	# Retrieving all SPO sites, including Personal sites
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Retrieving all SPO Sites (including personal sites)" -AppendContent
	$SitesIncludingPersonal = Get-SPOSite -IncludePersonalSite $true -Limit All -Detailed 

	# Adding User to the site collection admins (Fixing an rights error encountered)
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Adding Inventory requests to the site collection admins" -AppendContent
	
	foreach($site in $Sites){ 
	  Set-SPOUser -Site $site.Url -LoginName $Credentials.UserName -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue 
	} 

	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Retrieving all site users" -AppendContent

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
	Update-control -Synchash $synchash -control IMG_Report_AllSites -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"

	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Retrieving all site groups" -AppendContent

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
	Update-control -Synchash $synchash -control IMG_Report_AllSiteGroups -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"

	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Retrieving Libraries gt 100" -AppendContent

	$Databases = @()
	foreach($site in $site){ 
	  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($site.Url) 
	  #Authenticate 
	  $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName , $Creds.Password) 
	  $ctx.Credentials = $credentials 
 
	  #Fetch the users in Site Collection 
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

	Update-control -Synchash $synchash -control IMG_Report_SPOLibrariesgt100 -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Retrieving libraries gt 0" -AppendContent

	$Databases = @(); 
	foreach($site in $Sites){ 
	  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($site.url) 
	  
	  #Authenticate 
	  $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName , $Creds.Password) 
	  $ctx.Credentials = $credentials 
 
	  #Fetch the users in Site Collection 
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
	Update-control -Synchash $synchash -control IMG_Report_SPOLibrariesgt0 -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"

	# Output List of sites
	$SitesIncludingPersonal | Select * | Export-Csv -Path "$($VariableHash.OutputPath)\SPOSitesIncludingPersonal.csv"

	# Output list of site users
	$Databases | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - AllSiteUsers.csv" -NoTypeInformation -Force 

	# Output list of site groups
	$Databases | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - AllSiteGroups.csv" -NoTypeInformation -Force 

	# Output libraries
	$Databases | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - Librariesgt100.csv" -NoTypeInformation -Force
	$Databases | Export-Csv -Path "$($VariableHash.OutputPath)\SPO - Librariesgt0.csv" -NoTypeInformation -Force 
}

Function Get-AADInventory{

	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Retrieving libraries gt 0" -AppendContent

	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Azure Active Directory user report" -AppendContent

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
	Update-control -Synchash $synchash -control IMG_Report_AADUsers -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"

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

if(!$NoExcel){
    #Create File path if doesn't exist
    if(!(Test-Path (Split-Path -Path $File))){
	    New-Item -ItemType directory -Path (Split-Path -Path $File) | out-null
    }
    #Excel file check, will DELETE if exists!!!
    if($file.split('.').count -gt 1){
        if (test-path $file) { rm $file }   #delete existing file
    }
}

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
    
        if($Results){
            $Results = $Results.replace("ï»¿","") | ConvertFrom-Csv
            if($Results){
                $Results | add-member –membertype NoteProperty –name 'Office365Report' –Value $ReportName 
                $objectCollection.Add($($ReportName),$Results)    
            }
            else{
               $objResults = New-Object –Type PSObject 
               $objResults | add-member –membertype NoteProperty –name 'Office365Report' –Value $ReportName  
               $objectCollection.Add($($ReportName),$objResults)  
            }
        }    
    }
}
if(!$NoExcel){
    #Excel object
    $Excel = New-Object -Com Excel.Application
    $Excel.Visible = $true
    $Excel.Workbooks.Add(1) | out-null
    $Workbook = $Excel.Workbooks.Item(1)

    #Modified Export-Excel function from: https://social.technet.microsoft.com/Forums/windows/en-US/abcf63ba-ce01-4e91-8bad-d5c42d2234e9/how-to-write-in-excel-via-powershell?forum=winserverpowershell
    function Export-Excel {
	    [cmdletBinding()]
	    Param(
		    [Parameter(ValueFromPipeline=$true)]
		    [string]$junk        )
	    begin{
		    $header = $null
		    $row = 1
            if($Workbook.WorkSheets.item(1).name -eq "sheet1"){
                $Worksheet = $Workbook.WorkSheets.item(1)
            }
            else{
                $Worksheet = $Workbook.Worksheets.Add()
            }
	    }
	    process{
		    if(!$header){
			    $i = 0
			    $header = $_ | Get-Member -MemberType NoteProperty | select name
			    $header | %{$Worksheet.cells.item(1,++$i)=$_.Name}
		    }
		    $i = 0
		    ++$row
		    foreach($field in $header){
			    $Worksheet.cells.item($row,++$i)=$($_."$($field.Name)")
		    }
	    }
	    end{
            $Worksheet.Name = $($_."Office365Report")
            $Worksheet.Columns.AutoFit() | out-null
            $Worksheet = $null
            $header = $null
	    }
    }

    #Loop for creating a worksheet for each report
    Foreach ($item in $objectCollection.Keys | sort -Descending){
       $objectCollection[$item] | Export-Excel
    }

    #Save Excel file
    $Workbook.SaveAs($file)
}

return $objectCollection
}

Function get-AzureInventory{

	# Creating credentials object
	$secpswd = ConvertTo-SecureString $VariableHash.M365password -AsPlainText -Force
	$Creds = new-object -typename System.Management.Automation.PSCredential -argumentlist $VariableHash.M365Username,$secpswd

	# Login to Azure RM
	Login-AzureRmAccount -Credential $Creds

	# Retrieving subscription list
	$Subscriptions = Get-AzureRmSubscription

	Foreach($subscription in $Subscriptions){
		# Setting subscription
		Select-AzureRmSubscription -Subscription $subscription

		# Getting items
		$AzureResources = Get-AzureRmResource
		

	}

}

Function get-AADDeletedUsers{
	$AADDeletedUsers_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$AADDeletedUsers_Observable.Clear()
			
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Azure Active Directory deleted user report" -AppendContent
	
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
				
		Update-control -Synchash $synchash -control IMG_Report_AADDelUser -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"
		
		# export
		$AADDeletedUsers_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - Deleted User Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_AADDelUser -property Source -Value "$($VariableHash.IconsDir)\Check_Red.ico"
	}		
}

Function get-AADContacts{
	$AADContacts_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$AADContacts_Observable.Clear()
			
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Azure Active Directory contacts report" -AppendContent
	
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
				
		Update-control -Synchash $synchash -control IMG_Report_AADContacts -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"
		
		# Export
		$AADContacts_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - Contacts Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_AADContacts -property Source -Value "$($VariableHash.IconsDir)\Check_Red.ico"
	}
}

Function get-AADGroups{
	$AADGroups_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$AADGroups_Observable.Clear()
			
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Azure Active Directory groups report" -AppendContent
	
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
		
		Update-control -Synchash $synchash -control IMG_Report_AADGroups -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"

		# Export
		$AADGroups_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - Groups Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_AADGroups -property Source -Value "$($VariableHash.IconsDir)\Check_Red.ico"
	}
}

Function get-AADDomains{
	$AADDomains_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$AADDomains_Observable.Clear()
			
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Azure Active Directory Domains report" -AppendContent
	
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
				
		Update-control -Synchash $synchash -control IMG_Report_AADDomains -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"
		
		# Export
		$AADDomains_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\AAD - Domains Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_AADDomains -property Source -Value "$($VariableHash.IconsDir)\Check_Red.ico"
	}		
}

Function get-ExchangeMailboxes{
	$ExchangeMailboxes_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$ExchangeMailboxes_Observable.Clear()
			
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Exchange Online mailboxes report" -AppendContent
	
	try{
		$ExchangeMailboxes = Get-Mailbox | sort DisplayName | select DisplayName, Alias, PrimarySMTPAddress, ArchiveStatus, UsageLocation, WhenMailboxCreated
		
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
						LastLogonTime = $statistics.LastLogonTime
					}
				))   					
			}
		}
				
		Update-control -Synchash $synchash -control IMG_Report_EXOMailboxes -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"
		
		# Export
		$ExchangeMailboxes_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\EXO - Mailboxes Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_EXOMailboxes -property Source -Value "$($VariableHash.IconsDir)\Check_Red.ico"
	}			
}

Function get-ExchangeArchives{
	$ExchangeArchives_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$ExchangeArchives_Observable.Clear()
			
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Exchange Online Archives report" -AppendContent
	
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

		Update-control -Synchash $synchash -control IMG_Report_EXOArchives -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"
		
		# Export
		$ExchangeArchives_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\EXO - Archives Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_EXOArchives -property Source -Value "$($VariableHash.IconsDir)\Check_Red.ico"
	}		
}

Function get-ExchangeGroups{
	$ExchangeGroups_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
	$ExchangeGroups_Observable.Clear()
			
	# Reporting
	$Timestamp = (get-date -Format HH:mm:ss)
	Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Exchange Online Groups report" -AppendContent
	
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
		Update-control -Synchash $synchash -control IMG_Report_EXOGroups -property Source -Value "$($VariableHash.IconsDir)\Check_Green.ico"

		# Export
		$ExchangeGroups_Observable | Export-Csv -Path "$($VariableHash.OutputPath)\EXO - Groups Report.csv" -NoTypeInformation -Force
	}catch{
		Update-control -Synchash $synchash -control IMG_Report_EXOGroups -property Source -Value "$($VariableHash.IconsDir)\Check_Red.ico"	
	}		
}

Function Run-Reports{
	# Variables
					$VariableHash.tenantname = $SyncHash.txt_Settings_tenant.text
					$VariableHash.GraphAppID = $syncHash.txt_Settings_GraphAppID.Text
					$VariableHash.GraphRedirectUri = $syncHash.txt_Settings_GraphRedirectUri.Text
					$VariableHash.TenantRegion = $SyncHash.CB_Settings_TenantRegion.CurrentItem
					If($VariableHash.TenantRegion -like ""){
							$VariableHash.TenantRegion = "Default"
						}

					$Report_Runspace =[runspacefactory]::CreateRunspace()
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
							Update-control -synchash $synchash -control IMG_Wizard -property Source -Value "$($VariableHash.ImageDir)\wizard_redeyes.png"

							# Creating credentials object
							$secpswd = ConvertTo-SecureString $VariableHash.M365password -AsPlainText -Force
							$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $VariableHash.M365Username,$secpswd

							# Calling connection function
							Connect-M365 -Credentials $credentials -Region $VariableHash.TenantRegion -TenantName $VariableHash.tenantname -ExchangeOnline -AzureAD -MSOL -SharePointOnline

							# Azure Active Directory User Report
							get-AADInventory

							# SPO site inventories
							get-SPOInventory

							#region Graph reports
								$Timestamp = (get-date -Format HH:mm:ss)
								Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Running Microsoft Graph reports" -AppendContent
								$Graphreports = get-O365UsageReports -AzureAppClientId $variableHash.GraphAppID -Credential $credentials -Period D180 -NoExcel

								Foreach($item in $Graphreports.keys){
									$Graphreports[$item] | Export-Csv -Path "$($VariableHash.OutputPath)\Graph - $item.csv" -NoTypeInformation
								}

								Update-control -Synchash $synchash -control IMG_Report_Graph -property Source -value "$($VariableHash.IconsDir)\Check_Green.ico"
							#region Graph reports

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

							# Disconnecting exchange (Pesky little session limit...)
							$Timestamp = (get-date -Format HH:mm:ss)
							Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Disconnecting Exchange Online" -AppendContent
							Remove-PSSession $Session
							Update-control -Synchash $synchash -control IMG_Conn_EXO -property Source -value "$($VariableHash.IconsDir)\Check_Waiting.ico"
						
							# Reporting finished
							$Timestamp = (get-date -Format HH:mm:ss)
							Update-control -Synchash $synchash -control txt_output -property Text -value "[$timestamp] - Report run completed" -AppendContent
							Update-control -synchash $synchash -control IMG_Wizard -property Source -Value "$($VariableHash.ImageDir)\wizard.png"
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