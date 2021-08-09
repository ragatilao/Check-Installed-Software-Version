clear

#**************************************************************************************************************
# add the required .NET assembly
#**************************************************************************************************************
Add-Type -AssemblyName System.Windows.Forms

#**************************************************************************************************************
#WPF XML GUI Presentation Layout
#**************************************************************************************************************

Add-Type -AssemblyName PresentationFramework

[xml]$Form  = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowershellWPFs"
        Title="Check Installed Software Version" Height="600" Width="1250" Background="White" Grid.IsSharedSizeScope="True" Topmost="False">
    <Grid Margin="0,0,0,0" >
        <GroupBox Name="groupBox" Header="Server List Option" HorizontalAlignment="Left" Height="85" VerticalAlignment="Top" Width="745" Margin="5,5,0,0">
            <Grid HorizontalAlignment="Left" Height="70" Margin="0,0,0,0" VerticalAlignment="Top" Width="740">
                <StackPanel Name="sp_radiobuttons" HorizontalAlignment="Left" Height="68" Width="125" Margin="5,5,0,0" VerticalAlignment="Top" >
                    <RadioButton Name="rb_servername" Content="Server Name/s" GroupName="Group1" Height="30" IsChecked="True" VerticalContentAlignment="Center" />
                    <RadioButton Name="rb_serverlist" Content="Server List File Path" GroupName="Group1" Height="30" VerticalContentAlignment="Center" />
                </StackPanel>
                <TextBox Name="tb_servername" HorizontalAlignment="Left" Height="23" Margin="140,9,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="500" ToolTip="Enter server name or servers in comma separated string."/>
                <TextBox Name="tb_serverlist" HorizontalAlignment="Left" Height="23" Margin="140,39,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="500" IsEnabled="False"/>
                <Button Name="btn_browsefile" Content="Browse File..." HorizontalAlignment="Left" Margin="645,40,0,0" VerticalAlignment="Top" Width="85" IsEnabled="False"/>
            </Grid>
        </GroupBox>
        <GroupBox Name="groupBox1" Header="Search Option" HorizontalAlignment="Left" Height="85" Margin="755,5,0,0" VerticalAlignment="Top" Width="475">
            <Grid HorizontalAlignment="Left" Height="75" Margin="0,0,-2,-12" VerticalAlignment="Top" Width="465">
                <TextBox Name="tb_searchstring" HorizontalAlignment="Left" Height="23" Margin="85,5,0,0" TextWrapping="Wrap" Text="Microsoft SQL Server Management Studio*" VerticalAlignment="Top" Width="270"/>
                <Label Name="label" Content="Search String" HorizontalAlignment="Left" Margin="5,3,0,0" VerticalAlignment="Top"/>
                <Label Name="label1" Content="NOTE: To filter result set, use unique strings of the Windows Program Name/Title." HorizontalAlignment="Left" Margin="5,33,0,0" VerticalAlignment="Top" Background="#FFFFFFDE" Width="455"/>
                <Button Name="btn_runvalidation" Content="Run Search" HorizontalAlignment="Left" Margin="360,6,0,0" VerticalAlignment="Top" Width="100"/>
            </Grid>
        </GroupBox>
        <GroupBox Name="groupBox2" Header="Validation Results" HorizontalAlignment="Left" Height="460" Margin="5,95,0,0" VerticalAlignment="Top" Width="1225">
            <Grid HorizontalAlignment="Left" Height="429" Margin="0,0,0,0" VerticalAlignment="Top" Width="1220">
                <DataGrid Name="dataGrid" HorizontalAlignment="Left" Height="420" Margin="5,5,0,0" VerticalAlignment="Top" Width="1205" Style="{DynamicResource DGHeaderStyle}" IsReadOnly="True" HeadersVisibility="Column" CanUserSortColumns="False">
                    <DataGrid.Resources>
                        <Style TargetType="DataGridColumnHeader">  
                            <Setter Property="Background" Value="CornFlowerBlue" />  
                            <Setter Property="Foreground" Value="White"/>  
                            <Setter Property="BorderBrush" Value="DarkGray" />
                            <Setter Property="BorderThickness" Value="1" />
                            <Setter Property="HorizontalContentAlignment" Value="Center" />
                            <Setter Property="FontWeight" Value="Bold" />
                        </Style>
                    </DataGrid.Resources>
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="SERVER NAME" Width="120" Binding="{Binding Computername}"/>
                        <DataGridTextColumn Header="CLUSTER STATUS" Width="250" Binding="{Binding ClusterStatus}"/>
                        <DataGridTextColumn Header="PROGRAM NAME" Width="540" Binding="{Binding AppName}"/>
                        <DataGridTextColumn Header="VERSION" Width="*" Binding="{Binding AppVersion}"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </GroupBox>
    </Grid>

</Window>
"@

$NR=(New-Object System.Xml.XmlNodeReader $Form)
$Win=[Windows.Markup.XamlReader]::Load( $NR )
#$Win.Add_Closing({$Win.Closing; $Win.Close();})

#**************************************************************************************************************
#Progress Bar Function
#**************************************************************************************************************

Function Update-Window { 
    Param ( 
        $Control, 
        $Property, 
        $Value, 
        [switch]$AppendContent 
    ) 
 
   # This is kind of a hack, there may be a better way to do this 
   If ($Property -eq "Close") { 
      $syncHash.Window.Dispatcher.invoke([action]{$syncHash.Window.Close()},"Normal") 
      Return 
   } 
   
   # This updates the control based on the parameters passed to the function 
   $syncHash.$Control.Dispatcher.Invoke([action]{ 
      # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps 
       If ($PSBoundParameters['AppendContent']) { 
           $syncHash.$Control.AppendText($Value) 
       } Else { 
           $syncHash.$Control.$Property = $Value 
       } 
   }, "Normal") 
} 

#**************************************************************************************************************
#Funtions
#**************************************************************************************************************

function Get-WindowsClusterStatus
{
    param([string]$servername)

    $OwnerNodeName = Get-WMIObject -Class MSCluster_ResourceGroup -ComputerName $serverName -Namespace root\mscluster -ErrorAction SilentlyContinue 

    if ($OwnerNodeName -eq $null)
    {
        Write-Output 'Standalone'
    }
    else
    {
        $OwnerNodeName = $OwnerNodeName | Where-Object Name -NotIn ('Available Storage') | Select-Object -ExpandProperty OwnerNode -Unique -ErrorAction SilentlyContinue 

        if ($OwnerNodeName.count -eq 1)
        {

            if ($OwnerNodeName.Trim() -eq $serverName.Trim())
            {
                Write-Output 'Active'
            }
            else 
            {
                Write-Output 'Passive'
            }
        }
        else
        {
            Write-Output 'Cluster resources owned by multiple nodes. (Unknown)'
        } 
    }
}

Function Get-Software  {
    [OutputType('System.Software.Inventory')]
    #Visit below link for usage instructions and other details.
    # http://techibee.com/powershell/powershell-script-to-query-softwares-installed-on-remote-computer/1389
    [cmdletbinding()]
    param(
     [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
     [string[]]$ComputerName = $env:computername
    )

    begin {
     $UninstallRegKeys=@("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
					    "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
    }

    process {
     foreach($Computer in $ComputerName) {
      Write-Verbose "Working on $Computer"
	    if(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {
		    foreach($UninstallRegKey in $UninstallRegKeys) {
			    try {
				    $HKLM   = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computer)
				    $UninstallRef  = $HKLM.OpenSubKey($UninstallRegKey)
				    $Applications = $UninstallRef.GetSubKeyNames()
			    } catch {
				    Write-Verbose "Failed to read $UninstallRegKey"
				    Continue
			    }

			    foreach ($App in $Applications) {
			    $AppRegistryKey  = $UninstallRegKey + "\\" + $App
			    $AppDetails   = $HKLM.OpenSubKey($AppRegistryKey)
			    $AppGUID   = $App
			    $AppDisplayName  = $($AppDetails.GetValue("DisplayName"))
			    $AppVersion   = $($AppDetails.GetValue("DisplayVersion"))
			    $AppPublisher  = $($AppDetails.GetValue("Publisher"))
			    $AppInstalledDate = $($AppDetails.GetValue("InstallDate"))
			    $AppUninstall  = $($AppDetails.GetValue("UninstallString"))
			    if($UninstallRegKey -match "Wow6432Node") {
				    $Softwarearchitecture = "x86"
			    } else {
				    $Softwarearchitecture = "x64"
			    }
			    if(!$AppDisplayName) { continue }
			    $OutputObj = New-Object -TypeName PSobject 
			    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()
			    $OutputObj | Add-Member -MemberType NoteProperty -Name AppName -Value $AppDisplayName
			    $OutputObj | Add-Member -MemberType NoteProperty -Name AppVersion -Value $AppVersion
			    $OutputObj | Add-Member -MemberType NoteProperty -Name AppVendor -Value $AppPublisher
			    $OutputObj | Add-Member -MemberType NoteProperty -Name InstalledDate -Value $AppInstalledDate
			    $OutputObj | Add-Member -MemberType NoteProperty -Name UninstallKey -Value $AppUninstall
			    $OutputObj | Add-Member -MemberType NoteProperty -Name AppGUID -Value $AppGUID
			    $OutputObj | Add-Member -MemberType NoteProperty -Name SoftwareArchitecture -Value $Softwarearchitecture
			    $OutputObj
			    }
		    }	
	    }
     }
    }
}

#**************************************************************************************************************
#Declare variables for the WPF in Powershell
#**************************************************************************************************************

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 
$syncHash = [hashtable]::Synchronized(@{}) 
$newRunspace =[runspacefactory]::CreateRunspace() 
$newRunspace.ApartmentState = "STA" 
$newRunspace.ThreadOptions = "ReuseThread"           
$newRunspace.Open() 
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)           
$psCmd = [PowerShell]::Create().AddScript({    
    [xml]$xaml = @" 
    <Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
        Name="Window" Title="Progress..." WindowStartupLocation = "CenterScreen" Width = "335" Height = "70" ShowInTaskbar = "True" Topmost="True" ResizeMode="NoResize" WindowStyle="None" BorderBrush="Black" BorderThickness="1" Background="LightYellow">
        <Grid> 
           <ProgressBar Name = "ProgressBar" Height = "20" Width = "300" HorizontalAlignment="Left" VerticalAlignment="Top" Margin = "10,10,0,0"/> 
           <Label Name = "Label1" Height = "30" Width = "300" HorizontalAlignment="Left" VerticalAlignment="Top" Margin = "10,35,0,0"/> 
        </Grid> 
    </Window> 
"@ 

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader ) 
    $syncHash.ProgressBar = $syncHash.Window.FindName("ProgressBar") 
    $syncHash.Label1 = $syncHash.Window.FindName("Label1")
    $syncHash.Window.ShowDialog() | Out-Null 
    $syncHash.Error = $Error
    
}) 

#Group Box
$sp_radiobuttons = $Win.FindName("sp_radiobuttons")

#Server List Option
$rb_servername = $Win.FindName("rb_servername")
$rb_serverlist = $Win.FindName("rb_serverlist")
$tb_servername = $Win.FindName("tb_servername")
$tb_serverlist = $Win.FindName("tb_serverlist")
$btn_browsefile = $Win.FindName("btn_browsefile")

#Search Option
$tb_searchstring = $Win.FindName("tb_searchstring")
$btn_runvalidation = $Win.FindName("btn_runvalidation")

#Validation Results
$dataGrid = $Win.FindName("dataGrid")
$progressBar = $Win.FindName("progressBar")


#**************************************************************************************************************
#Powershell Body
#**************************************************************************************************************

$rb_servername.Add_Checked({
    #$tb_servername.Text = $_.source.name

    if ($rb_servername.IsChecked -eq $true)
    {
        $tb_serverlist.IsEnabled = $false 
        $btn_browsefile.IsEnabled = $false
        $tb_servername.IsEnabled = $true
        $tb_serverlist.Text = $null
        $tb_servername.Text = $null
        $tb_searchstring.Text = "Microsoft SQL Server Management Studio*"
        $dataGrid.ItemsSource = $null
    }
})

$rb_serverlist.Add_Checked({
    #$tb_servername.Text = $_.source.name

    if ($rb_serverlist.IsChecked -eq $true)
    {
        $tb_serverlist.IsEnabled = $true 
        $btn_browsefile.IsEnabled = $true
        $tb_servername.IsEnabled = $false
        $tb_serverlist.Text = $null
        $tb_servername.Text = $null
        $tb_searchstring.Text = "Microsoft SQL Server Management Studio*"
        $dataGrid.ItemsSource = $null
    }
})

$btn_browsefile.Add_Click({
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Multiselect = $false # Multiple files can be chosen
	    Filter = 'Text(*.txt)|*.txt' # Specified file types
    }

    [void]$FileBrowser.ShowDialog()

    $file = $FileBrowser.FileName;
    $tb_serverlist.Text = $file
})

$btn_runvalidation.Add_Click({
    
    $KBs = @()
    $KBNumber_or_TitleString = $tb_searchstring.Text
    $progressbarcount = 0 
    
    if ($tb_searchstring.Text -eq "")
    {
        [System.Windows.Forms.Messagebox]::Show("Please provide a search string.","Error",0,16)
        $tb_searchstring.Focus();
    }
    elseif (($tb_servername.Text -eq "") -and ($rb_servername.IsChecked -eq $true))
    {
        [System.Windows.Forms.Messagebox]::Show("Please provide a server name/s.","Error",0,16)
        $tb_servername.Focus();
    }
    elseif (($tb_serverlist.Text -eq "") -and ($rb_serverlist.IsChecked -eq $true))
    {
        [System.Windows.Forms.Messagebox]::Show("Please provide the file path of the server list.","Error",0,16)
        $tb_serverlist.Focus();
    }
    else
    {
        #[System.Windows.Forms.Messagebox]::Show("Validate the Servers","Success",0,0)
        $dataGrid.Items.Clear()
                                
        $psCmd.Runspace = $newRunspace 
        $data = $psCmd.BeginInvoke()
        While (!($syncHash.Window.IsInitialized)) { 
           Start-Sleep -m 500 
        } 

        if ($rb_servername.IsChecked -eq $true)
        {
            $serverlist  = $tb_servername.Text.Split(",")
            #$serverlist = $serverlists.Split(",")
        }
        else
        {
            $serverlistpath = $tb_serverlist.Text.ToString()
            $serverlist = get-content $serverlistpath
        }

        foreach ($svr in $serverlist)
        {

            $progressbarcount ++ 
            Update-Window Label1 Content "Processing $svr ..."   
            Update-Window ProgressBar Value "$(($progressbarcount/$serverlist.Count)*100)"   

            $svr = $svr.Trim()

            $ssms_version = $null
            $ssms_version = Get-Software ($svr) | SELECT -Unique Computername, AppName, AppVersion | Where-Object {$_.AppName -like "*$KBNumber_or_TitleString*"}
            #$ssms_version = Get-Software ($server) | SELECT -Unique Computername, AppName, AppVersion, InstalledDate | Where-Object AppName -Like "Microsoft SQL Server Management Studio*"
            #$ssms_version 
            $WindowsClusterStatus = $null
            $WindowsClusterStatus = Get-WindowsClusterStatus -servername $svr

            if ($ssms_version -eq $null)
            {
                $KB = New-Object -TypeName PSObject 
                $KB | Add-Member -MemberType NoteProperty -Name Computername -Value $svr
                if ($WindowsClusterStatus -eq $null)
                {
                    $KB | Add-Member -MemberType NoteProperty -Name ClusterStatus -Value '--'
                } else {
                    $KB | Add-Member -MemberType NoteProperty -Name ClusterStatus -Value $WindowsClusterStatus
                }
                $KB | Add-Member -MemberType NoteProperty -Name AppName -Value '--'
                $KB | Add-Member -MemberType NoteProperty -Name AppVersion -Value '--'
                #$KB | Add-Member -MemberType NoteProperty -Name InstalledDate -Value '--'
                $KBs += $KB
            } else {
                $counter = 0
                foreach($appversion in $ssms_version)
                {
                    #$KBs += $appversion 
                    $KB = New-Object -TypeName PSObject 

                    if ($counter -eq 0)
                    {
                        $KB | Add-Member -MemberType NoteProperty -Name Computername -Value $appversion.Computername
                        if ($WindowsClusterStatus -eq $null)
                        {
                            $KB | Add-Member -MemberType NoteProperty -Name ClusterStatus -Value '--'
                        } else {
                            $KB | Add-Member -MemberType NoteProperty -Name ClusterStatus -Value $WindowsClusterStatus
                        }
                    } else {
                        $KB | Add-Member -MemberType NoteProperty -Name Computername -Value ''
                        $KB | Add-Member -MemberType NoteProperty -Name ClusterStatus -Value ''

                    }
                    $counter++

                    $KB | Add-Member -MemberType NoteProperty -Name AppName -Value $appversion.AppName
                    $KB | Add-Member -MemberType NoteProperty -Name AppVersion -Value $appversion.AppVersion
                    $KBs += $KB
                }
            }
        }

        foreach ($KBRow in $KBs)
        {
            $dataGrid.Items.Add($KBRow)
        }
        
        #$dataGrid.ItemsSource=@($KBs)
        #$tb_servername.Text = $KBs.Count
        #$KBs | ft -AutoSize -Wrap

        # This closes the progress bar 
        Update-Window Window Close 
    }
})

[void]$Win.ShowDialog()
