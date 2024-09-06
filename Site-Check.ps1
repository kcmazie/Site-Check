Param(
    [Switch]$Console = $false,
    [Switch]$Debug = $false   
)
<#------------------------------------------------------------------------------ 
         File Name : Site-Check.ps1 
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com) 
                   : 
       Description : This script is designed to be a quick-check of a list of locations (sites) and network IP 
                   : addresses (targets). It is completely generic and is configured via an XML config file which
                   : must be located in the same folder and named the same as as the script but with a ".xml" extension.
                   : If the config file is NOT detected a sample XML file will be created that must be edited before use.
                   : The script dynamically adjusts it's size depending on the number of sites and/or targets in the config
                   : file up to a limit after which it will exceed the screen size.  The script uses a simple ICMP echo 
                   : (ping) to identify if a target is alive. Select one of the listed sites and the associated info from 
                   : the config file will populate the GUI. Each time you select a different site the info will dynamically
                   : update. Once the site is loaded simply click the Execute button to run the ping check. Results are noted
                   : along the right side of the screen.  After each run you can select a different site and execute again 
                   : continuously until you click the cancel button. 
                   : 
             Notes : Normal operation is with no command line options.   
                   : 
         Arguments : Command line options for testing: 
                   : - "-console $true" will enable local console echo for troubleshooting
                   : - "-debug $true" will only email to the first recipient on the list
                   : 
          Warnings : None 
                   : 
             Legal : Public Domain. Modify and redistribute freely. No rights reserved. 
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED. 
                   : That being said, please let me know if you find bugs or improve the script. 
                   : 
           Credits : Code snippets and/or ideas came from many sources including but 
                   : not limited to the following: n/a 
                   : 
    Last Update by : Kenneth C. Mazie 
   Version History : v1.00 - 09-06-24 - Original Release
    Change History : v2.00 - 00-00-00 - 
                   : #>
                   $ScriptVer = "1.00"    <#--[ Current version # used in script ]--
                   : 
------------------------------------------------------------------------------#>
#Requires -Version 5.1
Clear-Host 

#--[ Suppress Console ]-------------------------------------------------------
Add-Type -Name Window -Namespace Console -MemberDefinition ' 
[DllImport("Kernel32.dll")] 
public static extern IntPtr GetConsoleWindow(); 
 
[DllImport("user32.dll")] 
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow); 
' 
$ConsolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($ConsolePtr, 0) | Out-Null
#------------------------------------------------------------------------------#>
 
#--[ Runtime Variables For Testing ]-------------
#$Script:Console = $true
#$Script:Debug = $true
#------------------------------------------------

$ErrorActionPreference = "silentlycontinue"

#==[ Functions ]================================================================
Function GetConsoleHost ($ExtOption){  #--[ Detect if we are using a script editor or the console ]--
    Switch ($Host.Name){
        'consolehost'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $False -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell Console detected." -Force
        }
        'Windows PowerShell ISE Host'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell ISE editor detected.  Console mode enabled." -Force
        }
        'PrimalScriptHostImplementation'{
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "COnsoleMessage" -Value "PrimalScript or PowerShell Studio editor detected.  Console mode enabled." -Force
        }
        "Visual Studio Code Host" {
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ExtOption | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "Visual Studio Code editor detected.  Console mode enabled. " -Force
        }
    }
    If ($ExtOption.ConsoleState){
        StatusMsg "Script Status: Initializing and beginning script execution" "Yellow " $ExtOption
        StatusMsg "Detected session running from an editor..." "Cyan" $ExtOption
    }
    Return $ExtOption
}

Function ReloadForm ($Form){
    $Form.Close()
    $Form.Dispose()
    ActivateForm
}

Function KillForm ($Form) {
    $Form.Close()
    $Form.Dispose()
}
Function UpdateOutput ($Form){  #--[ Refreshes the infobox contents ]--
    $InfoBox.update()
    $InfoBox.Select($InfoBox.Text.Length, 0)
    $InfoBox.ScrollToCaret()
}

Function IsThereText ($TargetBox){  #--[ Checks for text in the text entry box(es) ]--
  if (($TargetBox.Text.Length -ge 8)){ 
    Return $true
  }else{
    Return $false
  }
}

Function LoadConfig ($ConfigFile,$SiteCode){  #--[ Read and load configuration file ]-------------------------------------
    [xml]$Config = Get-Content $ConfigFile  #--[ Read & Load XML ]--  
    $XmlData = New-Object -TypeName psobject 
    $Site = "Site"+($SiteCode.Split(",")[1])
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "TargetTemplate" -Value $Config.Settings.TargetTemplate  
    If ($SiteCode -eq ""){
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "TitleText" -Value $Config.Settings.General.TitleText
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "EmailRecipient" -Value $Config.Settings.General.EmailRecipient
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "EmailSender" -Value $Config.Settings.General.EmailSender
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "SmtpServer" -Value $Config.Settings.General.SmtpServer
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "SmtpPort" -Value $Config.Settings.General.SmtpPort
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "SiteList" -Value  $Config.Settings.Sites.Site
    }Else{ 
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Location" -Value $Config.Settings.$Site.Location
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Name" -Value $Config.Settings.$Site.Name
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Address" -Value $Config.Settings.$Site.Address
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Contact" -Value $Config.Settings.$Site.Contact
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Email" -Value $Config.Settings.$Site.Email
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "DeskPhone" -Value $Config.Settings.$Site.DeskPhone
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "CellPhone" -Value $Config.Settings.$Site.CellPhone
        $Counter = 1
        $Template = $Config.Settings.TargetTemplate.Split(';')
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "TargetCount" -Value $Template.Count
        While ($Counter -le $Template.count){
            $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Target$Counter" -Value $Config.Settings.$Site.("Target$Counter")
            $Counter++
        }
    }
    Return $XmlData
}

Function StatusMsg ($Msg, $Color, $ExtOption){
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    If ($ExtOption.ConsoleState){
        Write-Host "-- Script Status: $Msg" -ForegroundColor $Color
        $Msg = ""
    }
}

#==[ End of Functions ]=======================================================

#==[ Begin Processing ]=======================================================
#--[ Load external XML options file ]-----------------------------------------
$ConfigFile = $PSScriptRoot+"\"+($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"
If (Test-Path -Path $ConfigFile -PathType Leaf){
    $ExtOption = LoadConfig $ConfigFile ""  #--[ This creates the initial PSObject ]--
    StatusMsg "Loading external config file..." "Magenta" $ExtOption
    $SiteList = [Ordered]@{}  
    $Index = 1
    ForEach($Site in $ExtOption.Sitelist){
        $SiteList.Add($Index, @($Site))
        $Index++
    }
    $ExtOption = GetConsoleHost $ExtOption
    If ($Console){  
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ConsoleState" -Value $true   
    }
    If ($Debug){  
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $true   
    }
}Else{
    $Msg = "The XML config file is required by this script!`nA basic one has been created for you.`nScript operation has been aborted..."
    $Button = [System.Windows.Forms.MessageBoxButtons]::OK
    $Icon = [System.Windows.Forms.MessageBoxIcon]::Warning
    [System.Windows.Forms.MessageBox]::Show($Msg,"Attention!  XML CONFIG FILE IS MISSING",$Button,$Icon)
  
$XMLSample = '<?xml version="1.0" encoding="utf-8"?>
    <Settings>
        <General>
            <TitleText>Enterprise Network Engineering "Quick" Site Check</TitleText>            
            <SmtpServer>Not Used</SmtpServer>
            <SmtpPort>25</SmtpPort>
            <EmailRecipient>Not Used</EmailRecipient>
            <EmailSender>Not Used</EmailSender>
        </General>
        <TargetTemplate>LoopbackIP;GatewayIP;Circuit1;PrivateIP1;PublicIP1;Circuit2;PrivateIP2;PublicIP2</TargetTemplate> 
            <!--  NOTE: Target Template MUST match the order and number of targets in each site section. Below is an EXAMPLE. 
            <Target1>Loopback IP</Target1>
            <Target2>Gateway IP</Target2>
            <Target3>Circuit 1</Target3> 
            <Target4>Private IP 1</Target4>
            <Target5>Public IP 1</Target5>
            <Target6>Circuit 2</Target6>
            <Target7>Private IP 2</Target7>
            <Target8>Public IP 2</Target8>
            -->
        <Sites>
        	<!-- Site ID number must match site tag below, i.e "sitename,1" and "<Site1>" -->
            <site>New York,1</site>
            <site>San Francisco,2</site>
            <site>Los Angeles,3</site> 
            <site>Denver,4</site>
            <site>San Jose,5</site>
            <site>Seattle,6</site>
            <site>Atlanta,7</site>
            <site>Chicago,8</site>
            <site>London,9</site>
            <site>Paris,10</site> 
        </Sites>
        <Site1>
            <Designation>Site-01</Designation>
            <Name>New York Office</Name>
            <Address>123 Sesame Street, New York NY 12345</Address>
            <Contact>Joe Tech</Contact>
            <Email>techdude@123.com</Email>
            <CellPhone>123-456-7890</CellPhone>
            <DeskPhone>234-567-8901</DeskPhone>
            <Target1>10.0.1.1</Target1>
            <Target2>10.1.1.1</Target2>
            <Target3>Comcast: 45.YADA.12345..XYZL</Target3> 
            <Target4>10.2.2.7</Target4>
            <Target5>10.2.2.6</Target5>
            <Target6>ATT: 50ASDF666333PT</Target6>
            <Target7>10.3.3.7</Target7>
            <Target8>10.3.3.6</Target8>
        </Site1>
        <Site2>
            <Designation>Site-02</Designation>
            <Name>San Francisco Branch</Name>
            <Address>123 Sesame Street, San Francisco CA 92345</Address>
            <Contact>Joe Tech</Contact>
            <Email>techdude@123.com</Email>
            <CellPhone>123-456-7890</CellPhone>
            <DeskPhone>234-567-8901</DeskPhone>
            <Target1>20.0.1.1</Target1>
            <Target2>20.1.1.1</Target2>
            <Target3>Comcast: 23.YADA.12345..XYZL</Target3> 
            <Target4>30.2.2.7</Target4>
            <Target5>30.2.2.6</Target5>
            <Target6>ATT: 50ASDF444333PT</Target6>
            <Target7>10.3.3.7</Target7>
            <Target8>10.3.3.6</Target8>
        </Site2>
    </Settings>'
    Add-Content -Path $ConfigFile -Value $XMLSample
Break;Break;Break
}

#--[ Prep GUI ]------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$Icon = [System.Drawing.SystemIcons]::Information

#$ScreenSize = (Get-WmiObject -Class Win32_DesktopMonitor | Select-Object ScreenWidth,ScreenHeight)  --[ Depricated ]--
$ScreenSize = (Get-CimInstance -Class Win32_DesktopMonitor | Select-Object ScreenWidth,ScreenHeight)

If ($ScreenSize.Count -gt 1){  #--[ Detect multiple monitors ]--
    StatusMsg "More than 1 monitor detected..." "Magenta"
    ForEach ($Resolution in $ScreenSize){
        If ($Null -ne $Resolution.ScreenWidth){
            $ScreenWidth = $Resolution.ScreenWidth
            $ScreenHeight = $Resolution.ScreenHeight
            Break
        }
    }
}Else{
    $ScreenWidth = $Resolution.ScreenWidth
    $ScreenHeight = $Resolution.ScreenHeight
}

#--[ Define Form ]--------------------------------------------------------------
[int]$FormWidth = 727
If ($SiteList.Count/2 -is [int]){
    [int]$FormHeight = ((($ExtOption.SiteList.Count/2)*20)+255)   #--[ Dynamically Created Variable for Box Size (Even count) ]--
}Else{
    [int]$FormHeight = ((($ExtOption.SiteList.Count/2)*23)+255)   #--[ Dynamically Created Variable for Box Size (Odd count) ]--
}

[int]$FormHCenter = ($FormWidth / 2)   # 170 Horizontal center point
[int]$FormVCenter = ($FormHeight / 2)  # 209 Vertical center point
[int]$ButtonHeight = 25
[int]$TextHeight = 20

#--[ Create Form ]---------------------------------------------------------------------
$LineAdd = 24
$Form = New-Object System.Windows.Forms.Form    
$Form.AutoSize = $False
$Notify = New-Object system.windows.forms.notifyicon
$Notify.icon = $Icon              #--[ NOTE: Available tooltip icons are = warning, info, error, and none
$Notify.visible = $true
$Form.Text = "Script Version: $ScriptName v$ScriptVer"
$Form.StartPosition = "CenterScreen"
$Form.KeyPreview = $true
$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$Form.Close();$Stop = $true}})
$ButtonFont = new-object System.Drawing.Font("Microsoft Sans Serif Regular",9,[System.Drawing.FontStyle]::Bold)

#--[ Form Title Label ]-----------------------------------------------------------------
$BoxLength = 350
$LineLoc = 5
$FormLabelBox = new-object System.Windows.Forms.Label
$FormLabelBox.Font = $ButtonFont
$FormLabelBox.Location = new-object System.Drawing.Size(($FormHCenter-($BoxLength/2)-15),$LineLoc)
$FormLabelBox.size = new-object System.Drawing.Size($BoxLength,$ButtonHeight)
$FormLabelBox.TextAlign = 2 
$FormLabelBox.Text = $ExtOption.TitleText
$Form.Controls.Add($FormLabelBox)

#--[ CLOSE Button ]------------------------------------------------------------------------
$BoxLength = 100
$ButtonLineLoc = 12
$CloseButton = new-object System.Windows.Forms.Button
$CloseButton.Location = New-Object System.Drawing.Size(($FormHCenter-($BoxLength/2)-265),$ButtonLineLoc)
$CloseButton.Size = new-object System.Drawing.Size($BoxLength,$ButtonHeight)
$CloseButton.TabIndex = 1
$CloseButton.Text = "Cancel/Close"
$CloseButton.Add_Click({
    KillForm $Form
})
$Form.Controls.Add($CloseButton)

#--[ Instruction Box ]-------------------------------------------------------------
#$BoxLength = 673  #--[ NOTE: This box size sets the form width while form autosize is enabled ]--
$BoxLength = 400
$LineLoc = 27
$InfoBox = New-Object System.Windows.Forms.Label 
$InfoBox.Location = New-Object System.Drawing.Point(($FormHCenter-($BoxLength/2)-11),$LineLoc)
$InfoBox.Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
$InfoBox.ForeColor = "DarkCyan"
$InfoBox.Font = $ButtonFont
$InfoBox.Text = "Select a site below to run connectivity tests against:"
$InfoBox.TextAlign = 2 
$Form.Controls.Add($InfoBox) 

#==[ Define Site Information Group ]==================================================    

#--[ Address Label ]-------------------------------------------------------------------
$BoxLength = 100
$LineLoc = 21 
$AddressLabel = New-Object System.Windows.Forms.TextBox
$AddressLabel.Location = New-Object System.Drawing.Size((($FormHCenter-($BoxLength/2))-300),$LineLoc)
$AddressLabel.Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
$AddressLabel.Text = "Site Address:"
$AddressLabel.Enabled = $True
$AddressLabel.ReadOnly = $True
$AddressLabel.TextAlign = 2
$Form.Controls.Add($AddressLabel) 

#--[ Address Information Box ]-------------------------------------------------------------------
$BoxLength = 544  
$AddressBox = New-Object System.Windows.Forms.TextBox
$AddressBox.Location = New-Object System.Drawing.Size((($FormHCenter-($BoxLength/2))+30),$LineLoc)
$AddressBox.Size = New-Object System.Drawing.Size(($BoxLength-9),$TextHeight) 
$AddressBox.BackColor = $Form.BackColor
$AddressBox.ForeColor = "lightgray"
$AddressBox.Text = "-address-"
$AddressBox.Enabled = $True
$AddressBox.ReadOnly = $True
$AddressBox.TextAlign = 2
$Form.Controls.Add($AddressBox)

#--[ Contact Label ]-------------------------------------------------------------------
$BoxLength = 100
$BoxLeft = 10
$LineLoc = $LineLoc+$LineAdd
$ContactLabel = New-Object System.Windows.Forms.TextBox
$ContactLabel.Location = New-Object System.Drawing.Size((($FormHCenter-($BoxLength/2))-300),$LineLoc)
$ContactLabel.Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
$ContactLabel.Text = "Site Contact:"
$ContactLabel.Enabled = $True
$ContactLabel.ReadOnly = $True
$ContactLabel.TextAlign = 2
$Form.Controls.Add($ContactLabel) #>

#--[ User ID Text Input Box ]-------------------------------------------------------------
$BoxLength = 150
$BoxLeft = $BoxLeft+($BoxLength-18)
$ContactTextBox = New-Object System.Windows.Forms.TextBox 
$ContactTextBox.Location = New-Object System.Drawing.Size($BoxLeft,90) #$LineLoc)
$ContactTextBox.Size = New-Object System.Drawing.Size(($BoxLength),$TextHeight) 
$ContactTextBox.TabIndex = 2
$ContactTextBox.BackColor = $Form.BackColor
$ContactTextBox.ForeColor = "lightgray"
$ContactTextBox.ReadOnly = $true
$ContactTextBox.Text = "-name-" 
$ContactTextBox.TextAlign = 2
$ContactTextBox.Enabled = $True
$Form.Controls.Add($ContactTextBox) 

#--[ Phone Label ]-------------------------------------------------------------------
$BoxLength = 140
$BoxLeft = $BoxLeft+($BoxLength-1)
$EmailBox = New-Object System.Windows.Forms.TextBox
$EmailBox.Location = New-Object System.Drawing.Size($BoxLeft,$LineLoc)
$EmailBox.Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
$EmailBox.BackColor = $Form.BackColor
$EmailBox.ForeColor = "lightgray"
$EmailBox.Text = "-email-"
$EmailBox.Enabled = $True
$EmailBox.ReadOnly = $True
$EmailBox.TextAlign = 2
$Form.Controls.Add($EmailBox) #>

#--[ Text Input Box ]-------------------------------------------------------------
$BoxLength = 110
$BoxLeft = $BoxLeft+($BoxLength+38)
$DeskPhoneBox = New-Object System.Windows.Forms.TextBox 
$DeskPhoneBox.Location = New-Object System.Drawing.Point($BoxLeft,$LineLoc)
$DeskPhoneBox.Size = New-Object System.Drawing.Size(($BoxLength),$TextHeight) 
$DeskPhoneBox.TabIndex = 2
$DeskPhoneBox.ReadOnly = $true
$DeskPhoneBox.BackColor = $Form.BackColor
$DeskPhoneBox.ForeColor = "lightgray"
$DeskPhoneBox.Text = "-phone 1-" 
$DeskPhoneBox.TextAlign = 2
$DeskPhoneBox.Enabled = $True
$Form.Controls.Add($DeskPhoneBox) 

$BoxLeft = $BoxLeft+($BoxLength+8)
$CellPhoneBox = New-Object System.Windows.Forms.TextBox
$CellPhoneBox.Location = New-Object System.Drawing.Point($BoxLeft,$LineLoc)
$CellPhoneBox.Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
$CellPhoneBox.BackColor = $Form.BackColor
$CellPhoneBox.ForeColor = "lightgray"
$CellPhoneBox.Text = "-phone 2-"
$CellPhoneBox.Enabled = $true
$CellPhoneBox.ReadOnly = $True
$CellPhoneBox.TextAlign = 2
$Form.Controls.Add($CellPhoneBox) #>

$Range = @($AddressLabel,$AddressBox,$ContactLabel,$EmailBox,$CellPhoneBox,$DeskPhoneBox)
$LocationGroupBox = New-Object System.Windows.Forms.GroupBox
$LocationGroupBox.Location = New-Object System.Drawing.Point((($FormHCenter/2)-162),$LineLoc) 
$LocationGroupBox.size = '670,72'
$LocationGroupBox.AutoSize = $False
$LocationGroupBox.text = "Location Detail:"    
$LocationGroupBox.Controls.AddRange($Range)
$Form.controls.add($LocationGroupBox)

#==[ Define Initial Result Area Group ]=================================================
$ResultLines = @{}
$Counter = 1
ForEach ($Item in $ExtOption.TargetTemplate.Split(';')){
    $ResultLines.Add($Counter, $Item)
    $Counter++
}

#--[ Result Box Group Base Settings ]--
$BoxLength = 100
$LineLoc = 20
$LineRight = 5
$LineAdd = 29 
$Counter = 0
$Range = @()

While ($Counter -lt $ResultLines.Count){
    $Counter++    
    #--[ Left Info Box ]-------------------------------------------------------------------
    $Box = New-Object System.Windows.Forms.TextBox -Property @{
        Location = New-Object System.Drawing.Size($LineRight,$LineLoc)
        Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
        Text = $ResultLines[$Counter]
        Enabled = $True
        ReadOnly = $true
        BackColor = $Form.BackColor
        TextAlign = 2
    }    
    Remove-Variable -Name ("Line"+$Counter+"Left") -ErrorAction SilentlyContinue
    New-Variable -Name ("Line"+$Counter+"Left" ) -Value $Box 
    $Form.Controls.Add($Box) 
    $Range += $Box
  
    #--[ Center Info Box ]-------------------------------------------------------------------
    If ($ResultLines[$Counter] -like "*Circuit*"){
        $Length = $BoxLength+115
    }Else{
        $Length = $BoxLength
    }
    $Box = New-Object System.Windows.Forms.TextBox -Property @{
        Location = New-Object System.Drawing.Size(($LineRight+115),$LineLoc)
        Size = New-Object System.Drawing.Size($Length,$TextHeight)         
        Text = "--"
        Enabled = $True
        ReadOnly = $true
        BackColor = $Form.BackColor
        TextAlign = 2
    }
    Remove-Variable -Name ("Line"+$Counter+"Center") -ErrorAction SilentlyContinue
    New-Variable -Name ("Line"+$Counter+"Center" ) -Value $Box
    $Form.Controls.Add($Box) 
    $Range += $Box

    #--[ Right Info Box ]-------------------------------------------------------------------
    $Box = New-Object System.Windows.Forms.TextBox -Property @{
        Location = New-Object System.Drawing.Size(($LineRight+230),$LineLoc)
        Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
        Text = "--"
        Enabled = $True
        ReadOnly = $true
        TextAlign = 2
    }
    Remove-Variable -Name ("Line"+$Counter+"Right") -ErrorAction SilentlyContinue
    New-Variable -Name ("Line"+$Counter+"Right" ) -Value $Box
    $Form.Controls.Add($Box) 
    $Range += $Box
    $LineLoc = $LineLoc+$LineAdd 
}

$ResultGroupBox = New-Object System.Windows.Forms.GroupBox
$ResultGroupBox.Location = New-Object System.Drawing.Size(($FormHCenter-14),123) 
$ResultGroupBox.text = "Results:"
$ResultGroupBox.AutoSize = $true
$ResultGroupBox.Controls.AddRange($Range)
$Form.controls.add($ResultGroupBox)

#==[ Dynamically Create Location Radio Buttons from XML Contents ]=========================
#--[ Radio Button Group Base Settings ]--
$CbLeft = 17
$CbRight = 163
$CbHeight = 22
$CbVar = 20
$CbBox = 145
$Count = 0
$Range = @()

While ($SiteList.Count -gt $Count) {  
    Remove-Variable -Name "RadioButton$Count" -ErrorAction SilentlyContinue
    New-Variable -Name "RadioButton$Count" -value $RadBox
    $RadBox = Get-Variable -name "RadioButton$Count" -ValueOnly    
    $RadBox = new-object System.Windows.Forms.radiobutton
    $RadBox.Text = $SiteList[$Count]

    if (0,2,4,6,8 -contains "$Count"[-1]-48) {
        $RadBox.Location = new-object System.Drawing.Size($CbLeft,$CbHeight)
        $Count++
    }Else{
        $RadBox.Location = new-object System.Drawing.Size($CbRight,$CbHeight)
        $CbHeight = $CbHeight+$CbVar
        $Count++
    }
    $RadBox.Size = new-object System.Drawing.Size($CbBox,$TextHeight)
    $RadBox.Enabled = $true 
    $RadBox.Add_Click({
        $ProcessButton.Enabled = $true
        #==[ Create and load site specific PSObject ]===================
        If ($ExtOption.ConsoleState){  #--[ Optional output for debugging ]--
            StatusMsg "Finding and loading site specifc data." "Magenta" $ExtOption
        }
        $Script:SiteDetails = LoadConfig $ConfigFile $This.text
        #===============================================================
        $Text = "You have selected: "+$SiteDetails.Name
        $AddressBox.ForeColor = "Black"
        $AddressBox.Text = $SiteDetails.Address
        $InfoBox.Text = $Text
        $InfoBox.ForeColor = "DarkGreen"
        $ContactTextBox.ForeColor = "Black"
        $ContactTextBox.Text = $SiteDetails.Contact
        $CellphoneBox.ForeColor = "Black"
        $CellPhoneBox.Text = $SiteDetails.CellPhone
        $DeskPhoneBox.ForeColor = "Black"
        $DeskPhoneBox.Text = $SiteDetails.DeskPhone
        $EmailBox.ForeColor = "Black"
        $EmailBox.Text = $SiteDetails.Email
        $Marker = $ExtOption.TargetTemplate.Split(";").count
        $Counter = 0
        StatusMsg "Populating targets." "Magenta" $ExtOption
        Foreach ($Target in $SiteDetails.PSObject.Properties){
            If (($Target.Name -eq "TargetTemplate") -or ($Target.Name -eq "TargetCount")){
                #--[ Do nothing, bypass these. ]--
            }ElseIf ($Target.Name -like "*Target*"){              
                #--[ Blank the right boxes ]--
                $Rline = 'Line'+($Counter-$Marker)+'Right'
                $Rbox = (Get-Variable  | Where-Object {$_.Name -like "*$Rline*"} ) 
                $Rbox.value.text = "" 
                #--[ Populate center box with targets ]--              
                $Cline = 'Line'+($Counter-$Marker)+'Center'
                $Cbox = (Get-Variable  | Where-Object {$_.Name -like "*$Cline*"} ) 
                $Cbox.value.text = $Target.value  
                If ($ExtOption.Debug){  #--[ Optional output for debugging ]--
                    Write-Host "Counter         : " -ForegroundColor Yellow -NoNewline
                    Write-host $Counter -ForegroundColor Cyan
                    write-host "Target Name     : " -ForegroundColor Yellow -NoNewline
                    Write-host $Target.Name.PadRight(15," ") -ForegroundColor Cyan -NoNewline
                    write-host "Target Value    : " -ForegroundColor Yellow -NoNewline
                    Write-host $Target.Value -ForegroundColor Cyan
                    write-host "Center Box Name : " -ForegroundColor Yellow -NoNewline
                    Write-host $Rbox.name.PadRight(15," ")  -ForegroundColor Cyan -NoNewline
                    write-host "Center Box Value: " -ForegroundColor Yellow -NoNewline
                    Write-host "-Blanked-`n" -ForegroundColor Cyan
                }
            }
            $Counter++
        }
        StatusMsg "Ready to execute." "Magenta" $ExtOption
        Start-Sleep -millisec 2
    })
    $Range += $RadBox
}
$SiteGroupBox = New-Object System.Windows.Forms.GroupBox
$SiteGroupBox.Location = New-Object System.Drawing.Size(($FormHCenter-344),123) 
$SiteGroupBox.text = "Location List:"
$SiteGroupBox.AutoSize = $true
$SiteGroupBox.Controls.AddRange($Range)
$form.controls.add($SiteGroupBox)

#--[ Final Form Size Dynamic Determination ]--------------------------------------
$Form.AutoSize = $False
$Form.AutoScale = $False
StatusMsg "Determining form sizes." "Magenta" $ExtOption
If ([int]$ResultGroupBox.Size.Height -ge [int]$SiteGroupBox.Size.Height){
    If ($ExtOption.Debug){  #--[ Optional debugging output for form dimensioning ]--
        write-host "- Result Box is Larger -   "  -ForegroundColor Green
        write-host "Result Box Size : " $ResultGroupBox.Size -ForegroundColor Green
        write-host "Target Count    : " $ResultLines.Count -ForegroundColor Green
    }
    $FormHeight = ($ResultGroupBox.Size.Height+180)
}Else{    
    If ($ExtOption.Debug){  #--[ Optional debugging output for form dimensioning ]--
        write-host "- Site Box is Larger -"  -ForegroundColor Green
        write-host "Site Box Size   : " $SiteGroupBox.Size -ForegroundColor Green
        write-host "Site Count      : " $SiteList.Count -ForegroundColor Green
    }
    $FormHeight = ($SiteGroupBox.Size.Height+180)
}

If ($ExtOption.Debug){  #--[ Optional debugging output for form dimensioning ]--
    write-host "Form Size       : " $Form.Size -ForegroundColor Green
    write-host "Form Bottom     : " $Form.Bottom -ForegroundColor Green
}
$Form.minimumSize = New-Object System.Drawing.Size($FormWidth,($FormHeight))
$Form.maximumSize = New-Object System.Drawing.Size($FormWidth,($FormHeight))
#----------------------------------------------------------------------------

#--[ EXECUTE Button ]--------------------------------------------------------------
$ProcessButton = new-object System.Windows.Forms.Button
$ProcessButton.Location = new-object System.Drawing.Size(($FormHCenter-($BoxLength/2)+245),$ButtonLineLoc)
$ProcessButton.Size = new-object System.Drawing.Size($BoxLength,$ButtonHeight)
$ProcessButton.Enabled = $false 
$ProcessButton.Text = "Execute"
$ProcessButton.TabIndex = 2
$ProcessButton.Add_Click({
    $Counter = 1
    StatusMsg "Cycling through each target." "Magenta" $ExtOption
    Foreach ($Target in $SiteDetails.PSObject.Properties){
        If (($Target.Name -eq "TargetTemplate") -or ($Target.Name -eq "TargetCount")){
            #--[ Do nothing, bypass these. ]--
        }ElseIf ($Target.Name -like "*Target*"){
            $Rline = 'Line'+($Counter-9)+'Right'
            $Lline = 'Line'+($Counter-9)+'Left'
            $Rbox = (Get-Variable  | Where-Object {$_.Name -like "*$Rline*"} )
            $Lbox = (Get-Variable  | Where-Object {$_.Name -like "*$Lline*"} )
            If ($Lbox.value.Text -like "*circuit*"){
                #--[ Skip this, box is a double wide circuit ID ]--
            }Else{
                $Rbox.value.ForeColor = "Black"
                $Rbox.value.text = "--Testing --"
                $Rbox.value.BackColor = $Form.BackColor
                #--[ Ping routine ]---------------------
                Try{
                    $Result = Test-Connection -ComputerName $Target.Value -count 3 -ErrorAction:Stop
                    $Latency = [int]($Result | Measure-Object -Property ResponseTime -Average).Average
                    #$z.value.Font = $ButtonFont
                    $Rbox.value.ForeColor = "Green"
                    $Rbox.value.text = "Success ($Latency ms)"#$Result
                }Catch{
                    $Rbox.value.ForeColor = "Red"
                    $Rbox.value.text = "NO RESPONSE"
                }
                #--------------------------------------
                If ($ExtOption.Debug){  #--[ Optional output for debugging ]--
                    Write-Host "Counter         : " -ForegroundColor Yellow -NoNewline
                    Write-host $Counter -ForegroundColor Cyan
                    write-host "Target Name     : " -ForegroundColor Yellow -NoNewline
                    Write-host $Target.Name.PadRight(15," ") -ForegroundColor Cyan -NoNewline
                    write-host "Target Value    : " -ForegroundColor Yellow -NoNewline
                    Write-host $Target.Value -ForegroundColor Cyan
                    write-host "Right Box Name  : " -ForegroundColor Yellow -NoNewline
                    Write-host $Rbox.name.PadRight(15," ")  -ForegroundColor Cyan -NoNewline
                    write-host "Right Box Value : " -ForegroundColor Yellow -NoNewline
                    Write-host $Rbox.Value.Text"`n" -ForegroundColor Cyan                    
                }
            }
        }
    $Counter++
    }
    StatusMsg "Ready for next site" "Magenta" $ExtOption
    Start-Sleep -millisec 2
})
$Form.Controls.Add($ProcessButton)

#--[ Open Form ]-------------------------------------------------------------
$Form.topmost = $true
$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
if($Stop -eq $true){
    $Form.Close();break;break
}
StatusMsg "Script completed.  Exiting..." "Red" $ExtOption



