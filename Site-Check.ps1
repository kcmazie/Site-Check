Param(
    [Switch]$Script:Console = $false           #--[ Set to true to enable local console result display. Defaults to false ]--
)
<#------------------------------------------------------------------------------ 
         File Name : Whereabouts.ps1 
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com) 
                   : 
       Description : Programatically creates an email to send to the predetermined user or group for notification about where you are.
                   : Grabs Outlook email signature from current users profile.  Determines sender by current logged on user. 
                   : 
             Notes : Normal operation is with no command line options.  
                   : See end of script for detail about how to launch via shortcut. 
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
   Version History : v1.00 - 09-09-18 - Original 
    Change History : v2.00 - 10-18-20 - Added travel log tracking and PC console locking options.
                   : v3.00 - 10-27-20 - Added test mode option. 
                   : v3.10 - 11-13-20 - Added function to assume job of updating run data
                   : v4.00 - 11-13-20 - Added option to dynamically size form depending on number of sites in site list
                   : v5.00 - 11-25-20 - Removed descrete checkbox sections and replaced with dynamic creator based 
                   :                    on site list.
                   : v5.10 - 04-28-21 - Minor form layout adjustment.
                   : v5.20 - 03-30-22 - Added 0 mile logging, reordered variable location, adjusted some console message
                   :                    text messages for clarity                
                   : v6.00 - 05-05-23 - Relocated options out to XML file for publishing.  Refactored some sections.
                   : v6.10 - 10-02-23 - Switched out WMI for CIM
                   : #>
                   $ScriptVer = "6.10"    <#--[ Current version # used in script ]--
                   : 
------------------------------------------------------------------------------#>
#Requires -Version 5.1
Clear-Host 

<#--[ Suppress Console ]-------------------------------------------------------
Add-Type -Name Window -Namespace Console -MemberDefinition ' 
[DllImport("Kernel32.dll")] 
public static extern IntPtr GetConsoleWindow(); 
 
[DllImport("user32.dll")] 
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow); 
' 
$ConsolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($ConsolePtr, 0) | Out-Null
#------------------------------------------------------------------------------#>
 
#--[ Runtime Variables ]----------------------------------------------------
#--[ For Testing ]-------------
$Script:Console = $true
$Script:Debug = $true
#------------------------------
$ErrorActionPreference = "silentlycontinue"


#$DomainName = $env:USERDOMAIN      #--[ Pulls local domain as an alternate if the user leaves it out ]-------
$UN = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name        #--[ Current User   ]--
$DN = $UN.split("\")[0]                                                     #--[ Current Domain ]--    
$SenderEmail = $UN.split("\")[1]+"@"+$DN+".org"                      #--[ Correct this for email domain, .ORG, .COM, etc ]--

#--[ Functions ]--------------------------------------------------------------

Function GetConsoleHost ($ConfigFile){  #--[ Detect if we are using a script editor or the console ]--
    Switch ($Host.Name){
        'consolehost'{
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $False -force
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell Console detected." -Force
        }
        'Windows PowerShell ISE Host'{
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell ISE editor detected.  Console mode enabled." -Force
        }
        'PrimalScriptHostImplementation'{
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "COnsoleMessage" -Value "PrimalScript or PowerShell Studio editor detected.  Console mode enabled." -Force
        }
        "Visual Studio Code Host" {
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "Visual Studio Code editor detected.  Console mode enabled. " -Force
        }
    }
    If ($ConfigFile.ConsoleState){
        StatusMsg "Detected session running from an editor..." "Cyan" $ConfigFile
    }
    Return $ConfigFile
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

<#Function GetSiteDetails ($ConfigFile,$SiteCode){
    $Site = "Site"+$SiteCode.Split(",")[1]
    [xml]$Config = Get-Content $ConfigFile           #--[ Read & Load XML ]--  
    $Target = $Config.Settings.TargetTemplate.Split(';')
    $XmlData = New-Object -TypeName psobject 
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Location" -Value $Config.Settings.$Site.Location
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Name" -Value $Config.Settings.$Site.Name
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Address" -Value $Config.Settings.$Site.Address
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Contact" -Value $Config.Settings.$Site.Contact
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Email" -Value $Config.Settings.$Site.Email
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "DeskPhone" -Value $Config.Settings.$Site.DeskPhone
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "CellPhone" -Value $Config.Settings.$Site.CellPhone
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "ResultCount" -Value $Target.$Count
    $Counter = 0
    While ($Counter -lt $Target.count){
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Target$Counter" -Value $Config.Settings.$Site.($Target[$Counter])
        $Counter++
    }
    Return $XmlData
}#>

Function Inspect ($IPAddress){
    $PingIt = test-connection -ComputerName $IPAddress -count 1 -erroraction SilentlyContinue #-asjob  | Wait-Job -Timeout 1 | receive-job 
   # $PingIt2 = Test-NetConnection -ComputerName $IPAddress -InformationLevel Detailed
   # write-host $PingIt2


    If ($PingIt ){
        $Result = $IPAddress+";Success"
    }Else{
        $Result = $IPAddress+";FAILED"
    }
    Return $Result
}

Function StatusMsg ($Msg, $Color, $ExtOption){
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    Write-Host "-- Script Status: $Msg" -ForegroundColor $Color
    $Msg = ""
}

#==[ End of Functions ]=======================================================
#=============================================================================
#=[ Begin Processing ]========================================================

#--[ Load external XML options file ]-----------------------------------------
$ConfigFile = $PSScriptRoot+"\"+($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"

If (Test-Path -Path $ConfigFile -PathType Leaf){                       #--[ Error out if configuration file doesn't exist ]--
    $ExtOption = LoadConfig $ConfigFile ""
    StatusMsg "Loading external config file..." "Magenta" $ExtOption
    $SiteList = [Ordered]@{}  
    $Index = 1
    ForEach($Site in $ExtOption.Sitelist){
        $SiteList.Add($Index, @($Site))
        $Index++
    }
    If ($Debug){  
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $true   
    }
}Else{
    StatusMsg "MISSING XML CONFIG FILE.  File is required.  Script aborted..." " Red" $True
    $Message = (
'--[ External XML config file example ]-----------------------------------
--[ To be named the same as the script and located in the same folder as the script ]--

<?xml version="1.0" encoding="utf-8"?>
<Settings>
<General>
    <SmtpServer>mailserver.company.org</SmtpServer>
    <SmtpPort>25</SmtpPort>
    <RecipientEmail>InformationTechnology@company.org</RecipientEmail>
    <SourcePath>C:\folder</SourcePath>
    <ExcelSourceFile>+NetworkedDevice-Master-Inventory.xlsx</ExcelSourceFile>
    <ExcelWorkingCopy>NetworkedDevice-Master-Inventory.xlsx</ExcelWorkingCopy>
    <Domain>company.org</Domain>
</General>
<Credentials>
    <PasswordFile>c:\AESPass.txt</PasswordFile>
    <KeyFile>c:\AESKey.txt</KeyFile>
    <WAPUser>admin</WAPUser>
    <WAPPass>wappass</WAPPass>
    <AltUser>user1</AltUser>
    <AltPass>userpass1</AltPass>
</Credentials>    
<Recipients>
    <Recipient>me@company.org</Recipient>
    <Recipient>you@company.org</Recipient>
    <Recipient>them@company.org</Recipient>
</Recipients>
</Settings> ')
Write-host $Message -ForegroundColor Yellow
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
$DeskPhoneBox.Text = "-phone 1-" 
$DeskPhoneBox.TextAlign = 2
$DeskPhoneBox.Enabled = $True
$Form.Controls.Add($DeskPhoneBox) 

$BoxLeft = $BoxLeft+($BoxLength+8)
$CellPhoneBox = New-Object System.Windows.Forms.TextBox
$CellPhoneBox.Location = New-Object System.Drawing.Point($BoxLeft,$LineLoc)
$CellPhoneBox.Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
$CellPhoneBox.Text = "-phone 2-"
$CellPhoneBox.Enabled = $True
$CellPhoneBox.ReadOnly = $True
$CellPhoneBox.TextAlign = 2
$Form.Controls.Add($CellPhoneBox) #>

$Range = @($AddressLabel,$AddressBox,$ContactLabel,$EmailBox,$CellPhoneBox,$DeskPhoneBox)
$LocationGroupBox = New-Object System.Windows.Forms.GroupBox
$LocationGroupBox.Location = New-Object System.Drawing.Point((($FormHCenter/2)-162),$LineLoc) #'35,50'
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
        $Script:SiteDetails = LoadConfig $ConfigFile $This.text
        $Text = "You have selected: "+$SiteDetails.Name
        $InfoBox.Text = $Text
        $AddressBox.Text = $SiteDetails.Address
        $InfoBox.ForeColor = "DarkGreen"
        $ContactTextBox.Text = $SiteDetails.Contact
        $CellPhoneBox.Text = $SiteDetails.CellPhone
        $DeskPhoneBox.Text = $SiteDetails.DeskPhone
        $EmailBox.Text = $SiteDetails.Email

        $Counter = 0
        $Marker = $ExtOption.TargetTemplate.Split(";").count
        Foreach ($Target in $SiteDetails.PSObject.Properties){  #--[ Cycle through XML targets for site ]--
            If ($Target.Name -like "*Target*"){                
                #--[ Blank right boxes ]--
                $Rline = 'Line'+($Counter-$Marker)+'Right'
                $Rbox = (Get-Variable  | Where-Object {$_.Name -like "*$Rline*"} ) 
                $Rbox.value.text = "" 
                #--[ Populate center box with targets ]--              
                $Cline = 'Line'+($Counter-$Marker)+'Center'
                $Cbox = (Get-Variable  | Where-Object {$_.Name -like "*$Cline*"} ) 
                $Cbox.value.text = $Target.value                
                If ($ExtOption.Debug){  #--[ Optional debugging output for form layout ]--
                    Write-Host "`nCounter        : "($Counter-($ExtOption.TargetCount)) -ForegroundColor Yellow
                    write-host "Target Name    : "$($Target.Name) -ForegroundColor Yellow
                    write-host "Target Value   : "$($Target.Value)  -ForegroundColor Yellow  
                    write-host "Box Name       : "$Cbox.name -ForegroundColor cyan
                    write-host "Box Value      : "$Cbox.Value.text -ForegroundColor cyan
                }
            }
            $Counter++
        }
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

If ([int]$ResultGroupBox.Size.Height -ge [int]$SiteGroupBox.Size.Height){
    If ($ExtOption.Debug){  #--[ Optional debugging output for form dimensioning ]--
        write-host "`nResult Box is Larger    "  -ForegroundColor Green
        write-host "Result Box Size    " $ResultGroupBox.Size -ForegroundColor Green
        write-host "Result Count       " $ResultLines.Count -ForegroundColor Green
    }
    $FormHeight = ($ResultGroupBox.Size.Height+180)
}Else{    
    If ($ExtOption.Debug){  #--[ Optional debugging output for form dimensioning ]--
        write-host "`nSite Box is Larger    "  -ForegroundColor Green
        write-host "Site Box Size      " $SiteGroupBox.Size -ForegroundColor Green
        write-host  "Site Count        " $SiteList.Count -ForegroundColor Green
    }
    $FormHeight = ($SiteGroupBox.Size.Height+180)
}

If ($ExtOption.Debug){  #--[ Optional debugging output for form dimensioning ]--
    write-host "`nForm size          " $Form.Size -ForegroundColor Magenta
    write-host "Form bottom        " $Form.Bottom -ForegroundColor Magenta
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
    Foreach ($Property in $SiteDetails.PSObject.Properties){
        If ($Property.Name -like "*Target*"){
            If ($ExtOption.Debug){  #--[ Optional debugging output for form dimensioning ]--
                Write-Host "`nCounter        : "$Counter -ForegroundColor Yellow
                write-host "Target Name    : "$($Property.Name) -ForegroundColor Yellow
                write-host "Target Value   : " $($Property.Value)  -ForegroundColor Yellow
            }
            $Rline = 'Line'+($Counter-9)+'Right'
            $Lline = 'Line'+($Counter-9)+'Left'
            $Rbox = (Get-Variable  | Where-Object {$_.Name -like "*$Rline*"} )
            $Lbox = (Get-Variable  | Where-Object {$_.Name -like "*$Lline*"} )
            If ($Lbox.value.Text -like "*circuit*"){
                #--[ Skip this, box is 2 wide circuit ID ]--
            }Else{
                $Rbox.value.text = "--Testing --"
                $Rbox.value.BackColor = $Form.BackColor
                $Result = Inspect $Property.value
                If ($Result -like "*FAIL*"){
                    #$z.value.Font = $ButtonFont
                    $Rbox.value.ForeColor = "Red"
                    $Rbox.value.text = "NO RESPONSE"
                }Else{
                    $Rbox.value.ForeColor = "Green"
                    $Rbox.value.text = "- ONLINE -"
                }
                If ($ExtOption.Debug){  #--[ Optional debugging output for form dimensioning ]--
                    write-host "Right Box Name  : "$Rbox.name -ForegroundColor cyan
                    write-host "Right Box Value : "$Rbox.Value -ForegroundColor cyan
                }
            }
        }
    $Counter++
    }
    Start-Sleep -millisec 2

})
$Form.Controls.Add($ProcessButton)

#--[ Open Form ]-------------------------------------------------------------
$Form.topmost = $true
$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
if($Stop -eq $true){$Form.Close();break;break}



