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
<#PSScriptInfo 
.VERSION 6.10 
.AUTHOR Kenneth C. Mazie (kcmjr AT kcmjr.com) 
.DESCRIPTION 
Programatically creates an email to send to the predetermined user or group for notification about where you are.
Grabs Outlook email signature from current users profile.  Determines sender by current logged on user.  Optionally
writes a travel log to users Documents folder and optionally will lock the PC.  Dynamically grows or shrinks the
form depending on the number of sites in the site table.
#>
#Requires -Version 5.1

Clear-Host 

#--[ For Testing ]-------------
$Script:Console = $true
#$Script:Debug = $true
#------------------------------

#--[ Suppress Console ]-------------------------------------------------------
Add-Type -Name Window -Namespace Console -MemberDefinition ' 
[DllImport("Kernel32.dll")] 
public static extern IntPtr GetConsoleWindow(); 
 
[DllImport("user32.dll")] 
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow); 
' 
$ConsolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($ConsolePtr, 0) | Out-Null
#------------------------------------------------------------------------------#
 
#--[ Runtime Variables ]----------------------------------------------------
$ErrorActionPreference = "silentlycontinue"
$Icon = [System.Drawing.SystemIcons]::Information
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0] 
$ConfigFile = $PSScriptRoot+'\'+($ScriptName.Split("_")[0])+'.xml'
#$DomainName = $env:USERDOMAIN      #--[ Pulls local domain as an alternate if the user leaves it out ]-------
$UN = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name        #--[ Current User   ]--
$DN = $UN.split("\")[0]                                                     #--[ Current Domain ]--    
$SenderEmail = $UN.split("\")[1]+"@"+$DN+".org"                      #--[ Correct this for email domain, .ORG, .COM, etc ]--

#--[ Functions ]--------------------------------------------------------------
<#Function UpdateData ($ItemNum,$SiteList){  #--[ Adds checkbox selected locations to the email and log file ]--
    $Message += "- "+($SiteList["$ItemNum"])[0]+"<br>"
    If ((($SiteList["$ItemNum"])[1] -eq 0) -And ($LogZero)){  #--[ Forces locations with a zero milage to get logged ]--
        $RunData += "`n"+(Get-Date).toShortDateString()+","+($SiteList["$ItemNum"])[0]+","+($SiteList["$ItemNum"])[1]
    }ElseIf (($SiteList["$ItemNum"])[1] -gt 0){
        $RunData += "`n"+(Get-Date).toShortDateString()+","+($SiteList["$ItemNum"])[0]+","+($SiteList["$ItemNum"])[1]
    }
}#>
Function ReloadForm {
    $Form.Close()
    $Form.Dispose()
    ActivateForm
    $Stop = $True
}

Function KillForm {
    $Form.Close()
    $Form.Dispose()
    $Stop = $True
}
Function UpdateOutput {  #--[ Refreshes the infobox contents ]--
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

Function GetSiteDetails ($ConfigFile,$SiteCode){
    $Site = "Site"+$SiteCode.Split(",")[1]
    [xml]$Config = Get-Content $ConfigFile           #--[ Read & Load XML ]--  


    $Target = $Config.Settings.ResultTemplate.Split(';')
#clear-host
#write-host "Total Target IPs/Circuits: "$Target.count

    ForEach ($x in $Target){
    #    write-host "  Target: "$x
    }
    
   

    $XmlData = New-Object -TypeName psobject 
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Location" -Value $Config.Settings.$Site.Location
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Name" -Value $Config.Settings.$Site.Name
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Address" -Value $Config.Settings.$Site.Address
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Contact" -Value $Config.Settings.$Site.Contact
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Email" -Value $Config.Settings.$Site.Email
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "DeskPhone" -Value $Config.Settings.$Site.DeskPhone
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "CellPhone" -Value $Config.Settings.$Site.CellPhone
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "ResultCount" -Value $Target.$Count
   # $XmlData | Add-Member -Force -MemberType NoteProperty -Name "ResultTemplate" -Value $Config.Settings.ResultTemplate
   $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $true

    
    $Counter = 0
   While ($Counter -lt $Target.count){
  #  write-host $Counter"    " -ForegroundColor red -NoNewline
  #  write-host $Config.Settings.$Site.($Target[$Counter]) -ForegroundColor cyan
     #   $XmlData | Add-Member -Force -MemberType NoteProperty -Name $Target[$Counter] -Value $Config.Settings.$Site.($Target[$Counter])
        $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Target$Counter" -Value $Config.Settings.$Site.($Target[$Counter])
        $Counter++
    }

<#
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Circuit1" -Value $Config.Settings.$Site.Circuit1
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "Circuit2" -Value $Config.Settings.$Site.Circuit2
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "LoopbackIP" -Value $Config.Settings.$Site.LoopbackIP
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "GatewayIP" -Value $Config.Settings.$Site.GatewayIP
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "EC1PrivateIP" -Value $Config.Settings.$Site.EC1PrivateIP
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "EC1PublicIP" -Value $Config.Settings.$Site.EC1PublicIP
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "EC2PrivateIP" -Value $Config.Settings.$Site.EC2PrivateIP    
    $XmlData | Add-Member -Force -MemberType NoteProperty -Name "EC2PublicIP" -Value $Config.Settings.$Site.EC2PublicIP    
    #>
    
 
    Return $XmlData
}

Function Inspect ($IPAddress){
    $PingIt = test-connection -ComputerName $IPAddress -count 1 -erroraction SilentlyContinue #-asjob  | Wait-Job -Timeout 1 | receive-job  
    If ($PingIt ){
        $Result = $IPAddress+";Success"
    }Else{
        $Result = $IPAddress+";- FAILED -"
    }
    Return $Result
}

Function Inspect2 ($IPAddress,$Line,$Form){
#Inspect $SiteDetails.GatewayIP $Line2Center.Text $Form
$PingIt = test-connection -ComputerName $IPAddress -count 1 -erroraction SilentlyContinue #-asjob  | Wait-Job -Timeout 1 | receive-job  
$Form.$line.Text = "- Working -"

If ($PingIt ){
    $Form.$line.Text = "Success"
    $Result = $IPAddress+";Success"
}Else{
    $Form.$line.Text = "-- FAILED --"
    $Result = $IPAddress+";- FAILED -"
}
Return $Result
}

Function StatusMsg ($Msg, $Color){
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    If ($Script:Debug -or $Script:Console){Write-Host "-- Script Status: $Msg" -ForegroundColor $Color}
    $Msg = ""
}

#==[ End of Functions ]=====================================================

#--[ Read and load configuration file ]-------------------------------------
if (!(Test-Path $ConfigFile)){                       #--[ Error out if configuration file doesn't exist ]--
    StatusMsg "MISSING CONFIG FILE.  Script aborted." " Red"
    break;break;break
}else{
    $ErrorActionPreference = "stop"
    [xml]$Configuration = Get-Content $ConfigFile  #--[ Read & Load XML ]--    
    $TitleText = $Configuration.Settings.General.TitleText
    $SmtpServer = $Configuration.Settings.General.SmtpServer
    $SmtpPort = $Configuration.Settings.General.SmtpPort
    $RecipientEmail = $Configuration.Settings.General.RecipientEmail
    $ResultTemplate = $Configuration.Settings.ResultTemplate
   
    $SiteList = [Ordered]@{}  
    $Index = 1
    ForEach($Site in $Configuration.Settings.Sites.site){
        $SiteList.Add($Index, @($Site))
        $Index++
    }
    #--[ Email Recipient Options ]--------------------------------------------------
    $Recipients = @()   #--[ List of recipients in case a group can't be used ]--
    $Recipients +="maziekc@ah.org"
    If (!($Script:Debug)){     #--[ Use to block remaining recipients for test mode routing to sender ]--
        ForEach($Recipient in $Configuration.Settings.Recipients.Recipient){
            $Recipients +=$Recipient
        }
    }
}


#--[ Prep GUI ]------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

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
    [int]$FormHeight = ((($SiteList.Count/2)*20)+255)   #--[ Dynamically Created Variable for Box Size (Even count) ]--
}Else{
    [int]$FormHeight = ((($SiteList.Count/2)*23)+255)   #--[ Dynamically Created Variable for Box Size (Odd count) ]--
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
$Form.Text = "$ScriptName v$ScriptVer"
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
$FormLabelBox.Text = $TitleText
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
    KillForm
})
$Form.Controls.Add($CloseButton)

#--[ Instruction Box ]-------------------------------------------------------------
#$BoxLength = 673  #--[ NOTE: This box size sets the form width while form autosize is enabled ]--
$BoxLength = 400
$LineLoc = 27
$UserCredLabel = New-Object System.Windows.Forms.Label 
$UserCredLabel.Location = New-Object System.Drawing.Point(($FormHCenter-($BoxLength/2)-11),$LineLoc)
$UserCredLabel.Size = New-Object System.Drawing.Size($BoxLength,$TextHeight) 
$UserCredLabel.ForeColor = "DarkCyan"
$UserCredLabel.Font = $ButtonFont
$UserCredLabel.Text = "Select a site below to run connectivity tests against:"
$UserCredLabel.TextAlign = 2 
$Form.Controls.Add($UserCredLabel) 



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
$AddressBox.Text = ""
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
$CellPhoneBox.Text = "-phone2-"
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
ForEach ($Item in $ResultTemplate.Split(';')){
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
$ResultGroupBox.Location = New-Object System.Drawing.Size(($FormHCenter-14),123) #$LineLoc)
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

        $Script:SiteDetails = GetSiteDetails $ConfigFile $This.text
        
        $Text = "You have selected: "+$SiteDetails.Name
        $UserCredLabel.Text= $Text
        $AddressBox.Text = $SiteDetails.Address
        $UserCredLabel.ForeColor = "DarkGreen"
        $ContactTextBox.Text = $SiteDetails.Contact
        $CellPhoneBox.Text = $SiteDetails.CellPhone
        $DeskPhoneBox.Text = $SiteDetails.DeskPhone
        $EmailBox.Text = $SiteDetails.Email


#Get-Variable | Out-String
#Get-Variable | ForEach-Object { "Name : {0}`r`nValue: {1}`r`n" -f $_.Name,$_.Value }

#Get-Variable | Where-Object {$_.Name -like ('*Line*')} 

        $ErrorActionPreference = "stop" #ilentlycontinue"
        clear-host
        $Counter = 1
          Write-Host  "-----------------------------------------"
        Foreach ($Property in $SiteDetails.PSObject.Properties){ #} | where {$_.Name -like "Target*"}){
            #Write-Host "$Counter   $($Property.Name): $($Property.Value) " -ForegroundColor Yellow

            #New-Variable -Name "Line"+$Counter+"Center" -Value 98033
            If ($Property.Name -like "*Target*"){
              
                Write-Host "`ncounter : "$Counter -ForegroundColor Yellow
                write-host "name    : "$($Property.Name) -ForegroundColor Yellow
                write-host "value   : " $($Property.Value)  -ForegroundColor Yellow
                
                               
                Get-Variable |%{ "Name : {0}`r`nValue: {1}`r`n" -f $_.Name,$_.Value }
                $x = 'Line'+($Counter-9)+'Center'
                $z = (Get-Variable  | Where-Object {$_.Name -like "*$x*"} ) #| set-variable $_.text = $Property.Value
                #write-host ""#$z
                write-host "variable  : "$z.name -ForegroundColor cyan
                write-host "value     : "$z.Value -ForegroundColor cyan
                set-variable -name $($z.Name) -value $Property.value -force -PassThru
#$z.GetType()



#write-host $z.text # = $Property.Value

              #  $(($Line+$Counter+"Center")).Text = $SiteDetails.LoopbackIP 
              #  $(($Line+$Counter+"Right")).Text = "--"



           <# Switch ($($Property.Name)){
                "LoopbackIP" {
                    $Line1Center.Text = $SiteDetails.LoopbackIP 
                    $Line1Right.Text = "--"
                }
                "GatewayIP"{
                    $Line2Center.Text = $SiteDetails.GatewayIP
                    $Line2Right.Text = "--"
                }
                "Circuit1"{
                    $Line3Center.Text = $SiteDetails.Circuit1
                    $Line3Right.Text = "--"
                }
                "EC1PrivateIP"{
                    $Line4Center.Text = $SiteDetails.EC1PrivateIP
                    $Line4Right.Text = "--"
                }
                "EC1PublicIP"{
                    $Line5Center.Text = $SiteDetails.EC1PublicIP
                    $Line5Right.Text = "--"
                }
                "Circuit2"{
                    $Line6Center.Text = $SiteDetails.Circuit2
                    $Line6Right.Text = "--"
                }
                "EC2PrivateIP"{
                    $Line7Center.Text = $SiteDetails.EC2PrivateIP      
                    $Line7Right.Text = "--"
                }
                "EC2PublicIP"{
                    $Line8Center.Text = $SiteDetails.EC2PublicIP
                    $Line8Right.Text = "--"
                }
            }#>
            }
            $Counter++
        }
        Start-Sleep -millisec 2
    })
    $Range += $RadBox
}
$SiteGroupBox = New-Object System.Windows.Forms.GroupBox
$SiteGroupBox.Location = New-Object System.Drawing.Size(($FormHCenter-344),123) #$LineLoc)
$SiteGroupBox.text = "Location List:"
$SiteGroupBox.AutoSize = $true
$SiteGroupBox.Controls.AddRange($Range)
$form.controls.add($SiteGroupBox)



#--[ Final Form Size Determination ]--------------------------------------
$Form.AutoSize = $False
$Form.autoscale = $false

If ([int]$ResultGroupBox.Size.Height -ge [int]$SiteGroupBox.Size.Height){
    If ($ExtOption.Debug){
        write-host "`nresult box is larger    " 
        write-host "`nresult box size    " $ResultGroupBox.Size
        write-host  "result count      " $ResultLines.Count
    }
    $FormHeight = ($ResultGroupBox.Size.Height+180)
}Else{    
    If ($ExtOption.Debug){
        write-host "`nsite box is larger    " 
        write-host "`nsite box size      " $SiteGroupBox.Size
        write-host  "site count        " $SiteList.Count
    }
    $FormHeight = ($SiteGroupBox.Size.Height+180)
}

If ($ExtOption.Debug){
    write-host "`nform size          " $Form.Size
    write-host "form bottom        " $Form.Bottom
}
$Form.minimumSize = New-Object System.Drawing.Size($FormWidth,($FormHeight))
$Form.maximumSize = New-Object System.Drawing.Size($FormWidth,($FormHeight))
#----------------------------------------------------------------------------


#--[ EXECUTE Button ]------------------------------------------------------------------------
#$LineLoc = 11  --[ Set at top of form definition at cancel box ]--
$ProcessButton = new-object System.Windows.Forms.Button
$ProcessButton.Location = new-object System.Drawing.Size(($FormHCenter-($BoxLength/2)+245),$ButtonLineLoc)
$ProcessButton.Size = new-object System.Drawing.Size($BoxLength,$ButtonHeight)
$ProcessButton.Enabled = $false 
$ProcessButton.Text = "Execute"
$ProcessButton.TabIndex = 2
$ProcessButton.Add_Click({
Foreach ($Property in $SiteDetails.PSObject.Properties){
    If ($ExtOption.Debug){
        Write-Host "$($Property.Name): $($Property.Value)" -ForegroundColor Cyan
    }
    Switch ($($Property.Name)){
        "LoopbackIP" {
            $Line1Right.Text = "- Working -"
            $Result = Inspect $SiteDetails.LoopbackIP  
            $Line1Right.BackColor = $Form.BackColor
            $Line1Right.Text = $Result.Split(";")[0]
            If ($Result -like "*FAIL*"){
                $Line1Right.Font = $ButtonFont
                $Line1Right.ForeColor = 'Red'
                $Line1Right.Text = "- OFFLINE -"
            }Else{
                $Line1Right.ForeColor = 'Green'
                $Line1Right.Text = "- ONLINE -"
            }
        }
        "GatewayIP"{
            $Line2Right.Text = "- Working -"
            $Result = Inspect $SiteDetails.GatewayIP  
            $Line2Right.BackColor = $Form.BackColor
            $Line2Right.Text = $Result.Split(";")[0]
            If ($Result -like "*FAIL*"){
                $Line2Right.Font = $ButtonFont
                $Line2Right.ForeColor = 'Red'
                $Line2Right.Text = "- OFFLINE -"
            }Else{
                $Line2Right.ForeColor = 'Green'
                $Line2Right.Text = "- ONLINE -"
            }
        }
        "Circuit1"{
#            Inspect $SiteDetails.GatewayIP $Line2Center.Text $Form
           # $Line3Center.Text = $SiteDetails.Circuit1
            #$Result = Inspect $SiteDetails.Circuit1
        }
        "EC1PrivateIP"{
            $Line4Right.Text = "- Working -"
            $Result = Inspect $SiteDetails.EC1PrivateIP
            $Line4Right.BackColor = $Form.BackColor
            $Line4Right.Text = $Result.Split(";")[0]
            If ($Result -like "*FAIL*"){
                $Line4Right.Font = $ButtonFont
                $Line4Right.ForeColor = 'Red'
                $Line4Right.Text = "- OFFLINE -"
            }Else{
                $Line4Right.ForeColor = 'Green'
                $Line4Right.Text = "- ONLINE -"
            }
        }
        "EC1PublicIP"{
            $Line5Right.Text = "- Working -"        
            $Result = Inspect $SiteDetails.EC1PublicIP
            $Line5Right.BackColor = $Form.BackColor
            $Line5Right.Text = $Result.Split(";")[0]
            If ($Result -like "*FAIL*"){
                $Line5Right.Font = $ButtonFont
                $Line5Right.ForeColor = 'Red'
                $Line5Right.Text = "- OFFLINE -"
            }Else{
                $Line5Right.ForeColor = 'Green'
                $Line5Right.Text = "- ONLINE -"
            }
        }
        "Circuit2"{
            #$Line6Center.Text = $SiteDetails.Circuit2
            #$Result = Inspect $SiteDetails.Circuit2
        }
        "EC2PrivateIP"{
            $Line7Right.Text = "- Working -"            
            $Result = Inspect $SiteDetails.EC2PrivateIP
            $Line7Right.BackColor = $Form.BackColor
            If ($Result -like "*FAIL*"){
                $Line7Right.Font = $ButtonFont
                $Line7Right.ForeColor = 'Red'
                $Line7Right.Text = "- OFFLINE -"
            }Else{
                $Line7Right.ForeColor = 'Green'
                $Line7Right.Text = "- ONLINE -"
            }     
        }
        "EC2PublicIP"{
            $Line8Right.Text = "- Working -"            
            $Result = Inspect $SiteDetails.EC2PublicIP
            $Line8Right.BackColor = $Form.BackColor
            $Line8Right.Text = $Result.Split(";")[0]
            If ($Result -like "*FAIL*"){
               # $Line8Right.Font = $ButtonFont
                $Line8Right.ForeColor = 'Red'
                $Line8Right.Text = "- OFFLINE -"
            }Else{
                $Line8Right.ForeColor = 'Green'
                $Line8Right.Text = "- ONLINE -"
            }
        }
    }
}

})
$Form.Controls.Add($ProcessButton)



#--[ Open Form ]-------------------------------------------------------------
$Form.topmost = $true
$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
if($Stop -eq $true){$Form.Close();break;break}



