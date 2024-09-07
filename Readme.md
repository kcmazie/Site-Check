[![Minimum Supported PowerShell Version][powershell-minimum]][powershell-github]

[powershell-minimum]: https://img.shields.io/badge/PowerShell-5.1+-blue.svg
[powershell-github]:  https://github.com/PowerShell/PowerShell

# $${\color{Cyan}Powershell \space "Site-Check" \space Script}$$

#### $${\color{orange}Original \space Author \space : \space \color{white}Kenneth \space C. \space Mazie \space \color{lightblue}(kcmjr \space AT \space kcmjr.com)}$$

## $${\color{grey}Description:}$$ 
This script is designed to be a quick-check of a list of locations (sites) and network IP addresses (targets). 

My initial desire was to have a tool that when a WAN circuit outage alert was received all primary addresses at the site could be rapidly checked.  I wanted it simple, quick, and easy, as well as versitile enough that I could alter the lists with minimal effort, plus in order to share it could contain no proprietary data.

By itself the script is completely generic, there are no default settings.  Configuration is managed via a companion XML config file which must be located in the same folder and named the same as as the script (but obviously with a ".xml" extension).  If the config file is NOT detected a sample XML file will be created on first run that must be edited before use.  Data from the sample is shown in the screen shots below.  

The config file has fields to contain technical contact info as well as the location address for each site.  It also has the ability to list circuit IDs in the target list for reference.

The script dynamically adjusts it's size depending on the number of sites and/or targets in the config file up to a limit after which it will exceed the screen size.  

The script uses a simple ICMP echo (ping) to identify if a target is alive. Select one of the listed sites and the associated info from the config file will populate the GUI. Each time you select a different site the info will dynamically update. Once the site is loaded simply click the Execute button to run the ping check. Results are noted
along the right side of the screen.  After each run you can select a different site and execute again 
continuously until you click the cancel button. 

## $${\color{grey}Notes:}$$ 
* Normal operation is with no command line options.
* Powershell 5.1 is the minimal version required.

## $${\color{grey}Arguments:}$$ 
Command line options for testing: 
* "-console $true" will enable local console echo for troubleshooting
* "-debug $true" will only email to the first recipient on the list

### $${\color{grey}Screenshots:}$$ 
   This is the initial GUI.
   
![Initial GUI](https://github.com/kcmazie/Site-Check/blob/main/Screenshot1.jpg "Initial GUI")

   This is the GUI after selecting a site.
   
![GUI With Site Selected](https://github.com/kcmazie/Site-Check/blob/main/Screenshot2.jpg "GUI With Site Selected")

   This is the GUI during a run (with some live IPs substituted).
   
![GUI While Test is Running](https://github.com/kcmazie/Site-Check/blob/main/Screenshot3.jpg "GUI While Test is Running")

  
### $${\color{grey}Warnings:}$$ 
* None 

### $${\color{grey}Enhancements:}$$ 
Some possible future enhancements are:
* Ability to email the results
* Still need to add some error checking when IP addresses are missing or misspelled.

### $${\color{grey}Legal:}$$ 
Public Domain. Modify and redistribute freely. No rights reserved. 
SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED. 

That being said, please let me know if you find bugs, have improved the script, or would like to help. 

### $${\color{grey}Credits:}$$  
Code snippets and/or ideas came from many sources including but not limited to the following: 
* n/a 
  
### $${\color{grey}Version \\& Change History:}$$ 
* Last Update by  : Kenneth C. Mazie 
  * Initial release : v1.00 - 09-06-24
  * Change History  : v2.00 - 00-00-00 
 
