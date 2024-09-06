# Site-Check
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
 
