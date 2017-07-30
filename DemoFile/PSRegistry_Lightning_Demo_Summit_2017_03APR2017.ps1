# PowerShell Summit 2017

# PSRegistry Demo

# Traditional way of loading registry of all users on a system

$psISE.CurrentPowerShellTab.Files.Add("C:\_Github\CommunityLightningDemos2017\DemoFile\RegistryDemo_OLD.ps1")

# PSRegistry

# We can mount a unloaded user hive by referencing a path to the users NTSUSER.DAT file
$RegMount = Mount-RegistryHive -NTUSER C:\Users\defaultuser0\NTUSER.DAT -KeyType HKU

# We can add a new registry key
New-RegistryKey -KeyType HKU -KeyPath "$($RegMount.MountKey)\SOFTWARE\Test\Path"

# We can also remove that key or any other key from a mounted hive
Remove-RegistryKey -KeyType HKU -KeyPath "$($RegMount.MountKey)\SOFTWARE\Test\Path"

# we can also just specify the removal of a specific key value
Remove-RegistryKeyValue -KeyType HKU -KeyPath "$($RegMount.MountKey)\SOFTWARE\Test\Path" -Value 'Test'

# Once we are done, we can dismount the loaded registry hive by referencing the GUID 
# returned by Mount-RegistryHive
Dismount-RegistryHive -MountKey $RegMount.MountKey -KeyType HKU


# Work in Progress - Needs further testing
# Mounting a remote registry hive
Mount-RemoteRegistryHive -Computer 'Computer01' -KeyType HKU

# Creating a new registry key value



# still a lot to do
#
# 
# If you are interested in contributing or just want to ask questions, hit me up:



#################################################
#                                               #
# Name: Josh Rickard                            #
# Blog: MSAdministrator.com                     #
# GitHub: https://github.com/MSAdministrator    #
# Twitter: @MS_dministrator                     #
#                                               #
#################################################

# BONUS!  Did you know that your twitter handle can NOT 
# have the word "Admin" or "Administrator" in it
#
# FTW











Write-Host -Object ("")
Write-Host -Object ("")
Write-Host -Object ("         ,.=:^!^!t3Z3z.,                  ") -ForegroundColor Red
Write-Host -Object ("        :tt:::tt333EE3                    ") -ForegroundColor Red
Write-Host -Object ("        Et:::ztt33EEE ") -NoNewline -ForegroundColor Red
Write-Host -Object (" @Ee.,      ..,     $Date") -ForegroundColor Green
Write-Host -Object ("       ;tt:::tt333EE7") -NoNewline -ForegroundColor Red
Write-Host -Object (" ;EEEEEEttttt33#     ") -ForegroundColor Green
Write-Host -Object ("      :Et:::zt333EEQ.") -NoNewline -ForegroundColor Red
Write-Host -Object (" SEEEEEttttt33QL     ") -NoNewline -ForegroundColor Green
Write-Host -Object ("      ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("      it::::tt333EEF") -NoNewline -ForegroundColor Red
Write-Host -Object (" @EEEEEEttttt33F      ") -NoNewline -ForeGroundColor Green
Write-Host -Object ("          ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("     ;3=*^``````'*4EEV") -NoNewline -ForegroundColor Red
Write-Host -Object (" :EEEEEEttttt33@.      ") -NoNewline -ForegroundColor Green
Write-Host -Object ("    ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("     ,.=::::it=., ") -NoNewline -ForegroundColor Cyan
Write-Host -Object ("``") -NoNewline -ForegroundColor Red
Write-Host -Object (" @EEEEEEtttz33QF       ") -NoNewline -ForegroundColor Green
Write-Host -Object ("        ") -NoNewline -ForegroundColor Red
Write-Host -Object ("   ") -NoNewline -ForegroundColor Cyan
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("    ;::::::::zt33) ") -NoNewline -ForegroundColor Cyan
Write-Host -Object ("  '4EEEtttji3P*        ") -NoNewline -ForegroundColor Green
Write-Host -Object ("        ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("   :t::::::::tt33.") -NoNewline -ForegroundColor Cyan
Write-Host -Object (":Z3z.. ") -NoNewline -ForegroundColor Yellow
Write-Host -Object (" ````") -NoNewline -ForegroundColor Green
Write-Host -Object (" ,..g.        ") -NoNewline -ForegroundColor Yellow
Write-Host -Object ("       ") -NoNewline -ForegroundColor Red
Write-Host -Object ("           ") -ForegroundColor Cyan
Write-Host -Object ("   i::::::::zt33F") -NoNewline -ForegroundColor Cyan
Write-Host -Object (" AEEEtttt::::ztF         ") -NoNewline -ForegroundColor Yellow
Write-Host -Object ("     ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("  ;:::::::::t33V") -NoNewline -ForegroundColor Cyan
Write-Host -Object (" ;EEEttttt::::t3          ") -NoNewline -ForegroundColor Yellow
Write-Host -Object ("           ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("  E::::::::zt33L") -NoNewline -ForegroundColor Cyan
Write-Host -Object (" @EEEtttt::::z3F          ") -NoNewline -ForegroundColor Yellow
Write-Host -Object ("              ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -NoNewline -ForegroundColor Cyan
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object (" {3=*^``````'*4E3)") -NoNewline -ForegroundColor Cyan
Write-Host -Object (" ;EEEtttt:::::tZ``          ") -NoNewline -ForegroundColor Yellow
Write-Host -Object ("        ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("             ``") -NoNewline -ForegroundColor Cyan
Write-Host -Object (" :EEEEtttt::::z7            ") -NoNewline -ForegroundColor Yellow
Write-Host -Object ("               ") -NoNewline -ForegroundColor Red
Write-Host -Object ("") -ForegroundColor Cyan
Write-Host -Object ("                 'VEzjt:;;z>*``           ") -ForegroundColor Yellow
Write-Host -Object ("                      ````                  ") -ForegroundColor Yellow
Write-Host -Object ("")