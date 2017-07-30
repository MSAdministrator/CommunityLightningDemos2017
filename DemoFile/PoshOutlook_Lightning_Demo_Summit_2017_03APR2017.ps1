# PowerShell Summit 2017

# PSOutlook Demo

# Class

. C:\_Github\PoshOutlook\PoshOutlook\Classes\OutlookClass.ps1

# Show new method
# Properties exposed

[outlook]::new()

# Create new Outlook COM Object

[Microsoft.Office.Interop.Outlook.Application] $Outlook = New-Object -ComObject Outlook.Application

$PSOutlook = [Outlook]::new($Outlook)

$PSOutlook

# Accounts Property
$PSOutlook.Accounts

# Attached Mailbox
$PSOutlook.AttachedMailbox

# Cached Mode on?
$PSOutlook.CachedMode

# Get Outlook Add-Ins Installed
$PSOutlook.ComAddins

# Get Current User information
$PSOutlook.CurrentUser

# Get your Global Address List
$PSOutlook.GlobalAddressList

# Get your outlook folders
$PSOutlook.OutlookFolders


#Other Properties
$PSOutlook.ConversationView
$PSOutlook.ExchangeConnectionMode
$PSOutlook.ExchangeVersion
$PSOutlook.FilePath
$PSOutlook.Instance
$PSOutlook.InstantSearchEnabled
$PSOutlook.OfflineStatus
$PSOutlook.OutlookFolders
$PSOutlook.OutlookVersion
$PSOutlook.ProductCode
$PSOutlook.ProfileName



# PoshOutlook Module

# Create a new Outlook COM Object
$Outlook = Create-OutlookObject -Debug

# Get Outlook Accounts
Get-OutlookAddIns -Outlook $Outlook

# Get Attached Mailbox

Get-OutlookAttachedMailbox -Outlook $Outlook

# Get Outlook Calendar Item
#Get-OutlookCalendarItem -Outlook $Outlook -start 04/01/2017 -Verbose

# Get Outlook Folders
Get-OutlookFolders -Outlook $Outlook
# Note, I'm working on a way to view Outlook folders in their native layer

# Get Outlook Global Address List
$GAL = Get-OutlookGAL -Outlook $Outlook
$GAL.count

# Get Outlook Root Folders
Get-OutlookRootFolders -Outlook $Outlook

# Get Outlook Folder Heirachy (WIP) 
# If anyone wants to help out, hit me up!

. C:\_Github\PoshOutlook\_DEV\Get-OutlookFolderHeirachy.ps1
. C:\_Github\PoshOutlook\_DEV\Get-OutlookFolderStructure.ps1

# Search Outlook attachments based on keyword or phrase
Search-OutlookAttachments -Outlook $Outlook -Item docx

Search-OutlookAttachments -Outlook $Outlook -Item pdf -Folder 'Deleted Items'

# still a lot to do
# Eventually, this will ported to a DSL (Domain Specific Language) to enhance usability
# the Outlook COM Object is tricky and has many many layers (like a wrotten onion)
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