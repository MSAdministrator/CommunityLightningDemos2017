# PowerShell Summit 2017

# PPRT (PowerShell Phishing Resposne Toolkit) Demo

# Show new method
# Properties exposed

# Creating a new MessageObject that can be used throughout the module
$logPath = 'C:\Users\Josh\Desktop'
$MsgObject = @()
$props = @{
    'Message' = 'C:\Users\Josh\Desktop\MSG_FROM_TRIAGE\PHISHING EMAIL Trending40.com Top 40 Cyber Innovators  Entrepreneurs.msg'
    'LogPath' = $logPath
    'FullDetails' = $true
}

$MsgObject = New-MessageObject @props

# Get URL from Message object
Get-URLFromMessage -Message $MsgObject -LogPath $logPath

# Get IP address of that URL
Get-IPaddress -hostname $URL.URL[0] -LogPath $logPath




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