<#
changelog
1.0.0 Back up your own outlook files to a folder
1.0.1 Back up outlook files for every user to a folder
1.0.2 Back up outlook files to the user's I drive
1.1.0 Added a .bat that can be run on startup
1.1.1 Edited the script to create its own .bat file on my personal drive
1.1.2 File is automatically transferred to destination computer
1.2.0 The functionality of my printerscript has been added (printer backup & restore)
1.2.1 Feedback to user has been added
1.3.0 Script is now portable and can be run on remote computers
1.4.0 Script automatically creates a script on pc-cb adapted to the computer which moves startup bat file to the new pc.
1.4.1 pc-cb functionality discontinued, it was too buggy. Remote still works.
1.5.0 startup- en restorescript are both created from the main script, backuplocation edited, bugfixes.
1.5.1 When useraccount already exists on new PC, data gets copied instantly.
1.5.2 Add option to skip restore
1.5.3 

Bugs:
When user doesn't have a backup on new pc, the script will always be run with an error
The script only runs succesfully once, then the startup file gets deleted

Future releases:
More feedback to user.
Graphic interface 
#>
cls

#Outlook backup starts here
$PCold = Read-Host "Voer oude PC in. Als $env:computername de oude pc is druk op enter"
if ($PCold -lt 1) 
    {$PCold = $env:computername}
$users = dir "\\$PCold\C$\Users\"
$restore = Read-Host "wil je al een restore uitvoeren? ja of nee"
if ($restore -eq 'ja')
    {$PCnew = Read-Host "Voer de nieuwe PC in waar de backup op hersteld moet worden"}
$startupbat = "Powershell.exe -executionpolicy remotesigned -File `"I:\Windows\backup\Backup $PCold\RESTORE.ps1`""
$startuppath = "\\$PCnew\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
$restorescript = '$path = "$env:APPDATA\Microsoft\Outlook\"
$flag = "$path\flag.txt"
$pathbackup = "I:\Windows\backup\backup*\out*"
if ( -Not (Test-Path "$path"))
    {
    New-Item "$path" -ItemType directory
    }
if ( -not (Test-Path "$flag"))
    {
    Copy-Item "$pathbackup" "$path" -Force
    New-Item $path\flag.txt -itemtype file
    }
Remove-Item -Literalpath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\start restore COM040.bat"
#Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force
'
foreach ($_ in $users)
    {
    $outlook = "\\$PCold\C$\Users\$_\Appdata\roaming\Microsoft\Outlook\"
    $pathbackup = "\\cifs01\Users\$_\Windows\backup\backup $PCold\"
    #BACK UP
    if (Test-path "$outlook")
        {
        if ( -Not (Test-Path "$pathbackup"))
            {
            New-Item "$pathbackup" -ItemType directory | Out-Null
            }
        Copy-Item "$outlook\out*" "$pathbackup"
        Out-File -FilePath "$pathbackup\RESTORE.ps1" -encoding oem -force -InputObject $restorescript
        Write-Host "`nEr is een backup gemaakt van de outlookgegevens van $_ in I:\Windows\backup\backup $PCold\"
        #RESTORE
        if ($restore -eq 'ja')
            {
            if (Test-Connection -ComputerName $PCnew -count 1)
                {
                if (Test-Path "\\$PCnew\C$\Users\$_\")
                    {
                    Copy-Item "$pathbackup\out*" "$outlook"
                    Write-Host "De outlookgegevens zijn naar $PCnew gekopieerd!"
                    }
                else
                    { 
                    Out-File -FilePath "$startuppath\start restore $PCold.bat" -Encoding oem -Force -InputObject $startupbat
                    Write-Host "account bestaat nog niet, de eerste keer dat de gebruiker inlogt op $pcnew krijgt hij een outlook restore"
                    }
                }
            else
                {
                Write-Host "$PCnew is niet beschikbaar, voer restore handmatig uit"
                }
            }
        }
    }
#Printerscript     <WEKRT ALLEEN OP PC WAAR SCRIPT OP WORDT UIGEVOERD>
$PCname = (Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name
if ($PCname -eq $PCold)
    {
    $restorepath = "I:\windows\backup\backup $PCold\RESTORE.ps1"
    Write-Host "`ndeze printers zijn meegenomen in de backup:"
    $Printarray = Get-ChildItem  -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections\' | Get-ItemProperty

    foreach ($x in $Printarray) 
        {
        $X.Printer
        $Printer = $X.Printer
        Out-File -FilePath $restorepath -Append -Encoding ascii -InputObject "rundll32 printui.dll ,PrintUIEntry /ga /n $Printer" -NoClobber
        }
    }
Else {Write-host "Printer backup werkt alleen op lokale pc"}
Read-Host Press enter to exit
#\\PC\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup
