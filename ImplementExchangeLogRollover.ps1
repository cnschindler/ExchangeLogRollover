[cmdletbinding()]

Param(
[Parameter(Mandatory=$true,Helpmessage = "Please specify the group managed service account username with a '$'-sign at the end")]
[ValidatePattern('^.{1,19}(\$)$')]
$ServiceAccountName,
$ScriptFolder = "C:\Admin\Scripts",
$RolloverTimeInDays = 30,
[Parameter(Mandatory = $false,HelpMessage = "Please use 'hh:mm:ss' format for the StartTaskAt Parameter")]
[ValidatePattern('^([0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]$')]
$StartTaskAt = "00:00:00"
)

$RolloverTaskName = "ExchangeLogRollover"
$RolloverScriptFullPath = ($ScriptFolder + "\LogRollover.cmd")

# Create LogRollover Script
$RolloverTimeInDaysWithMinus = "-" + $RolloverTimeInDays
$Rollovercontent = "forfiles /P 'c:\Inetpub\logs' /S /M *.log /D $($RolloverTimeInDaysWithMinus) /C 'cmd /c del @path'"
$Rollovercontent = $Rollovercontent + "`r`n"
$Rollovercontent = $Rollovercontent + "forfiles /P 'C:\Program Files\Microsoft\Exchange Server\V15\Logging' /S /M *.log /D $($RolloverTimeInDaysWithMinus) /C 'cmd /c del @path'"

# Which Domain does the computer belong to
$ComputerDomain = (Get-CimInstance -ClassName Win32_ComputerSystem).Domain

# Build the Pre-Windows2000 username (Domain\Username) - needed for Task Scheduler...
$ServiceAccountNameWithDomain = $ComputerDomain + "\" + $ServiceAccountName

# Check if GMSA exists and local system is allowed to use it
try
{
    $TestResult = Test-ADServiceAccount -Identity $ServiceAccountName -ErrorAction Stop
}

catch
{
    Write-Host -ForegroundColor Red -Object "Failure when accessing group managed service account credentials! Error: $_"
    Exit
}

if ($TestResult -eq $false)
{
    Exit
}

# Check if the script folder exists and create it if necessary
if (-Not (Test-Path $ScriptFolder))
{
    Write-Host -ForegroundColor Green -Object "`nCreating Folder $($ScriptFolder)"
    mkdir $ScriptFolder
}

# Write $Rollovercontent to batchfile
Write-Host -ForegroundColor Green -Object "`nWriting Rollover batchfile $($RolloverScriptFullPath)"
Set-Content -Value $Rollovercontent -Path $RolloverScriptFullPath

# Grant Permissions for Service Account in File System locations
Write-Host -ForegroundColor Green -Object "`nGranting $($ServiceAccountNameWithDomain) delete Permissions on Folder 'c:\inetpub\logs\logfiles\w3svc1'`n"
Start-Process -FilePath "C:\Windows\System32\icacls.exe" -ArgumentList "c:\inetpub\logs\logfiles\w3svc1 /Grant ${ServiceAccountNameWithDomain}:(OI)(CI)(D,RD,REA,X,RA)" -NoNewWindow -Wait
Write-Host -ForegroundColor Green -Object "`nGranting $($ServiceAccountNameWithDomain) delete Permissions on Folder 'c:\inetpub\logs\logfiles\w3svc2'`n"
Start-Process -FilePath "C:\Windows\System32\icacls.exe" -ArgumentList "c:\inetpub\logs\logfiles\w3svc2 /Grant ${ServiceAccountNameWithDomain}:(OI)(CI)(D,RD,REA,X,RA)" -NoNewWindow -Wait
Write-Host -ForegroundColor Green -Object "`nGranting $($ServiceAccountNameWithDomain) delete Permissions on Folder 'c:\Program Files\Microsoft\Exchange Server\V15\Logging' and subfolders`n"
Start-Process -FilePath "C:\Windows\System32\icacls.exe" -ArgumentList "`"c:\Program Files\Microsoft\Exchange Server\V15\Logging`" /Grant ${ServiceAccountNameWithDomain}:(OI)(CI)(D,RD,REA,X,RA)" -NoNewWindow -Wait

# Create Scheduled Task
Write-Host -ForegroundColor Green -Object "`nRegistering Scheduled Task $($RolloverTaskName)"
$a = New-ScheduledTaskAction -Execute $RolloverScriptFullPath
$p = New-ScheduledTaskPrincipal -UserId $ServiceAccountNameWithDomain -LogonType Password
$t = New-ScheduledTaskTrigger -At $StartTaskAt -Daily
$Task = New-ScheduledTask -Action $a -Principal $p -Trigger $t
Register-ScheduledTask -TaskName $RolloverTaskName -InputObject $Task

Write-Host -ForegroundColor Yellow -Object "`nScript complete. Please make sure to assign the 'Logon as a batch file' user right (SeBatchLogonRight) to your group managed service account on this machine!`n"
 
