$config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json

$Action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-NonInteractive -NoLogo -NoProfile -File `"$($config.AppPath)\script.ps1`"" -WorkingDirectory $config.AppPath
$Trigger = New-ScheduledTaskTrigger -Daily -At '7:00 AM'
$Settings = New-ScheduledTaskSettingsSet -DontStopOnIdleEnd -RestartInterval (New-TimeSpan -Minutes 1) -RestartCount 10 -StartWhenAvailable

$SecurePassword = ConvertTo-SecureString $config.TaskPassword -AsPlainText -Force
$UserName = $config.TaskUsername
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
$Password = $Credentials.GetNetworkCredential().Password

$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
$Task | Register-ScheduledTask -TaskName $config.TaskName -User $config.TaskUsername -Password $Password