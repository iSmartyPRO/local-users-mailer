$config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json

Import-Module ".\libs\iScript.psm1"

$users = Get-iLocalUsers
$parsedUsers = ConvertTo-ParsedUsers -users $users

$html = "<h2>$($config.EmailSubject)</h2>"
$html += "<p><i>Note: task is running everyday at 7:00AM</i></p>"
$html += "<table cellpadding=`"2`" cellspacing=`"0`" border=`"1`" style=`"width: 100%`" style=`"border-collapse: collapse; border: 2px solid #000000;`">"
$html += "<tr bgcolor=`"#000000`" style=`"color: #FFFFFF; height: 35px; border-collapse: collapse; border: 2px solid #000000;`"><td style=`"text-align: center; font-weight: 700;`">Username</td><td style=`"text-align: center; font-weight: 700;`">Full Name</td><td style=`"text-align: center; font-weight: 700;`">Tabel Number</td><td style=`"text-align: center; font-weight: 700;`">Department</td><td style=`"text-align: center; font-weight: 700;`">Position</td></tr>"
ForEach($u in $parsedUsers) {
  $html += "<tr><td>$($u.UserName)</td><td>$($u.FullName)</td><td>$($u.TabelNumber)</td><td>$($u.Department)</td><td>$($u.Position)</td></tr>"
}
$html += "</table>"

$params = [PSCustomObject]@{
    Sender      = "$($config.EmailSenderName) <$($config.EmailUser)>"
    Recipients   = $config.Recipients
    SMTPServer  = $config.EmailServer
    SMTPPort    = $config.EmailServerPort
    SMTPUser    = $config.EmailUser
    SMTPPass    = $config.EmailPass
    Subject     = $config.EmailSubject
    EnableSsl   = $config.EnableSsl
    Text        = $html
  }
iSendMail $params

Remove-Module iScript

