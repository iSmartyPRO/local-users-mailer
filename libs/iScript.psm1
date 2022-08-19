# define Error handling
# note: do not change these values
$global:ErrorActionPreference = "Stop"
if($verbose){ $global:VerbosePreference = "Continue" }


#==========================================================================
# FUNCTION iSendMail
#==========================================================================
Function iSendMail {
    <#
        .SYNOPSIS
        Send an e-mail to recipient
        .DESCRIPTION
        Send an e-mail to recipient
    #>
    [CmdletBinding()]
    Param(
        [Object]$params
    )

    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Send mail
        try {
            chcp 866 | Out-Null
            $encoding = [Console]::OutputEncoding
            [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("cp866")
            [Console]::OutputEncoding = $encoding
          # Create an SMTP Server Object
          $Smtp = New-Object System.Net.Mail.SMTPClient($params.SMTPServer,$params.SMTPPort)
          $Smtp.EnableSsl = $params.EnableSsl
          $Smtp.Credentials = New-Object System.Net.NetworkCredential($params.SMTPUser, $params.SMTPPass)

          # Create the message

          $mail = New-Object System.Net.Mail.Mailmessage
          $mail.From = $params.SMTPUser
          ForEach($recipient in $params.Recipients) {
            $mail.To.Add($recipient)
          }
          $mail.Subject = $params.Subject
          $mail.Body = $params.Text
          $mail.IsBodyHTML=$true
    $mail
          # Send send the Mail
          $result = $smtp.send($mail)
          return $result
          #}
            Exit 0
        } catch {
            Write-Host "Error description:"
            Write-Host $_
        }
    }
    end {
        Write-Host "Finished"
    }
}

#==========================================================================
# FUNCTION Get-iLocalUsers
#==========================================================================
Function Get-iLocalUsers {
    <#
        .SYNOPSIS
        Get Local Users with needed info
        .DESCRIPTION
        Get Local Users with needed info
    #>
    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Send mail
        try {
            return Get-LocalUser | ? {$_.Enabled -eq $true } | Select Name, FullName, Description, PasswordExpires, UserMayChangePassword, PasswordRequired, PasswordLastSet, LastLogon
        } catch {
            Write-Host "Error description:"
            Exit 1
        }
    }
    end {

    }
}

#==========================================================================
# FUNCTION ConvertTo-ParsedUsers
#==========================================================================
Function ConvertTo-ParsedUsers {
    <#
        .SYNOPSIS
        Send an e-mail to recipient
        .DESCRIPTION
        Send an e-mail to recipient
    #>
    [CmdletBinding()]
    Param(
        [Object]$users
    )
    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Send mail
        try {
            $parsed = @()
            ForEach($user in $users) {
                <# $user #>
                $desc = $user.Description
                if($desc -gt 0) {
                    $descSplited = $desc.Split(";")
                    if($descSplited.Length -eq 1) {
                        $department = $descSplited[0]
                    } else {
                        $position, $TabelNumber = ""
                    }
                    if($descSplited.Length -eq 2) {
                        $department = $descSplited[0]
                        $position = $descSplited[1]
                    } else {
                        $TabelNumber = ""
                    }
                    if($descSplited.Length -eq 3) {
                        $department = $descSplited[0]
                        $position = $descSplited[1]
                        $TabelNumber = $descSplited[2]
                    }
                } else {
                    $department, $position, $TabelNumber = ""
                }
                $parsed += [pscustomobject] @{
                    UserName = $user.Name
                    FullName = $user.FullName
                    Department = $department
                    Position = $position
                    TabelNumber = $TabelNumber
                }
            }
            return $parsed
        } catch {
            Write-Host $_
        }
    }
    end {

    }
}
#==========================================================================