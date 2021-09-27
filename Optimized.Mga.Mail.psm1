#region functions
function Send-MgaMail {
    <#
    .LINK
    https://github.com/baswijdenes/Optimized.Mga.Mail/tree/main

    .SYNOPSIS
    Sends emails with Microsoft Graph.
    
    .DESCRIPTION
    Send-MgaMail uses the Microsoft Graph v1.0 REST API.
    You need Mail.Send permissions.

    .PARAMETER To
    To accepts an array of addresses.
    
    .PARAMETER Subject
    Is the email subject.
    
    .PARAMETER Body
    Is the email body.
    
    .PARAMETER From
    Add the From address when logged in with Application permissions.
    When logged in with user credentials the From address will automatically be the userLogon. This is also displayed in a warning message.
    
    .PARAMETER Attachments
    Attachments accepts an Array. Make sure to use the FullName (Including Path).
    Example:
    'C:\Temp\Attachment.txt','C:\Temp\Attachment2.txt'

    .EXAMPLE
    Send-MgaMail -From 'John.Doe@XXXXXXXXXXX.onmicrosoft.com' -To 'Jack.Doe@contoso.com' -Subject 'Test message' -Body 'This is a test message'

    .EXAMPLE
    Send-MgaMail -To 'Jack.Doe@contoso.com' -Subject 'Test message' -Body 'This is a test message'
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript( { $_ -like "*@*" })]
        [string[]]
        $To,
        [Parameter(Mandatory = $true)]
        [string]
        $Subject,
        [Parameter(Mandatory = $true)]
        [string]
        $Body,
        [Parameter(Mandatory = $false)]
        [ValidateScript( { $_ -like "*@*" })]
        [string]
        $From,
        [Parameter(Mandatory = $false)]
        [object]
        [Alias('Attachment')]
        $AttachmentPaths,
        [Parameter(Mandatory = $false)]
        [object]
        $AttachmentObjects
    )
    begin {
        try {
            $URL = 'https://graph.microsoft.com/v1.0/me/sendMail'
            $ToList = [System.Collections.Generic.List[System.Object]]::new()
            foreach ($Address in $To) {
                $Object = [PSCustomObject]@{
                    emailAddress = [PSCustomObject] @{
                        'address' = $Address
                    }
                }
                $ToList.Add($Object)
            }
            $Message = [PSCustomObject] @{
                message = [PSCustomObject] @{
                    subject      = $subject
                    body         = [PSCustomObject] @{
                        contentType = 'HTML'
                        content     = $body
                    }
                    ToRecipients = @($ToList)
                }
            }
            if (($AttachmentPaths) -or ($AttachmentObjects)) {
                Write-Verbose "Send-MgaMail: process: Attachment parameter found."
                $AttachmentsList = [System.Collections.Generic.List[System.Object]]::new()
                if ($AttachmentPaths) {
                    foreach ($Attachment in $AttachmentPaths) {
                        try {
                            Write-Verbose "Send-MgaMail: process: Testing path to $Attachment."
                            $FileBytes = Get-Content -Path $Attachment -Encoding Byte -ErrorAction stop
                            $AttachmentName = $Attachment.split('\') | Select-Object -Last 1
                            $Base64String = ([System.Convert]::ToBase64String($FileBytes))
                            $AttachmentsNode = [PSCustomObject]@{
                                "@odata.type"  = "#microsoft.graph.fileAttachment"
                                "name"         = $AttachmentName
                                "contentBytes" = $Base64String
                            }
                            $AttachmentsList.Add($AttachmentsNode)
                        }
                        catch {
                            Write-Warning "There is an error with $AttachmentPath. We will continue despite error."
                            continue
                        }
                    }
                }
                elseif ($AttachmentObjects) {
                    foreach ($AttachmentObject in $AttachmentObjects) {
                        try {
                            Write-Verbose "Send-MgaMail: process: Converting object to Base64String."
                            $Bytes = [System.Text.Encoding]::Unicode.GetBytes($AttachmentObject.Content)
                            $Base64String = [System.Convert]::ToBase64String($Bytes)
                            $AttachmentsNode = [PSCustomObject]@{
                                "@odata.type"  = "#microsoft.graph.fileAttachment"
                                "name"         = $AttachmentObject.Name
                                "contentBytes" = $Base64String
                            }
                            $AttachmentsList.Add($AttachmentsNode)
                        }
                        catch {
                            Write-Warning "There is an error with $AttachmentObject. We will continue despite error."
                            continue
                        }
                    }
                }
                $Message = [PSCustomObject] @{
                    message = [PSCustomObject] @{
                        subject      = $subject
                        body         = [PSCustomObject] @{
                            contentType = 'HTML'
                            content     = $body
                        }
                        toRecipients = @($ToList)
                        Attachments  = @($AttachmentsList)
                    }
                }
            }
            if ($From.length -gt 0) {
                if (($global:MgaRU.result.length -ge 1) -and ($global:MgaRU.result.account.Username -ne $From)) {
                    Write-Warning "You have logged in with Credentials. We cannot use a different Email Address than $From. We will use $From to send email from." 
                    $From = $global:MgaRU.result.account.Username
                }
                if (($global:MgaBasic.access_token.length -ge 1) -and ($global:MgaUserCredentials.UserName -ne $From)) {
                    Write-Warning "You have logged in with Credentials. We cannot use a different Email Address than $From. We will use $From to send email from." 
                    $From = $global:MgaUserCredentials.UserName 
                }
                Write-Verbose "Send-MgaMail: From address is $From."
                $FromNode = [PSCustomObject] @{
                    emailAddress = [PSCustomObject] @{
                        'address' = $From
                    }
                }
                $Message | Add-Member -MemberType NoteProperty -Name 'From' -Value $FromNode
                $URL = "https://graph.microsoft.com/v1.0/users/$($From)/sendMail"
            }
        }
        catch {
            throw $_.Exception.Message
        }
    }
    process {
        try {
            Write-Verbose 'Send-MgaMail: Sending email...'
            Post-Mga -URL $URL -InputObject $Message
        }
        catch {
            throw $_.Exception.Message
        }
    }
    end {
        return "Email to $To with subject $Subject has been sent succesfully."
    }
}
#endregion