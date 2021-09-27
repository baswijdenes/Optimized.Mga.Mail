# Optimized.Mga.Mail
This is a submodule for [Optimized.Mga](https://github.com/baswijdenes/Optimized.Mga).  
Optimized.Mga.Mail is dependant on Optimized.Mga

* [Send-MgaMail](#Send-MgaMail)

## Send-MgaMail
Send-MgaMail speaks for itself. 

The -From addres can only be used when you connect with an application permissions.

### Examples 
```PowerShell
Send-MgaMail -From 'John.Doe@XXXXXXXXXXX.onmicrosoft.com' -To 'Jack.Doe@contoso.com' -Subject 'Test message' -Body 'This is a test message'
```
```PowerShell
Send-MgaMail -To 'Jack.Doe@contoso.com' -Subject 'Test message' -Body 'This is a test message'
```
```PowerShell
$Object = [PSCustomObject] @{
    Name = 'Testfile.csv'
    Content = (Get-Service | ConvertTo-Csv -NoTypeInformation)
}
Send-MgaMail 'Jack.Doe@contoso.com' -Subject 'Test message' -Body 'This is a test message' -AttachmentObjects $Object
```