####Script by Danijel Gerbez##########
####Last Change on 03/23/2021#########
#Added $port info and $smtp.port to put everything together. So messages can be sent through non-standard smtp port

$filename = "_4740Query"
$ext = ".txt"
$timestamp = (Get-Date -Format yyy-MM-dd-hh-mm)

cmd /c wevtutil qe Security "/q:*[System [(EventID=4740)]]" /f:text /rd:true /c:1 > C:\LockedOut\$timestamp$filename$ext

$textFile = Get-Content -Path C:\LockedOut\$timestamp$filename$ext
$callingComputer = $textFile[26].Substring(22).Trim()
$result = Get-Content -Path C:\LockedOut\$timestamp$filename$ext -Raw

If ($callingComputer = $textFile[26].Substring(22).Trim() | Where-Object {$_ -like '*REG1*'})
{
Write-Host "It is region 1"
$toaddress = "ITSupport_Region1@somedomain.com"
}
elseif ($callingComputer = $textFile[26].Substring(22).Trim() | Where-Object {$_ -like '*REG2*'})
{
Write-Host "It is region 2 computer"
$toaddress = "ITSupport_Region2@somedomain.com"
}
elseif ($callingComputer = $textFile[26].Substring(22).Trim() | Where-Object {$_ -like '*REG3*'})
{
Write-Host "It is Region 3 computer"
$toaddress = "ITSupport_Region3@somedomain.com"
}
elseif ($callingComputer = $textFile[26].Substring(22).Trim() | Where-Object {$_ -like '*REG4*'})
{
Write-Host "It is Region 4 Computer"
$toaddress = "ITSupport_Region4@somedomain.com"
}
elseif ($callingComputer = $textFile[26].Substring(22).Trim() | Where-Object {$_ -like '*REG5*'})
{
Write-Host "It is Region 5 Computer"
$toaddress = "ITSupport_Region5@somedomain.com"
}
elseif ($callingComputer = $textFile[26].Substring(22).Trim() | Where-Object {$_ -like '*REG6*'})
{
Write-Host "It is Region 6 Computer"
$toaddress = "ITSupport_Region6@somedomain.com"
}
else
{
Write-Host "It is General USA Computer"
$toaddress = "ITSupport_USA@somedomain.com"
}

####################EMAIL SCRIPT ###########################################################

$fromaddress = "EventManager@somedomain.com"
#$toaddress = "ITSupport_USA@somedomain.com"
#$bccaddress = ""
#$CCaddress = ""
$Subject = "User account locked from $env:COMPUTERNAME"
$body = "$result"
$attachment = "C:\LockedOut\$timestamp$filename$ext"
$smtpserver = "SMTPRelay.somedomain.com"
$port = 25

$message = new-object System.Net.Mail.MailMessage
$message.From = $fromaddress
$message.To.Add($toaddress)
#$message.CC.Add($CCaddress)
#$message.Bcc.Add($bccaddress)
#$message.IsBodyHtml = $True
$message.Subject = $Subject
#$attach = new-object Net.Mail.Attachment($attachment)
#$message.Attachments.Add($attach)
$message.body = $body
$smtp = new-object Net.Mail.SmtpClient($smtpserver)
$smtp.Port = $port
$smtp.Send($message)

##############################Keep only last 20 records######################################
gci C:\LockedOut\*.txt -Recurse| where{-not $_.PsIsContainer}| sort CreationTime -desc| 
    select -Skip 20 | Remove-Item -Force
