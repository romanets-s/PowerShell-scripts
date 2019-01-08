Param(
     [string]$mailTo,
     [string]$subject,
     [string]$mailBody,
     [string]$attachment
    )
try {
    #create and send mail
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.to = $mailTo
    $mail.Subject = $subject
    $mail.HTMLBody = $mailBody
    if ($attachment -ne ""){
        $mail.attachments.add($attachment)
    }
    $mail.Send()
    $outlook.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
} catch {
       return $False
}
return $True

