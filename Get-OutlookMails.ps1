# Get-OutlookMails.ps1
param(
    [string]$OutputFile = "C:\MailFetcher\Mails.txt",
    [int]$CheckInterval = 10
)

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

Write-Host "📧 Mail fetcher started. Saving to $OutputFile"

while ($true) {
    $messages = $Inbox.Items | Sort-Object ReceivedTime -Descending | Select-Object -First 5
    $output = ""

    foreach ($msg in $messages) {
        $output += "[{0}] From: {1}`r`nSubject: {2}`r`nBody: {3}`r`n`r`n" -f `
            (Get-Date $msg.ReceivedTime -Format "yyyy-MM-dd HH:mm"),
            $msg.SenderName,
            $msg.Subject,
            ($msg.Body -replace "`r`n"," ")
    }

    $output | Out-File $OutputFile -Encoding UTF8
    Start-Sleep -Seconds $CheckInterval
}
