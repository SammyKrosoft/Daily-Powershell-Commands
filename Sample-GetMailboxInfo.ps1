#Example 1.4

$AllMailboxes = Get-Mailbox -ResultSize Unlimited | select Identity,Name,AdminDisplayVersion,PrimarySMTPAddress

# TotalMailboxSize -> Get-MailboxStatistics
$AllCollection = @()

Foreach ($Mailbox in $AllMailboxes){
    $MailboxDetails = Get-MailboxStatistics $Mailbox.Identity | select TotalItemSize,ItemCount
    $TemporaryObject = [PSCustomObject]@{
        Name = $Mailbox.Name
        ServerVersion = $MAilbox.AdminDisplayVersion
        SMTPAddress = $Mailbox.PrimarySMTPAddress
        MailboxSize = $MailboxDetails.TotalItemSize
        NumberOfItems = $MailboxDetails.ItemCount
        }
$AllCollection += $TemporaryObject
}

$AllCollection | Export-Csv -NoTypeInformation c:\temp\MyMailboxes.Csv

    