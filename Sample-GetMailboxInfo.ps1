<#
This little script is to showcase how to get the information returned
by 2 different command lines into one PowerShell object, in order to, later,
export the result in a CSV file.

In return, that CSV file can also be used to populate another command line to
execute custom instructions based on the data from that CSV file...
#>

#================================================================
#STEP 1 - populate a first variable with a collection of objects.
#================================================================
#In this example, we store mailboxes information with a subset of properties in the $AllMailboxes PowerShell variable.
#NOTE: selecting some properties only of an object will dramatically reduce the footprint of your PowerShell script in memory (RAM)
$AllMailboxes = Get-Mailbox -ResultSize Unlimited | select Identity,Name,AdminDisplayVersion,PrimarySMTPAddress

#=================================================================================
#STEP2 - initialize the future variable that will contain all your Custom objects.
#=================================================================================
#NOTE: this variable has to be initialized as an empty Array - the PowerShell code for empty array is @()
$AllCollection = @()

#===========================================================================================
#STEP3 - For each object in the collection, we execute another PowerShell command line that.
# will return other properties that are not available in the first PowerShell command line.
#===========================================================================================
#In this example, the second PowerShell command line is Get-MailboxStatistics.
Foreach ($Mailbox in $AllMailboxes){
    #Note the FOREACH structure : we use a "temporary variable"  named "$Mailbox"
    # that will scan each object stored in the "$AllMailboxes" variable

    #Then, we call "Get-MailboxStatistics" against each mailbox:
    $MailboxDetails = Get-MailboxStatistics $Mailbox.Identity | select TotalItemSize,ItemCount

    #And here's the key of operations: we create a temporary object that we will
    # populate with custom properties. We will store values from both Get-Mailbox and Get-MailboxStatistics
    # in these custom properties.
    $TemporaryObject = [PSCustomObject]@{
        Name = $Mailbox.Name
        ServerVersion = $MAilbox.AdminDisplayVersion
        SMTPAddress = $Mailbox.PrimarySMTPAddress
        MailboxSize = $MailboxDetails.TotalItemSize
        NumberOfItems = $MailboxDetails.ItemCount
    }

    #After having one "temporary object" with properties coming from both Get-MAilbox and Get-MailboxStatistics,
    # we add that temporary object in the big global "$AllCollection", that will eventually contain all of the
    # mailbox and mailboxstatistics information:
    $AllCollection += $TemporaryObject
}

#And all the above makes it easier to export the whole data, with information from both Get-Mailbox and Get-MailboxStatistics
# that we stored in a common PSObject, which we used to populate a collection in the form of a
# PowerShell variable (named $AllCollection) that is typed as an array, which we initialized
# with @() earlier in the script.
$AllCollection | Export-Csv -NoTypeInformation c:\temp\MyMailboxes.Csv