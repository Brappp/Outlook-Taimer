# Create an Outlook application object
$outlook = New-Object -ComObject Outlook.Application

# Get the MAPI namespace
$namespace = $outlook.GetNamespace("MAPI")

# Get all the accounts in the profile
$accounts = $namespace.Accounts

# Initialize an empty array to store the mailbox names
$mailboxNames = @()

# Loop through the accounts and add the display names to the mailbox names array
foreach ($account in $accounts) {
    $mailboxNames += $account.DisplayName
}

# Create an empty array to store the mailbox list
$mailboxList = @()

# Loop through each mailbox name in the $mailboxNames array
foreach ($mailboxName in $mailboxNames) {

    # Compute the index of the current mailbox name in the $mailboxNames array
    $index = [array]::IndexOf($mailboxNames, $mailboxName)

    # Increment the index by 1 to start the numbered list at 1 instead of 0
    $index += 1

    # Format the mailbox name and index as a numbered list item
    $numberedMailboxName = "{0}. {1}" -f $index, $mailboxName

    # Add the numbered mailbox name to the $mailboxList array
    $mailboxList += $numberedMailboxName
}

# Prompt the user to select a mailbox from the list of mailbox names
$selectedMailboxIndex = Read-Host "Select a mailbox:`n$($mailboxList -join "`n")"
if ($selectedMailboxIndex -lt 1 -or $selectedMailboxIndex -gt $mailboxNames.Count) {
    Write-Host "Invalid selection. Exiting script."
    return
}
$selectedMailboxIndex = $selectedMailboxIndex - 1

# Get the selected mailbox
$selectedMailboxName = $mailboxNames[$selectedMailboxIndex]
$selectedMailbox = $accounts | Where-Object { $_.DisplayName -eq $selectedMailboxName }

# Ask the user for the email count to check for
$countToMove = Read-Host "Enter the minimum number of emails to move"

# Get the Inbox folder of the selected mailbox
$inbox = $selectedMailbox.DeliveryStore.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Get all the emails in the Inbox folder
$emails = $inbox.Items

# Get the total number of emails
$totalCount = $emails.Count

# Initialize variables to store the number of emails to move and a list of senders to move
$countToMove = [int]$countToMove
$sendersToMove = @()

# Loop through the emails and count the number of emails to move and add the senders to the list
foreach ($email in $emails) {
    if ($email.SenderEmailAddress -ne $null) {
        $sender = $email.SenderEmailAddress
        $emailCount = ($emails | Where-Object { $_.SenderEmailAddress -eq $sender }).Count
        if ($emailCount -ge $countToMove -and $sendersToMove -notcontains $sender) {
            $sendersToMove += $sender
        }
    }
}

# If there are senders to move, show a summary and prompt the user to confirm whether to move the emails
if ($sendersToMove.Count -gt 0) {
    Write-Host "The following senders have $countToMove or more emails and will be moved to a new folder:"
    foreach ($sender in $sendersToMove) {
        Write-Host "- $sender"
    }

    $moveEmails = Read-Host "Do you want to move the emails to the new folders? (Y/N)"
    if ($moveEmails -eq "Y") {
        # Initialize an empty array to store the results
        $results = @()

        # Loop through the sender groups and move the emails to a new folder
        foreach ($sender in $sendersToMove) {
            # Check if the new folder with the sender name already exists
            $newFolder = $inbox
            $newFolder = $inbox.Folders | Where-Object { $_.Name -eq $sender }
            if (!$newFolder) {
                # Create a new folder with the sender name
                $newFolder = $inbox.Folders.Add($sender)
            }

            # Loop through the sender's emails and move them to the new folder
            $senderGroup = $emails | Where-Object { $_.SenderEmailAddress -eq $sender }
            foreach ($email in $senderGroup) {
                $email.Move($newFolder)
            }

            # Add the email sender and count to the results array
            $results += [PSCustomObject] @{
                Folder = $inbox.FolderPath
                Sender = $sender
                Count = $senderGroup.Count
            }
        }

        Write-Host "Emails have been moved to the new folders."
    } else {
        Write-Host "Emails have not been moved."
    }
} else {
    Write-Host "No senders have $countToMove or more emails to be moved."
}

# Export the results to a CSV file
$results | Export-Csv -Path "D:\Output.csv" -NoTypeInformation
