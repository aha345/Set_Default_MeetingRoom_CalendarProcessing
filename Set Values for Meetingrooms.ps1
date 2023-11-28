# Connect to Exchange Online with the provided credentials
try {
    Get-AcceptedDomain -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
} catch {
     if {
        if (-not $Credential) {
            # Prompt the user for AzureAD credentials
            $Credential = Get-Credential -Message "Enter your AzureAD credentials"
        }

        try {
            Connect-ExchangeOnline -Credential $Credential -ErrorAction Stop | Out-Null
        } catch {
            $ConditionalAccess = $True
        }
    } else ($ConditionalAccess) {
        Connect-ExchangeOnline | Out-Null
    }
}

#-----------------------------------------------------------------------------#

# Change this to specify
$AddOrganizerToSubject = $true
$RemovePrivateProperty = $false
$DeleteSubject = $false
$DeleteComments = $true

#-----------------------------------------------------------------------------#

if ($AddOrganizerToSubject) {
    $AddOrganizerToSubjectvalue = "True"
} else {
    $AddOrganizerToSubjectvalue = "False"
}

if ($RemovePrivateProperty) {
    $RemovePrivatePropertyvalue = "True"
} else {
    $RemovePrivatePropertyvalue = "False"
}

if ($DeleteSubject) {
    $DeleteSubjectvalue = "True"
} else {
    $DeleteSubjectvalue = "False"
}

if ($DeleteComments) {
    $DeleteCommentsvalue = "True"
} else {
    $DeleteCommentsvalue = "False"
}

# Blank list for changed meeting rooms
$Changed = $null

# Set success to false
$Success = $false

while (-not $Success) {
    try {
        # Get name of mailbox
        $RoomName = $(Write-Host "What is the UPN of the room? " -NoNewLine) + $(Write-Host "(Type 'All' for all rooms or specify UPN) : " -ForegroundColor Yellow -NoNewline; Read-Host)

        if ($RoomName -ne "All") {
            # Get mailbox info
            $Rooms = Get-MailBox -Identity $RoomName -RecipientTypeDetails RoomMailbox -ErrorAction Stop

            # If found set to True and give info
            Write-Host "Room found"
            $Success = $True
        } elseif ($RoomName -eq "All") {
            # Get Room mailboxes
            $Rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox
            $Success = $True
        }
    } catch {
        Write-Warning "Room not found, try again!"
        Start-Sleep -Seconds 3
    }
}

$Changed = @()

#Loop each room in $Rooms
foreach ($Room in $Rooms) {

    $success = $false

    try {

        $CalendarName=($CalendarFolder=Get-MailboxFolderStatistics -Identity $Room.PrimarySMTPAddress -FolderScope Calendar | Where-Object {$_.FolderType -eq 'Calendar'}|Select-Object Name).Name

        Set-MailBoxFolderPermission -Identity "$($Room.PrimarySMTPAddress):\$($CalendarName)" -User "Default" -AccessRights "Reviewer"

    } catch {

        Write-Warning "Could not set permission for meetingroom"

    }

    while ($success -ne $true) {
        
        Write-Host ""
        Write-Host "Trying for $($Room.UserPrincipalName)" -ForegroundColor Green

        $RoomValue = Get-CalendarProcessing -Identity $room.UserPrincipalName

        # Output Values
        Write-Host ""
        Write-Host "What is the current values:" -ForegroundColor Cyan
        Write-Host "    AddOrganizerToSubject: $($RoomValue.AddOrganizerToSubject) (Should be $AddOrganizerToSubjectvalue)" -ForegroundColor Gray
        Write-Host "    RemovePrivateProperty: $($RoomValue.RemovePrivateProperty) (Should be $RemovePrivatePropertyvalue)" -ForegroundColor Gray
        Write-Host "    DeleteSubject: $($RoomValue.DeleteSubject) (Should be $DeleteSubjectvalue)" -ForegroundColor Gray
        Write-Host "    DeleteComments: $($RoomValue.DeleteComments) (Should be $DeleteCommentsvalue)" -ForegroundColor Gray
        Write-Host ""

        try{

            if (($RoomValue.AddOrganizerToSubject -ne $AddOrganizerToSubject) -or ($RoomValue.RemovePrivateProperty -ne $RemovePrivateProperty) -or ($RoomValue.DeleteSubject -ne $DeleteSubject) -or ($RoomValue.DeleteComments -ne $DeleteComments)) {

                try {

                    Set-CalendarProcessing -Identity $Room.UserPrincipalName -AddOrganizerToSubject $AddOrganizerToSubject -RemovePrivateProperty $RemovePrivateProperty -DeleteSubject $DeleteSubject -DeleteComments $DeleteComments -ErrorAction Stop

                    Write-Host "Values set for $($Room.UserPrincipalName)" -ForegroundColor Yellow

                    $NewRoomValue = Get-CalendarProcessing -Identity $room.UserPrincipalName
                    Write-Host "    AddOrganizerToSubject: $($RoomValue.AddOrganizerToSubject) -> $($NewRoomValue.AddOrganizerToSubject)" -ForegroundColor Gray
                    Write-Host "    RemovePrivateProperty: $($RoomValue.RemovePrivateProperty) -> $($NewRoomValue.RemovePrivateProperty)" -ForegroundColor Gray
                    Write-Host "    DeleteSubject: $($RoomValue.DeleteSubject) -> $($NewRoomValue.DeleteSubject)" -ForegroundColor Gray
                    Write-Host "    DeleteComments: $($RoomValue.DeleteComments) -> $($NewRoomValue.DeleteComments)" -ForegroundColor Gray
                    Write-Host ""            

                    $Changed += $Room.UserPrincipalName

                    $success = $true

                } catch {

                Start-Sleep -Seconds 10
            
                }

            } else {

                $success = $true

                Write-Host "No new values set"

            }

        } catch {

            Start-Sleep -Seconds 5

        }

        Start-Sleep -Seconds 3

    }

}

if ($Changed -ne $null) {
    Write-Host ""
    Write-Host "These rooms have been changed:" -ForegroundColor Yellow    
    foreach ($Change in $Changed) {
        Write-Host "    - $($Change)" -ForegroundColor Cyan
    }
}
