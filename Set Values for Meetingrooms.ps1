# Check if AzureAD credentials are already stored
if (-not $Credential) {
    # Prompt the user for AzureAD credentials
    $Credential = Get-Credential -Message "Enter your AzureAD credentials"
}

# Connect to Exchange Online with the provided credentials
Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false

# Blank list for changed meeting rooms
$Changed = $null

# Get Room mailboxes
$Rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox

#Loop each room in $Rooms
foreach ($Room in $Rooms) {

    $success = $false

    while ($success -ne $true) {
        
        Write-Host ""
        Write-Host "Trying for $($Room.UserPrincipalName)" -ForegroundColor Green

        $RoomValue = Get-CalendarProcessing -Identity $room.UserPrincipalName

        # Output Values
        Write-Host ""
        Write-Host "What is the current values:" -ForegroundColor Cyan
        Write-Host "    AddOrganizerToSubject: $($RoomValue.AddOrganizerToSubject) (Should be 'True')" -ForegroundColor Gray
        Write-Host "    RemovePrivateProperty: $($RoomValue.RemovePrivateProperty) (Should be 'True')" -ForegroundColor Gray
        Write-Host "    DeleteSubject: $($RoomValue.DeleteSubject) (Should be 'False')" -ForegroundColor Gray
        Write-Host "    DeleteComments: $($RoomValue.DeleteComments) (Should be 'True')" -ForegroundColor Gray
        Write-Host ""

        try{

            if (($RoomValue.AddOrganizerToSubject -ne $true) -or ($RoomValue.RemovePrivateProperty -ne $true) -or ($RoomValue.DeleteSubject -ne $false) -or ($RoomValue.DeleteComments -ne $True)) {

                try {

                    Set-CalendarProcessing -Identity $Room.UserPrincipalName -AddOrganizerToSubject $true -RemovePrivateProperty $true -DeleteSubject $false -DeleteComments $true -ErrorAction Stop

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
