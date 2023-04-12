# Uwierzytelnienie z aplikacją zarejestrowaną
$tenantId = "9be8208b-1b85-4ac6-ac80-bbe2d6889d37"
$appId = "f3dc1fe5-d774-4ac0-960f-67138f65a42d"
$clientSecretString = "7H_8Q~xfYahdoMGiSdqNGwAKfsMA2J0Q8tgP2bRg"
$secureClientSecret = ConvertTo-SecureString -String $clientSecretString -AsPlainText -Force

# Instalacja modułu MSAL.PS
if (-not (Get-Module -ListAvailable -Name "MSAL.PS")) {
    Install-Module -Name MSAL.PS -Scope CurrentUser -Force
}

# Importowanie modułu MSAL.PS
Import-Module MSAL.PS

# Uzyskanie tokena dostępu
$authority = "https://login.microsoftonline.com/$tenantId"
$scope = "https://graph.microsoft.com/.default"
$token = Get-MsalToken -ClientId $appId -ClientSecret $secureClientSecret -Authority $authority -Scope $scope

# Użycie tokena dostępu do wywołania API Graph
$headers = @{
    "Authorization" = "Bearer $($token.AccessToken)"
    "Content-Type"  = "application/json"
}

# Pobierz listę wszystkich użytkowników
$usersUri = "https://graph.microsoft.com/v1.0/users"
$usersResponse = Invoke-RestMethod -Uri $usersUri -Headers $headers -Method Get
$users = $usersResponse.value

# Pobierz listę wszystkich użytkowników
$usersUri = "https://graph.microsoft.com/v1.0/users"
$usersResponse = Invoke-RestMethod -Uri $usersUri -Headers $headers -Method Get
$users = $usersResponse.value

# Dodaj przedział czasowy, dla którego chcesz pobrać spotkania
$startDate = "2023-04-12T00:00:00Z"
$endDate = "2023-04-12T23:59:59Z"


# Pobierz spotkania cykliczne dla każdego użytkownika w określonym przedziale czasowym
foreach ($user in $users) {
    $userId = $user.id
    $eventsUri = "https://graph.microsoft.com/v1.0/users/$userId/events?$filter=recurrence ne null and start/dateTime ge '$startDate' and end/dateTime le '$endDate'"
    $eventsResponse = Invoke-RestMethod -Uri $eventsUri -Headers $headers -Method Get

    # Wypisz wyniki
    Write-Host "User: $($user.displayName)"
    foreach ($event in $eventsResponse.value) {
        Write-Host "Id: $($event.id)"
        Write-Host "Subject: $($event.subject)"
        Write-Host "Start: $($event.start.dateTime)"
        Write-Host "End: $($event.end.dateTime)"
        Write-Host "Location: $($event.location.displayName)"
        Write-Host "Organizer: $($event.organizer.emailAddress.name) - $($event.organizer.emailAddress.address)"
        Write-Host "Attendees:"
        foreach ($attendee in $event.attendees) {
            Write-Host "  - $($attendee.emailAddress.name) - $($attendee.emailAddress.address) - $($attendee.type)"
        }
 #       Write-Host "Body: $($event.body.content)"
        Write-Host "Recurrence: $($event.recurrence)"
        Write-Host "ResponseStatus: $($event.responseStatus.response) - $($event.responseStatus.time)"
        Write-Host "Sensitivity: $($event.sensitivity)"
        Write-Host "IsCancelled: $($event.isCancelled)"
        Write-Host "IsOnlineMeeting: $($event.isOnlineMeeting)"
        Write-Host "IsAllDay: $($event.isAllDay)"
        Write-Host "ShowAs: $($event.showAs)"
        Write-Host "Categories: $($event.categories -join ', ')"
        Write-Host "ReminderMinutesBeforeStart: $($event.reminderMinutesBeforeStart)"
        Write-Host "HasAttachments: $($event.hasAttachments)"
        Write-Host "WebLink: $($event.webLink)"
        Write-Host "-------------------"
    }
    Write-Host "===================================="
}
