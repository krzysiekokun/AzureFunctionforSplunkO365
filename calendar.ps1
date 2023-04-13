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

# Dodaj przedział czasowy dla createdDateTime
$createdStartDate = Get-Date -Year 2023 -Month 3 -Day 1
$createdEndDate = Get-Date -Year 2023 -Month 4 -Day 30

# Wypisz wyniki w formacie JSON
$allResults = @()

foreach ($user in $users) {
    $userEmail = $user.mail
    $eventsUri = "https://graph.microsoft.com/v1.0/users/$userEmail/calendar/calendarView?startDateTime=$($createdStartDate.ToString("yyyy-MM-ddTHH:mm:ssZ"))&endDateTime=$($createdEndDate.ToString("yyyy-MM-ddTHH:mm:ssZ"))"
    $eventsResponse = Invoke-RestMethod -Uri $eventsUri -Headers $headers -Method Get

    foreach ($event in $eventsResponse.value) {
        $selectedValues = $event | Select-Object -Property @{
                Name = 'organizer'
                Expression = {$_.organizer.emailAddress.name}
            },
            subject,
            @{Name='recurrence'; Expression={$_.recurrence.pattern.type}},
            @{Name='start'; Expression={$_.start.dateTime}},
            @{Name='end'; Expression={$_.end.dateTime}},
            createdDateTime, lastModifiedDateTime,
            @{Name='attendees'; Expression={$_.attendees.Count}}
        $allResults += $selectedValues
    }
}

# Sortuj wyniki według liczby uczestników (attendeesCount) w kolejności malejącej
$sortedResults = $allResults | Sort-Object -Property attendees -Descending

# Zapisz wynik do pliku JSON
$jsonResults = $sortedResults | ConvertTo-Json
$jsonFilePath = "results.json"
Set-Content -Path $jsonFilePath -Value $jsonResults

Write-Host "Wyniki zapisane do pliku: $jsonFilePath"
