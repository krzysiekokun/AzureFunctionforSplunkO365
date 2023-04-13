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

# Wczytaj wszystkie adresy e-mail z pliku userIdOrEmail.txt
$allUserIdsOrEmails = Get-Content -Path "D:\events_calendar\userIdOrEmail.txt"

# Zmienna przechowująca wyniki dla wszystkich użytkowników
$allUsersResults = @()

foreach ($userIdOrEmail in $allUserIdsOrEmails) {
    # Pobierz wydarzenia dla aktualnego użytkownika
    $eventsUri = "https://graph.microsoft.com/v1.0/users/$userIdOrEmail/calendar/events"
    $eventsResponse = Invoke-RestMethod -Uri $eventsUri -Headers $headers -Method Get
    $events = $eventsResponse.value

    # Wypisz wyniki w formacie JSON
    $allResults = @()
    foreach ($event in $eventsResponse.value) {
        $eventCreatedDate = Get-Date -Date $event.createdDateTime
        if ($null -ne $event.recurrence -and $eventCreatedDate -ge $createdStartDate -and $eventCreatedDate -le $createdEndDate) {
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

    # Dodaj wyniki dla bieżącego użytkownika do zbiorczych wyników
    $allUsersResults += $allResults
}

# Sortuj wyniki według liczby uczestników (attendeesCount) w kolejności malejącej
$sortedResults = $allUsersResults | Sort-Object -Property attendees -Descending

# Zapisz wynik do pliku CSV
$csvFilePath = "results_all_users.csv"
$sortedResults | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8

Write-Host "Wyniki zapisane do pliku: $csvFilePath"