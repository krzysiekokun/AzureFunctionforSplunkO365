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
    "Accept"        = "application/json; charset=utf-8"
}

# Pobierz użytkowników z pliku userIdOrEmail.txt
$userIdsOrEmails = Get-Content -Path "userIdOrEmail.txt" -Encoding UTF8

# Dodaj przedział czasowy dla createdDateTime
$createdStartDate = Get-Date -Year 2023 -Month 3 -Day 1
$createdEndDate = Get-Date -Year 2023 -Month 6 -Day 30

# Wspólny plik CSV dla wyników wszystkich użytkowników
$combinedCsvFilePath = "combined-results.csv"

foreach ($userIdOrEmail in $userIdsOrEmails) {
    Write-Host "Pobieranie wydarzeń dla użytkownika: $userIdOrEmail"

    # Wypisz wyniki w formacie CSV
    $allResults = @()

    # Pobierz wydarzenia dla wskazanego użytkownika
    $eventsUri = "https://graph.microsoft.com/v1.0/users/$userIdOrEmail/calendar/events"
    $eventsResponse = Invoke-RestMethod -Uri $eventsUri -Headers $headers -Method Get
    $events = $eventsResponse.value

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
                @{Name='attendees'; Expression={$_.attendees.Count}},
                @{Name='eventCreatedDate'; Expression={$eventCreatedDate}},
                @{Name='createdEndDate'; Expression={$createdEndDate}}
            $allResults += $selectedValues
        }
    }

    # Sortuj wyniki według liczby uczestników (attendeesCount) w kolejności malejącej
    $sortedResults = $allResults | Sort-Object -Property attendees -Descending

    # Dodaj wyniki do wspólnego pliku CSV
    $sortedResults | Export-Csv -Path $combinedCsvFilePath -NoTypeInformation -Encoding UTF8 -Append

    Write-Host "Wyniki zapisane do pliku: $combinedCsvFilePath"

    # Opuść tę instrukcję, jeśli to jest ostatni użytkownik na liście
    if ($userIdOrEmail -ne $userIdsOrEmails[-1]) {
        Write-Host "Oczekiwanie 15 sekund przed pobieraniem kolejnego użytkownika..."
        Start-Sleep -Seconds 1
}
}