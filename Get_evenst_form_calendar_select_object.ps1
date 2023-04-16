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
$createdStartDate = Get-Date -Year 2023 -Month 4 -Day 15
$createdEndDate = Get-Date -Year 2023 -Month 4 -Day 16

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

# Uwzględnij zmianę czasu na letni (czas UTC) w 2023 roku
$daylightSavingDate = Get-Date -Year 2023 -Month 3 -Day 26 -Hour 1 -Minute 0 -Second 0

foreach ($event in $eventsResponse.value) {
    $eventCreatedDate = Get-Date -Date $event.createdDateTime
    
    # Dodaj różnicę czasu dla czasu polskiego (CET/CEST)
    if ($eventCreatedDate -lt $daylightSavingDate) {
        $eventCreatedDate = $eventCreatedDate.AddHours(1)
    } else {
        $eventCreatedDate = $eventCreatedDate.AddHours(2)
    }
    
    if ($null -ne $event.recurrence -and $eventCreatedDate -ge $createdStartDate -and $eventCreatedDate -le $createdEndDate) {
        $selectedValues = $event | Select-Object -Property @{
                Name = 'Organizator spotkania'
                Expression = {$_.organizer.emailAddress.name}
            },
            subject,
            @{Name='recurrence'; Expression={$_.recurrence.pattern.type}},
            @{Name='Poczatek spotkania'; Expression={(Get-Date -Date $_.start.dateTime).AddHours(($_.start.dateTime -lt $daylightSavingDate) ? 1 : 2)}},
            @{Name='Koniec spotkania'; Expression={(Get-Date -Date $_.end.dateTime).AddHours(($_.end.dateTime -lt $daylightSavingDate) ? 1 : 2)}},
            @{Name='Anulowane?'; Expression={$_.isCancelled}},
            @{Name='Data utworzenia spotkania'; Expression={$eventCreatedDate}},
            @{Name='Data ostatniej modyfikacji spotkania'; Expression={(Get-Date -Date $_.lastModifiedDateTime).AddHours(($_.lastModifiedDateTime -lt $daylightSavingDate) ? 1 : 2)}},
            @{Name='Od kiedy spotkanie'; Expression={$_.recurrence.range.startDate}},
            @{Name='Do kiedy jest spotkanie zaplanowane'; Expression={$_.recurrence.range.endDate}},
            @{Name='Uczestnicy spotkania'; Expression={$_.attendees.Count}}
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
    Write-Host "Oczekiwanie 1 sekund przed pobieraniem kolejnego użytkownika..."
    Start-Sleep -Seconds 1
}