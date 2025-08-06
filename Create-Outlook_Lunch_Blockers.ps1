# Outlook-Anwendung initialisieren und MAPI-Namespace abrufen
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Zeitraum definieren: 1. bis 23. Dezember 2025
$StartDate = Get-Date "2025-12-01"
$EndDate = Get-Date "2025-12-23"

# Startzeiten für Termine: zwischen 11:00 und 12:15 Uhr in 15-Minuten-Schritten
# (Dezimalwerte: .25 = 15 Min, .5 = 30 Min, .75 = 45 Min)
$TimeSlots = @(11, 11.25, 11.5, 11.75, 12, 12.25)

# Mögliche Dauern der Termine in Minuten: 90 oder 120 Minuten
$Durations = @(90, 120)

# Schleife über jeden Tag im definierten Zeitraum
for ($Date = $StartDate; $Date -le $EndDate; $Date = $Date.AddDays(1)) {

    # Nur an Werktagen (Montag bis Freitag)
    if ($Date.DayOfWeek -in @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")) {

        # Zufällige Auswahl eines Startzeit-Slots
        $RandomSlot = Get-Random -InputObject $TimeSlots
        $Hour = [int]$RandomSlot
        $Minutes = [int](($RandomSlot - $Hour) * 60)

        # Startzeit des Termins auf Basis des Datums + gewähltem Zeitslot berechnen
        $StartTime = Get-Date "$($Date.ToString("yyyy-MM-dd"))"
        $StartTime = $StartTime.AddHours($Hour).AddMinutes($Minutes)

        # Zufällige Auswahl einer Dauer (90 oder 120 Minuten)
        $Duration = Get-Random -InputObject $Durations

        # Outlook-Termin erstellen
        $Appointment = $Outlook.CreateItem(1) # olAppointmentItem

        # Termineigenschaften setzen
        $Appointment.Start = $StartTime                               # Startzeit
        $Appointment.End = $StartTime.AddMinutes($Duration)          # Endzeit berechnen
        $Appointment.Subject = "Blocker ($($Duration / 60) Std)"     # Betreff mit Dauer in Stunden
        $Appointment.BusyStatus = 2                                   # Status: Beschäftigt
        $Appointment.Importance = 2                                   # Priorität: Hoch
        $Appointment.Sensitivity = 2                                  # Sensitivität: Privat
        $Appointment.ReminderSet = $false                             # Keine Erinnerung
        $Appointment.Categories = "Mittag"                            # Kategorie: Mittag

        # Termin speichern
        $Appointment.Save()
    }
}
