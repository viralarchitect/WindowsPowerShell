# Load the SpeechSynthesizer class from the System.Speech assembly
Add-Type -AssemblyName System.Speech

# Create an instance of the SpeechSynthesizer class
$synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer

# Get current date and time information
$currentDate = Get-Date
$dayOfWeek = $currentDate.DayOfWeek
$month = $currentDate.ToString("MMMM")
$day = $currentDate.Day
$hour = $currentDate.Hour

# Determine the closest quarter-hour
$minute = $currentDate.Minute
$quarterHour = switch ($minute) {
    { $_ -eq 0 } { "exactly" }
    { $_ -ge 1 -and $_ -lt 15 } { "just after" }
    { $_ -ge 15 -and $_ -lt 30 } { "a quarter past" }
    { $_ -ge 30 -and $_ -lt 45 } { "half past" }
    { $_ -ge 45 -and $_ -lt 55 } { "a quarter to" }
    { $_ -ge 55 -and $_ -lt 60 } { "almost" }
    default { "approximately" }
}

# Adjust hour if it's close to the next hour for "a quarter to" or at the tail end for "o'clock"
if ($quarterHour -eq "a quarter to" -or ($quarterHour -eq "o'clock" -and $minute -ge 53)) {
    $hour++
    if ($hour -eq 24) { $hour = 0 } # Wrap around if hour exceeds 23 (i.e., midnight)
}

# Format the hour into 12-hour format and determine AM/PM
$amPm = if ($hour -ge 12) { "PM" } else { "AM" }
$formattedHour = if ($hour -gt 12) { $hour - 12 } elseif ($hour -eq 0) { 12 } else { $hour }

# Construct the phrase
if ($quarterHour -eq "approximately" -and $minute -lt 60) {
    $phrase = "Today is $dayOfWeek, $month $day. The time is $formattedHour $amPm."
} else {
    $phrase = "Today is $dayOfWeek, $month $day. The time is $quarterHour $formattedHour $amPm."
}

# Use the Speak method to say the phrase
$synthesizer.Speak($phrase)
