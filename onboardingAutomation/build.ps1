$exclude = @("venv", "onboardingAutomation.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "onboardingAutomation.zip" -Force