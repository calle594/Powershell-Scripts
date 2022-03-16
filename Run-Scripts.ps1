#run script one

Write-Host "running script 1"
 
& "C:\Users\lawsonc\Desktop\powershell\Extract-Email.ps1"

Write-Host "running script 2"

& "C:\Users\lawsonc\Desktop\powershell\Unzip.ps1"

Write-Host "running script 3"

& "C:\Users\lawsonc\Desktop\powershell\Format-csv.ps1"

Remove-item -Path "H:\test\test\SpecialCompanies1.csv"
Remove-item -Path "H:\test\test\*.gz"

EXIT