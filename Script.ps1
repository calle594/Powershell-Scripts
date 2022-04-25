Add-Type -assembly "Microsoft.Office.Interop.Outlook"
add-type -assembly "System.Runtime.Interopservices"

try {
    $outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
    $outlookWasAlreadyRunning = $true
} catch {
    try {
        $Outlook = New-Object -comobject Outlook.Application
        $outlookWasAlreadyRunning = $false
    } catch {
        write-host "You must exit Outlook first."
        exit
    }
}

# Setup hooking ourselves onto the user's Outlook process + Set the variables for the Inbox folder, as well as the Subject of the report we want.
# Start background job
$job = Start-Job -Name "Job1" -ScriptBlock {

    $o = New-Object -comobject outlook.application
    $n = $o.GetNamespace("MAPI")
    $Inbox = $n.GetDefaultFolder(6)
    $filePath = "C:\Users\lawsonc\Desktop\Powershell\Test\" 

    $Inbox.Items | ForEach-Object {
        $_.attachments | Where-Object { $_.filename -like "*.gz" } | ForEach-Object {
            $fileName = $_.filename
            $_.saveasfile((Join-Path $filePath $fileName)) 
        }
    }
}
    
Start-Sleep -s 5
Stop-Job $job    

try {
    #WOW! Operations are letting Automation do magic, it's still in the inbox! Let's process that!
    Write-Host "EQ+ Report Found!"
    Function Format-File {
        Param(
            $infile,
            $outfile = ($folder -replace '\.gz$', ''),
            $path = $env:USERPROFILE
        )
        # Provides stream for file read and write operations. 
        $input = New-Object System.IO.FileStream $inFile, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::Read)
        $output = New-Object System.IO.FileStream $outFile, ([IO.FileMode]::Create), ([IO.FileAccess]::Write), ([IO.FileShare]::None)
        $gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)
        $buffer = New-Object byte[](1024)
        while ($true) {
            $read = $gzipstream.Read($buffer, 0, 1024)
            if ($read -le 0) { break }
            $output.Write($buffer, 0, $read)
        }
        $gzipStream.Close()
        $output.Close()
        $input.Close()
    }

    #Let's extract it to get the CSV
    $infile = Get-ChildItem -Path "C:\Users\lawsonc\Desktop\Powershell\Test\*.gz" 
    $outfile = "C:\Users\lawsonc\Desktop\Powershell\Test\SpecialCompanies_Unformatted.csv"
    Format-File $infile $outfile

    # Now let's format the CSV with Text-To-Columns
    $FormattedCsv = "C:\Users\lawsonc\Desktop\Powershell\Test\SpecialCompanies.csv"
    import-csv $outfile -Delimiter ";" -Header A,B,C,D,E,F,G | Export-Csv $FormattedCsv | Format-Table

    if (Test-Path  "W:\CS Operations\Equatex\Job Report\SpecialCompanies.csv"){
        Remove-item "W:\CS Operations\Equatex\Job Report\SpecialCompanies.csv"
        Move-Item -Path $FormattedCsv -Destination "W:\CS Operations\Equatex\Job Report"
    } else {
        Move-Item -Path $FormattedCsv -Destination "W:\CS Operations\Equatex\Job Report"
    }
}

catch {
    # # Looks like someone has deleted the EQ Report (likely), tell Operations it's failed so they can do some work...
    # Send-MailMessage -Subject "EQ+ Automated Process Failure" -Body "The automated EQ+ Report process has failed, please refer to manual procedure. 
    # `r https://computershare.service-now.com/kb_view.do?sys_kb_id=1aab1677db69fc584d9c496d139619ab&sysparm_rank=3&sysparm_tsqueryId=9fac7f43dbbe4d942de1af2913961927" `
    # -From "operations@computershare.co.uk" -To "operations@computershare.co.uk" -smtpserver webmail.emea.cshare.net

}