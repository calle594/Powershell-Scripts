$job = Start-Job -Name "Job1" -ScriptBlock {

    # Load Outlook
    $o = New-Object -comobject outlook.application
    $n = $o.GetNamespace("MAPI")
    
    # Pick Inbox
    $Inbox = $n.GetDefaultFolder(6) #.Item("#GL SO EMEA Operations").Folders.Item("Inbox")
    
    $SaveToFolder = "H:\test\test"
    
    $Inbox.Items | ForEach-Object {
    
        # Save attachment
        $_.attachments | Where-Object { $_.filename -like "*.gz" } | ForEach-Object {
    
            $file = $_.filename
            $_.saveasfile((Join-Path $SaveToFolder $file))
        }
    }
}
    
Start-Sleep -s 5
Stop-Job $job