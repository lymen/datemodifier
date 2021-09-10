<#
How To:
Copy this file to the folder where the doc files you want to modify are located. Right-click then select "Run with PowerShell"
#>

echo "Syncing document 'Last Modified Date' to 'Date Last Saved' . . ."
$includeExtensions = @(".doc", ".docx") 
$path = "."
$docs = Get-ChildItem -Recurse -Path $path | ?{$includeExtensions -contains $_.Extension}
$fileCount = 0

foreach($doc in $docs) {
    $application = New-Object -ComObject word.application
    $application.Visible = $false
    $document = $application.documents.open($doc.Fullname, $false, $true)
    $binding = "System.Reflection.BindingFlags" -as [type]
    $properties = $document.BuiltInDocumentProperties

    $lastsavetime = $null
    $creationdate = $null

    foreach($property in $properties)
    {
     $pn = [System.__ComObject].invokemember("name",$binding::GetProperty,$null,$property,$null)
      trap [system.exception]
       {
        continue
       }
       if($pn -eq "Last save time") {
            $lastsavetime = [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$property,$null)
       } elseif ($pn -eq "Creation date") {
            $creationdate = [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$property,$null)
       }                
    }

    $document.Close($false)
    $application.quit()

    "`nSetting " + $doc.FullName
    Set-ItemProperty $doc.FullName -Name "Creationtime" -Value $creationdate
    echo "Last save time: " $lastsavetime
    Set-ItemProperty $doc.FullName -Name "LastWriteTime" -Value $lastsavetime
    echo "Creation date: " $creationdate
    $fileCount++
}
"`n`nTotal Documents Synced: " + $fileCount
Read-Host -Prompt "Press Enter to exit"