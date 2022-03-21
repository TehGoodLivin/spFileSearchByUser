$sitePath = "https://usaf.dps.mil/sites/52msg/CS/" # SITE PATH
$email = "austin.livengood@spaceforce.mil" # SEARCH FOR USER BY EMAIL
$parentSiteOnly = $false # SEARCH ONLY PARENT SITE AND IGNORE SUB SITES

Connect-PnPOnline -Url $sitePath -UseWebLogin # CONNECT TO SPO
$subSites = Get-PnPSubWeb -Recurse # GET ALL SUBSITES
$getDocLibs = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }

$reportPath = "C:\users\$env:USERNAME\Desktop\$((Get-Date).ToString("yyyyMMdd_HHmmss"))_SiteFileEditSearch.csv" # REPORT PATH (DEFAULT IS TO DESKTOP)
$results = @() # RESULTS

Write-Host "Searching: $($sitePath)" -ForegroundColor Green

# GET PARENT DOCUMENT LIBRARIES
foreach ($DocLib in $getDocLibs) { 
    $allItems = Get-PnPListItem -List $DocLib -PageSize 1000 | Where {$_[ "FileLeafRef"] -like "*.*" }

    if ($allItems -eq $null) {
    } else {
        # LOOP THROUGH EACH DOCMENT IN THE PARENT SITES
        foreach ($Item in $allItems) {
		    if ($Item -eq $null) {
                Write-Host "Error: 'Unable to pull file information'."
            } else {
                if ($Item['Author'].Email -eq $email -or $Item['Editor'].Email -eq $email) {
                    if ($Item["FileLeafRef"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                    if ($Item["File_x0020_Type"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                    if ($Item["File_x0020_Size"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                    if ($Item["FileRef"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                    if ($Item["Created"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                    if ($Item["Modified"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                    if ($Item["Author"] -eq $null){ $Item["Author"] = 'INFO NOT FOUND' }
                    if ($Item["Editor"] -eq $null){ $Item["Editor"] = 'INFO NOT FOUND' }
           
                    $results = New-Object PSObject -Property @{
                        FileName = $Item["FileLeafRef"]
                        FileExtension = $Item["File_x0020_Type"]
                        FileSize = $Item["File_x0020_Size"]
                        Path = $Item["FileRef"]
                        Created = $Item["Created"]
                        Modified = $Item["Modified"]
                        CreatedBy = $Item["Author"].Email
                        ModifiedBy = $Item["Editor"].Email
                    }

                    if (test-path $reportPath) {
                        $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Created", "Modified", "CreatedBy", "ModifiedBy" | Export-Csv -Path $reportPath -Force -NoTypeInformation -Append
                    } else {
                        $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Created", "Modified", "CreatedBy", "ModifiedBy" | Export-Csv -Path $reportPath -Force -NoTypeInformation
                    }
                }
            }
        }
    }
}

# GET ALL SUB SITE DOCUMENT LIBRARIES
if ($parentSiteOnly -eq $false) {
    foreach ($site in $subSites) {
        Connect-PnPOnline -Url $site.Url -UseWebLogin # CONNECT TO SPO SUBSITE
        $getSubDocLibs = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101}

        Write-Host "Searching: $($site.Url)" -ForegroundColor Green

        foreach ($subDocLib in $getSubDocLibs) {
            $allSubItems = Get-PnPListItem -List $subDocLib -PageSize 1000 | Where {$_["FileLeafRef"] -like "*.*"}

            if ($allSubItems -eq $null) {
            } else {
                # LOOP THROUGH EACH DOCMENT IN THE PARENT SITES
                foreach ($Item in $allSubItems) {
		            if ($Item -eq $null) {
                        Write-Host "Error: 'Unable to pull file information'."
                    } else {
                        if ($Item['Author'].Email -eq $email -or $Item['Editor'].Email -eq $email) {
                            if ($Item["FileLeafRef"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                            if ($Item["File_x0020_Type"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                            if ($Item["File_x0020_Size"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                            if ($Item["FileRef"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                            if ($Item["Created"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                            if ($Item["Modified"] -eq $null){ $Item["FileLeafRef"] = 'INFO NOT FOUND' }
                            if ($Item["Author"] -eq $null){ $Item["Author"] = 'INFO NOT FOUND' }
                            if ($Item["Editor"] -eq $null){ $Item["Editor"] = 'INFO NOT FOUND' }
           
                            $results = New-Object PSObject -Property @{
                                FileName = $Item["FileLeafRef"]
                                FileExtension = $Item["File_x0020_Type"]
                                FileSize = $Item["File_x0020_Size"]
                                Path = $Item["FileRef"]
                                Created = $Item["Created"]
                                Modified = $Item["Modified"]
                                CreatedBy = $Item["Author"].Email
                                ModifiedBy = $Item["Editor"].Email
                            }

                            if (test-path $reportPath) {
                                $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Created", "Modified", "CreatedBy", "ModifiedBy" | Export-Csv -Path $reportPath -Force -NoTypeInformation -Append
                            } else {
                                $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Created", "Modified", "CreatedBy", "ModifiedBy" | Export-Csv -Path $reportPath -Force -NoTypeInformation
                            }
                        }
                    }
                }
            }
        }
    }
}

Disconnect-PnPOnline
