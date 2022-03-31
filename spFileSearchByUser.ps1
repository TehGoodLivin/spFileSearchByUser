#    DATE: 22 Mar 2022
#    UPDATED: 21 Mar 2022
#    
#    MIT License
#    Copyright (c) 2021 Austin Livengood
#    Permission is hereby granted, free of charge, to any person obtaining a copy
#    of this software and associated documentation files (the "Software"), to deal
#    in the Software without restriction, including without limitation the rights
#    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#    copies of the Software, and to permit persons to whom the Software is
#    furnished to do so, subject to the following conditions:
#    The above copyright notice and this permission notice shall be included in all
#    copies or substantial portions of the Software.
#    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#    SOFTWARE.
#
#    CHANGABLE VARIABLES
$sitePath = "https://usaf.dps.mil/sites/52msg/CS/SCX/SCXK/BRM/" # SITE PATH
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

                    $permissions = @()
                    $perm = Get-PnPProperty -ClientObject $Item -Property RoleAssignments       
                    foreach ($role in $Item.RoleAssignments) {
                        $loginName = Get-PnPProperty -ClientObject $role.Member -Property LoginName
                        $rolebindings = Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings
                        $permissions += "$($loginName) - $($rolebindings.Name)"
                    }
                    $permissions = $permissions | Out-String
           
                    $results = New-Object PSObject -Property @{
                        FileName = $Item["FileLeafRef"]
                        FileExtension = $Item["File_x0020_Type"]
                        FileSize = $Item["File_x0020_Size"]
                        Path = $Item["FileRef"]
                        Permissions = $permissions
                        Created = $Item["Created"]
                        Modified = $Item["Modified"]
                        CreatedBy = $Item["Author"].Email
                        ModifiedBy = $Item["Editor"].Email
                    }

                    if (test-path $reportPath) {
                        $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Created", "Modified", "CreatedBy", "ModifiedBy" | Export-Csv -Path $reportPath -Force -NoTypeInformation -Append
                    } else {
                        $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Created", "Modified", "CreatedBy", "ModifiedBy" | Export-Csv -Path $reportPath -Force -NoTypeInformation
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

                            $permissions = @()
                            $perm = Get-PnPProperty -ClientObject $subItem -Property RoleAssignments       
                            foreach ($role in $Item.RoleAssignments) {
                                $loginName = Get-PnPProperty -ClientObject $role.Member -Property LoginName
                                $rolebindings = Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings
                                $permissions += "$($loginName) - $($rolebindings.Name)"
                            }
                            $permissions = $permissions | Out-String
           
                            $results = New-Object PSObject -Property @{
                                FileName = $Item["FileLeafRef"]
                                FileExtension = $Item["File_x0020_Type"]
                                FileSize = $Item["File_x0020_Size"]
                                Path = $Item["FileRef"]
                                Permissions = $permissions
                                Created = $Item["Created"]
                                Modified = $Item["Modified"]
                                CreatedBy = $Item["Author"].Email
                                ModifiedBy = $Item["Editor"].Email
                            }

                            if (test-path $reportPath) {
                                $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Created", "Modified", "CreatedBy", "ModifiedBy" | Export-Csv -Path $reportPath -Force -NoTypeInformation -Append
                            } else {
                                $results | Select-Object "FileName", "FileExtension", "FileSize", "Path", "Permissions", "Created", "Modified", "CreatedBy", "ModifiedBy" | Export-Csv -Path $reportPath -Force -NoTypeInformation
                            }
                        }
                    }
                }
            }
        }
    }
}
Disconnect-PnPOnline

Write-Host "`nScript Completed: " -ForegroundColor DarkYellow -nonewline; Write-Host "$(get-date -format yyyy/MM/dd-HH:mm:ss)" -ForegroundColor White;
Write-Host "Report Saved: " -ForegroundColor DarkYellow -nonewline; Write-Host "$($reportPath)" -ForegroundColor White;
