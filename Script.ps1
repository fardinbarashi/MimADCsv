<# 
1.The script retrieves active users from MIM and saves it in a csv file $FromMim.
2.The script retrieves active users from AD and saves it in a csv file$FromAD.
3.The script creates a third csv file, $CombinedCsv
#>

Start-Transcript -Path "$PSScriptRoot\TranscriptLogFiles\Transcript_PowerShell_Log.txt"

Write-host  "===================================="
Write-host  "Initialization script..."
Write-host  ""

Try
 { # Start Try, Step 1 : Import Modules, Create Filepath setting
   Write-host "Step 1 : Import Modules, Create Filepath settings.. 0%" -ForegroundColor Yellow
   Import-Module LithnetRMA
   Write-host "Import Module LithnetRMA" 

   Import-Module ActiveDirectory
   Write-host "Import Module ActiveDirectory" 
   
   # Init connection to fim
   Write-host  "Initializing connection to MIM"
   Set-ResourceManagementClient -BaseAddress "http://localhost:5725" # Set host to query, should be localhost unless run from another server.

   
   # Export location
   $FromMim = "$PSScriptRoot\CsvFiles\FromMim\FromMim.csv"
   $FromAD = "$PSScriptRoot\CsvFiles\FromAD\FromAD.csv"
   $CombinedCsv = "$PSScriptRoot\CsvFiles\CombinedCsv\Allusers.csv"
   
   # FileCheck $FromMim
   $FileCheckFromMim = Get-ChildItem -Path $FromMim -Recurse -Force | Test-Path -PathType Leaf
   If( $FileCheckFromMim -Eq $Null )
    {
     # $FileCheckFromMim Is Empty
    } 
   Else 
    {
     Write-host "Removing Previous $FromMim" -ForegroundColor Magenta
     Get-ChildItem -Path $FromMim -Force -Recurse | Remove-Item -Force
    }
   
   # FileCheck $FromAD 
   $FileCheckFromAD = Get-ChildItem -Path $FromAD -Recurse -Force | Test-Path -PathType Leaf
   If( $FileCheckFromAD -Eq $Null )
    {
     # $FileCheckFromAD Is Empty
    } 
   Else 
    {
     Write-host "Removing Previous $FromAD" -ForegroundColor Magenta
     Get-ChildItem -Path $FromAD -Force -Recurse | Remove-Item -Force
    }
   
   #  FileCheck $CombinedCsv
   $FileCheckCombinedCsv = Get-ChildItem -Path $CombinedCsv -Recurse -Force | Test-Path -PathType Leaf
   If( $FileCheckCombinedCsv -Eq $Null )
    {
     # $FileCheckCombinedCsv Is Empty
    } 
   Else 
    {
     Write-host "Removing Previous $CombinedCsv" -ForegroundColor Magenta
     Get-ChildItem -Path $CombinedCsv -Force -Recurse | Remove-Item -Force
    }

   Write-host "Step 1 : Import Modules, Create Filepath settings... 100%" -ForegroundColor Green
   Write-host  ""

 } # End Try, Import Modules, Create Filepath setting

Catch
 { # Start Catch, Import Modules
  Write-Warning "Error"
  Write-Warning "Could not Import Modules, Create Filepath settings "
  Write-host "Step 1 : Import Modules, Create Filepath settings.. %"  -ForegroundColor Red
  Write-Host $Error[0];
 } # End Catch, Import Modules

Try
 { # Start Try, Query FIM, Save as CSV in $FromMim
   Write-host "Step 2 : Query FIM, Save as CSV in $FromMim.. 0%" -ForegroundColor Yellow
   Write-host "Querying MIM for active users"

   $XPath = "/Person[starts-with(EmployeeID,'%')]"   
   $Persons = Search-Resources -XPath $XPath -AttributesToGet (
   "FirstName","LastName",
   "AccountName","ManagerAccountName")
   
   #filter out generic accounts that are not people.
    $Personsdetails = $Persons | 
    Where-Object{ 
    # Start Where-Object
     $_.AccountName -ne "XXX" -and
     $_.AccountName -ne "ZZZ" -and
     $_.AccountName -ne "OOO"
    # End Where-Object 
    } |
    Select-Object @{N='FirstName';E={$_.FirstName}},
                  @{N='LastName';E={$_.LastName}},
                  @{N='Username';E={$_.AccountName}},               
                  @{N='MimManager';E={$_.ManagerAccountName}}, 
                  @{N='UserPrincipalName';E={}} |
    Export-csv -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Path $FromMim

    Write-host ("{0} People added to export" -f $persons.Count)

    Write-host  "Step 2 : Query FIM, Save as CSV in $FromMim... 100%" -ForegroundColor Green
    Write-host  ""
 } # End Try, Query FIM, Save as CSV in $FromMim

Catch
 { # Start Catch, Query FIM, Save as CSV in $FromMim
  Write-Warning "Error"
  Write-Warning "Could not Import Modules, Create Filepath settings "
  Write-host "Step 2 : Import Modules, Create Filepath settings.. %"  -ForegroundColor Red
  Write-Host $Error[0];
 } # End Catch, Query FIM, Save as CSV in $FromMim 

Try
 { # Start Try, Export Active AD-Users with Manager,save as csv $FromAD
   Write-host "Step 3 : Export Active AD-Users with Manager,save as csv $FromAD.. 0%" -ForegroundColor Yellow
   Write-host "Querying AD for active users" 
    $GetAdusers = Get-ADUser -SearchBase "SELECT YOUR OU" -Properties * -Filter * | 
    Where-Object { 
     $_.Enabled -Eq $True -and
     $_.SamAccountName -ne "XXX" -and
     $_.SamAccountName -ne "ZZZ" -and
     $_.SamAccountName -ne "OOO" 
     } |
    Select-Object @{Name='ADManager';Expression={(Get-ADUser $_.Manager).SamAccountName}},
    @{Name='UserPrincipalName';Expression={(Get-ADUser $_.Manager).UserPrincipalName}} |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Path $FromAD
   
   $CountFromAD = Get-Content $FromAD | Measure-Object | % { $_.Count }
   Write-host "$CountFromAD People added to $FromAD"
   Write-host "Step 3 : Export Active AD-Users with Manager,save as csv $FromAD... 100%" -ForegroundColor Green
   Write-host  ""
 } # End Try, IExport Active AD-Users with Manager,save as csv $FromAD

Catch
 { # Start Catch, Import $FromMim csv file, AD-Query Get out Users manager, save as csv $FromAD
  Write-Warning "Error"
  Write-Warning "Could not Export Active AD-Users with Manager,save as csv $FromAD... "
  Write-host "Step 3 : Export Active AD-Users with Manager,save as csv $FromAD.. %"  -ForegroundColor Red
  Write-Host $Error[0];
 } # End Catch, Import $FromMim csv file, AD-Query Get out Users manager, save as csv $FromAD

Try
 { # Start Try, Combine $FromMim and $FromAD. Create $CombinedCsv csv
  Write-host "Step 4 : Create CsvFile $CombinedCsv.. 0%" -ForegroundColor Yellow
   Write-Host "Import CSV-File to $FromAD"
   $CsvFromAD = Import-Csv -Path $FromAD -Delimiter ";" -Encoding UTF8  
   
   Write-Host "Import CSV-File to $FromMim"
   $CsvFromMim = Import-Csv -Path $FromMim -Delimiter ";" -Encoding UTF8 
   
   Write-Host "..Get UserPrincipalName From AD.. 0%" -ForegroundColor Yellow  
   ForEach ( $Object In $CsvFromMim )
    { # Start ForEach ( $Chefer In $CsvFromAD )
     $CombineCsv = "" | select "FirstName","LastName","Username","MimManager"       
     
     $FilterManagers = $CsvFromAD | Where-Object { $_.ADManager -Eq $Object.MimManager }
       If($FilterManagers)
        { # Start If($FilterManagers)             
            $CombineCsv.'FirstName' = $Object.'FirstName'
            $CombineCsv.'LastName' = $Object.'LastName'
            $CombineCsv.'Username' = $Object.'Username'
            $CombineCsv.'MimManager' = $Object.'MimManager'
            
            
            # Add Data From AD
            $AddManagerFromAD = ( ( Get-ADUser $CombineCsv.'MimManager' ).UserPrincipalName -split " " )[0]
            
            $CombineCsv.'UserPrincipalName' = $AddManagerFromAD

            # Create CombineCsv 
            $CombineCsv | Export-Csv -Path $CombinedCsv -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Append
        } # End If($FilterManagers)
       
       else
        { # Start Else

        } # End Else

    } # End ForEach ( $Chefer In $CsvFromAD )
  Write-Host "..Get UserPrincipalName From AD.. 100%" -ForegroundColor Green 
  Write-host  "Step 4 : Create CsvFile $CombinedCsv... 100%" -ForegroundColor Green
  Write-host  ""
 } # End Try, Combine $FromMim and $FromAD. Create $CombinedCsv csv

Catch
 { # Start Catch, Combine $FromMim and $FromAD. Create $CombinedCsv csv
  Write-Warning "Error"
  Write-Warning "Could not Create CsvFile $CombinedCsv csv "
  Write-host "Step 4 : Create CsvFile $CombinedCsv..  %"  -ForegroundColor Red
  Write-Host $Error[0];
 } # End Catch, Combine $FromMim and $FromAD. Create $CombinedCsv csv


Write-host  "Ending script..."
Write-host  "===================================="

Stop-Transcript

# Rename Logfiles
Get-ChildItem -Path "$PSScriptRoot\TranscriptLogFiles\Transcript_PowerShell_Log.txt" | 
Rename-Item -NewName {"Transcript_PowerShell_Log" + "- Date -" + (Get-Date -Format yyMMdd) + "- Time -" + (Get-Date -Format HHmmss) + ".txt"} 



