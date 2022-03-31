#################################################################################
#    ___      ___ ___  ________   ________  _______   ________   _________      #
#   |\  \    /  /|\  \|\   ___  \|\   ____\|\  ___ \ |\   ___  \|\___   ___\    #
#   \ \  \  /  / | \  \ \  \\ \  \ \  \___|\ \   __/|\ \  \\ \  \|___ \  \_|    #
#    \ \  \/  / / \ \  \ \  \\ \  \ \  \    \ \  \_|/_\ \  \\ \  \   \ \  \     #
#     \ \    / /   \ \  \ \  \\ \  \ \  \____\ \  \_|\ \ \  \\ \  \   \ \  \    #
#      \ \__/ /     \ \__\ \__\\ \__\ \_______\ \_______\ \__\\ \__\   \ \__\   #
#       \|__|/       \|__|\|__| \|__|\|_______|\|_______|\|__| \|__|    \|__|   #
#                                                                               #
#                                                                               #                                                                              #
#    ___  __    ___       _______   ___  ________                               #
#   |\  \|\  \ |\  \     |\  ___ \ |\  \|\   ___  \                             #
#   \ \  \/  /|\ \  \    \ \   __/|\ \  \ \  \\ \  \                            #
#    \ \   ___  \ \  \    \ \  \_|/_\ \  \ \  \\ \  \                           #
#     \ \  \\ \  \ \  \____\ \  \_|\ \ \  \ \  \\ \  \                          #
#      \ \__\\ \__\ \_______\ \_______\ \__\ \__\\ \__\                         #
#       \|__| \|__|\|_______|\|_______|\|__|\|__| \|__|                         #
#                                                                               #
#                                                                               #
#                        OfficeGrip SD Toolset v0.9                             #
#                               22/03/2022                                      #
#                                                                               #                                                              
#################################################################################


function Show-Menu {
    param (
        [string]$Title = 'Office Grip SD Toolset v.0.9'
    )
    Clear-Host
    Write-Host "================ $Title ================"

    Write-Host ""
    Write-Host "Computer Specifications." -ForegroundColor Yellow
    Write-Host "1: Basic Info."
    Write-Host "2: Advanced Info."
    Write-Host "3: Processor Info."
    Write-Host "4: Storage Info"
    Write-Host ""
    Write-Host "Network Specifications" -ForegroundColor Green
    Write-Host "5: Network Adapter Info."
    Write-Host "6: IPCONFIG."
    Write-Host ""
    Write-Host "Cleanup" -ForegroundColor Cyan
    Write-Host "7: Teams Cleanup"
    Write-Host "8: Chrome Cleanup"
    Write-Host "9: Edge Cleanup"
    Write-Host ""
    Write-Host "Q: Press 'Q' to quit." -ForegroundColor Red
    Write-Host ""
}
do
    {
        Show-Menu
        $selection= Read-Host "Please make a selection"
        switch ($selection)
    {

                          ###########################
                          #                         #
                          #       Selection         #
                          #                         #
                          ###########################

                '1' {
                    Write-Host ""
                    for ($i = 1; $i -le 100; $i++ )
                    {
                        Write-Progress -Activity "Scan in Progress" -Status "$i% Complete:" -PercentComplete $i
                        Start-Sleep -Milliseconds 15
                    }     
                    $pc = Invoke-Expression hostname
                    $computerSystem = get-wmiobject Win32_ComputerSystem -ComputerName $pc
                    $computerBIOS = get-wmiobject Win32_BIOS -ComputerName $pc
                    $computerOS = get-wmiobject Win32_OperatingSystem -ComputerName $pc
                    "System Information for: " + $computerSystem.Name
                    "Manufacturer: " + $computerSystem.Manufacturer
                    "Model: " + $computerSystem.Model
                    "BIOS Version: " + $computerBIOS.SMBIOSBIOSVersion
                    "Serial Number: " + $computerBIOS.SerialNumber
                    "Operating System: " + $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
                    "Total Memory in Gigabytes: " + $computerSystem.TotalPhysicalMemory/1gb
                    "User logged In: " + $computerSystem.UserName
                    "Last Reboot: " + $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
                    Write-Host "Scan complete!" -ForegroundColor Green
                }
                
                # 2nd selection..
                '2' {
                    Get-ComputerInfo
                    Write-Host "Scan complete!" -ForegroundColor Green

                }
                # 3rd selection..
                '3' {
                    Get-ComputerInfo Property "*processor*"
                    Write-Host "Scan complete!" -ForegroundColor Green

                }
                # 4th selection..
                '4' {
                    for ($i = 1; $i -le 100; $i++ )
                    {
                        Write-Progress -Activity "Scan in Progress" -Status "$i% Complete:" -PercentComplete $i
                        Start-Sleep -Milliseconds 3
                    }
                    Get-Disk
                    Write-Host "Scan complete!" -ForeGroundColor Green
                }
                # 5th selection..
                '5' {
                    for ($i = 1; $i -le 100; $i++ )
                    {
                        Write-Progress -Activity "Scan in Progress" -Status "$i% Complete:" -PercentComplete $i
                        Start-Sleep -Milliseconds 0.5
                    }
                    Write-Host "Visible Adapters:" -ForegroundColor Green
                    Get-NetAdapter -Name *
                    Write-Host ""
                    Write-Host "Physical Adapters:" -ForegroundColor Green
                    Get-NetAdapter -Name * -Physical
                    Write-Host ""
                    Write-Host "Hidden Adapters:" -ForegroundColor Green
                    Get-NetAdapter -Name * -IncludeHidden
                    Write-Host ""
                    Write-Host "Scan complete!" -ForegroundColor Green
                }
                # 6th selection..
                '6' {
                    for ($i = 1; $i -le 100; $i++ )
                    {
                        Write-Progress -Activity "Scan in Progress" -Status "$i% Complete:" -PercentComplete $i
                        Start-Sleep -Milliseconds 0.1
                    }
                    Write-Host "IPCONFIG" -ForegroundColor Green
                    ipconfig
                    Write-Host ""
                    Write-Host "IP Configuration" -ForeGroundColor Green
                    Get-NetIPConfiguration
                }
                # 7th selection..
                '7' {
                    $challenge = Read-Host "Office Grip Support - Are you sure you want to delete the Teams cache (Y/N)?"
                    $challenge = $challenge.ToUpper()
                    if ($challenge -eq "N"){
                        Stop-Process -Id $PID
                        }elseif ($challenge -eq "Y"){
                            Write-Host "Office Grip Support - Stopping Teams Process" -ForegroundColor Yellow
                            try{
                                Get-Process -ProcessName Teams | Stop-Process -Force
                                Start-Sleep -Seconds 3
                                Write-Host "Office Grip Support - Teams Process Sucessfully Stopped" -ForegroundColor Green
                            }catch{
                                Write-Output $_
                            }
                            Write-Host "Office Grip Support - Clearing Teams Disk Cache" -ForegroundColor Yellow
                            Try{
                                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\application cache\cache" | Remove-Item -Confirm:$false
                                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\blob_storage" | Remove-Item -Confirm:$false
                                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\databases" | Remove-Item -Confirm:$false
                                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\cache" | Remove-Item -Confirm:$false
                                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\gpucache" | Remove-Item -Confirm:$false
                                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Indexeddb" | Remove-Item -Confirm:$false
                                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Local Storage" | Remove-Item -Confirm:$false
                                Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\tmp" | Remove-Item -Confirm:$false
                                Write-Host "Office Grip Support - Teams Disk Cache Cleaned" -ForegroundColor Green
                            }catch{
                                Write-Output $_
                            }
                            Write-Host "Office Grip Support - Cleanup Complete... Launching Teams" -ForegroundColor Green
                            Start-Process -FilePath $env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe
                            Stop-Process -Id $PID
                        }else{
                            Stop-Process -Id $PID
                        }
                        Write-Host "Complete" -ForegroundColor Green
                            }
                # 8th selection...
                '8'{
                    Write-Host "Office Grip Support - Stopping Chrome Process" -ForegroundColor Yellow
                    try{
                        Get-Process -ProcessName Chrome | Stop-Process -Force
                        Start-Sleep -Seconds 3
                        Write-Host "Office Grip Support - Chrome Process Sucessfully Stopped" -ForegroundColor Green
                    }catch{
                        echo $_
                    }
                    Write-Host "Office Grip Support - Clearing Chrome Cache" -ForegroundColor Yellow
                    try{
                        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Code Cache\Js" | Remove-Item -Confirm:$false
                        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Cache" -File | Remove-Item -Confirm:$false
                        Write-Host "Office Grip Support - Chrome Cache cleared" -ForegroundColor Green
                    }catch{
                        echo $_
                    }
                    Write-Host "Complete" -ForegroundColor Green
                    }
                # 9th selection
                '9'{
                    Write-Host "Office Grip Support - Stopping Edge Process" -ForegroundColor Yellow
                    try{
                        Get-Process -ProcessName MicrosoftEdge | Stop-Process -Force
                        Get-Process -ProcessName msedge | Stop-Process -Force
                        Get-Process -ProcessName IExplore | Stop-Process -Force
                        Write-Host "Office Grip Support - Edge and Internet Explorer Processes Sucessfully Stopped" -ForegroundColor Green
                    }catch{
                        echo $_
                    }
                    Write-Host "Office Grip Support - Clearing Edge Cache" -ForegroundColor Yellow
                    try{
                        RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 8
                        RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 2
                        Write-Host "Office Grip Support - Edge and IE Cleaned" -ForegroundColor Green
                    }catch{
                        echo $_
                    }
                    Write-Host "Complete" -ForegroundColor Green
                }
                    }
                    pause
                }
                until ($selection -eq 'q')