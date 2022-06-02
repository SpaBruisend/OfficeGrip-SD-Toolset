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
#                        OfficeGrip Cache Cleaner v0.3                          #
#                               31/05/2022                                      #
#                                                                               #                                                              
#################################################################################


function Show-Menu {
    param (
        [string]$Title = "OfficeGrip Cache Cleanup"
    )
    Clear-Host
    Write-Host "================ $Title ================"

    Write-Host ""
    Write-Host "Cleanup" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1: Teams Cleanup"
    Write-Host ""
    Write-Host "2: Chrome Cleanup"
    Write-Host ""
    Write-Host "3: Edge Cleanup"
    Write-Host ""
    Write-Host "======================="
    Write-Host "Q: Press 'Q' to quit." -ForegroundColor Red
    Write-Host "======================="
    Write-Host ""
}
do
{
    Show-Menu
    $selection = Read-Host "Please make a selection"
    switch ($selection)
    {
        '1' {
            # Teams Cache Cleanup
            $challenge = Read-Host "OfficeGrip Support - Do you want to clear the Teams Cache (Y/N)?"
            $challenge = $challenge.ToUpper()
            if ($challenge -eq "N"){
                Stop-Process -Id $PID
            }elseif ($challenge -eq "Y"){
                Write-Host "OfficeGrip Support - Stopping Teams Process" -ForegroundColor Yellow
            try{
                Get-Process -ProcessName Teams | Stop-Process -Force
                Start-Sleep -Seconds 3
                Write-Host "OfficeGrip Support - Teams Process Succesfully Stopped" -ForegroundColor Green
            }catch{
                Write-Output $_
            }
            Write-Host "OfficeGrip Support - Clearing Teams Disk Cache" -ForegroundColor Yellow
            try{
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
            Write-Host "OfficeGrip Support - Cleanup Complete... Launching Teams" -ForegroundColor Green
            Start-Process -FilePath $env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe
            Stop-Process -Id $PID
        }else{
            Stop-Process -Id $PID
        }
        Write-Host "Complete" -ForegroundColor Green
    }
        '2' {
            #Chrome Cache Cleanup
            Write-Host "OfficeGrip Support - Stopping Chrome Process" -ForgroundColorYellow
            try{
                Get-Process -ProcessName Chrome| Stop-Process -Force
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
        '3'{
            #Edge and IE Cache Cleanup
            Write-Host "OfficeGrip Support - Stopping Edge and IE Process" -ForegroundColor Yellow
            try{
                Get-Process -ProcessName MicrosoftEdge | Stop-Process -Force
                Get-Process -ProcessName msedge | Stop-Process -Force
                Get-Process -ProcessName IExplore | Stop-Process -Force
                Write-Host "Office Grip Support - Edge and Internet Explorer Processes Sucessfully Stopped" -ForegroundColor Green
            }catch{
                echo $_
            }
            Write-Host "OfficeGrip Support - Clearing Edge and IE Cache" -ForegroundColor Yellow
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