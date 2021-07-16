$ServerInfo=Get-ComputerInfo
$OS=$ServeInfo.WindowsProductName

new-item -type directory -path C:\temp\patches -Force

if ( $OS -like "*Microsoft Windows Server 2012 R2*" )
 {
                #PATCH1
                $source1 = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/04/windows8.1-kb5001403-x64_7f15c4b281f38d43475abb785a32dbaf0355bad5.msu'
                $destination1 = "C:\temp\patches\windows8.1-kb5001403-x64_7f15c4b281f38d43475abb785a32dbaf0355bad5.msu"
                Invoke-WebRequest -Uri $source1 -OutFile (New-Item -path $destination1 -force)
                    If (Test-Path -Path $destination1 ) {
                         Write-Host "kb5001403 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5001403 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        }
                #PATCH2
                $source2 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows8.1-kb5004954-x64_691dc48f8697e7dd2d138d8c6ac2a92d27927467.msu'
                $destination2 = "C:\temp\patches\windows8.1-kb5004954-x64_691dc48f8697e7dd2d138d8c6ac2a92d27927467.msu"
                Invoke-WebRequest -Uri $source2 -OutFile (New-Item -path $destination2 -force)
                    If (Test-Path -Path $destination2 ) {
                         Write-Host "kb5004954 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004954 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        }
                #PATCH3
                $source3 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows8.1-kb5004958-x64_8b73440b9c53bcea2660d9409b6ad3920f104cd2.msu'
                $destination3 = "C:\temp\patches\windows8.1-kb5004958-x64_8b73440b9c53bcea2660d9409b6ad3920f104cd2.msu"
                
                Invoke-WebRequest -Uri $source3 -OutFile (New-Item -path $destination3 -force)
                    If (Test-Path -Path $destination3 ) {
                         Write-Host "kb5004958 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004958 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        }
}
elseif ( $OS -like "*Microsoft Windows Server 2012*" )
{
                #PATCH1
                $source1 = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/04/windows8-rt-kb5001401-x64_1027ae2c9888c2dfe0caadeafc506b3012789c56.msu'
                $destination1 = "C:\temp\patches\windows8-rt-kb5001401-x64_1027ae2c9888c2dfe0caadeafc506b3012789c56.msu"
                
                Invoke-WebRequest -Uri $source1 -OutFile (New-Item -path $destination1 -force)
                    If (Test-Path -Path $destination1 ) {
                         Write-Host "kb5001401 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5001401 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        }
                        #PATCH2
                $source2 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows8-rt-kb5004956-x64_3291813fb4b21c3308452ee0914e0dd06427d608.msu'
                $destination2 = "C:\temp\patches\windows8-rt-kb5004956-x64_3291813fb4b21c3308452ee0914e0dd06427d608.msu"
                
                Invoke-WebRequest -Uri $source2 -OutFile (New-Item -path $destination2 -force)
                     If (Test-Path -Path $destination2 ) {
                         Write-Host "kb5004956 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004956 Please download manually" -ForegroundColor red -backgroundcolor black
                        }    
                        #PATCH3           
                $source3 = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/07/windows8-rt-kb5004960-x64_15b7362a1146198852b2be37c4997f81e16495b6.msu'
                $destination3 = "C:\temp\patches\windows8-rt-kb5004960-x64_15b7362a1146198852b2be37c4997f81e16495b6.msu"
                
                Invoke-WebRequest -Uri $source3 -OutFile (New-Item -path $destination3 -force)
                     If (Test-Path -Path $destination3 ) {
                         Write-Host "kb5004960 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004960 Please download manually" -ForegroundColor red -backgroundcolor black
                        } 
            
}

elseif ( $OS -like "*Microsoft Windows Server 2016*" )
{
                
                $source1 = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/04/windows10.0-kb5001402-x64_0108fcc32c0594f8578c3787babb7d84e6363864.msu'
                $destination1 = "C:\temp\patches\windows10.0-kb5001402-x64_0108fcc32c0594f8578c3787babb7d84e6363864.msu"
                
                Invoke-WebRequest -Uri $source1 -OutFile (New-Item -path $destination1 -force)
                     If (Test-Path -Path $destination1 ) {
                         Write-Host "kb5001402 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5001402 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        } 
                $source2 = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/07/windows10.0-kb5004948-x64_206b586ca8f1947fdace0008ecd7c9ca77fd6876.msu'
                $destination2 = "C:\temp\patches\windows10.0-kb5004948-x64_206b586ca8f1947fdace0008ecd7c9ca77fd6876.msu"
                
                Invoke-WebRequest -Uri $source2 -OutFile (New-Item -path $destination2 -force)
                      If (Test-Path -Path $destination2 ) {
                         Write-Host "kb5004948 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004948 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        } 
            }
            
                elseif ( $OS -like "*Microsoft Windows Server 2019*" )
            {
                
                $source1 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/06/windows10.0-kb5003711-x64_577dc9cfe2e84d23b193aae2678b12e777fc7e55.msu'
                $destination1 = "C:\temp\patches\windows10.0-kb5003711-x64_577dc9cfe2e84d23b193aae2678b12e777fc7e55.msu"
                
                Invoke-WebRequest -Uri $source1 -OutFile (New-Item -path $destination1 -force)
                      If (Test-Path -Path $destination1 ) {
                         Write-Host "kb5003711 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5003711 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        } 
                $source2 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows10.0-kb5004947-x64_c00ea7cdbfc6c5c637873b3e5305e56fafc4c074.msu'
                $destination2 = "C:\temp\patches\windows10.0-kb5004947-x64_c00ea7cdbfc6c5c637873b3e5305e56fafc4c074.msu"
                
                Invoke-WebRequest -Uri $source2 -OutFile (New-Item -path $destination2 -force)
                      If (Test-Path -Path $destination2 ) {
                         Write-Host "kb5004947 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004947 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        } 
            }

        #Write-Host "$($pserver) OS: $($OS)"

        $UpdatePath = "C:\temp\patches"

        # Old hotfix list
        #Get-HotFix > "$UpdatePath\old_hotfix_list.txt"

        # Get all updates
        $Updates = Get-ChildItem -Path $UpdatePath -Recurse | Where-Object {$_.Name -like "*msu*"}

        # Iterate through each update
        ForEach ($update in $Updates) {

            # Get the full file path to the update
            $UpdateFilePath = $update.FullName

            # Logging
            write-host "Installing update $($update.BaseName)"

            # Install update - use start-process -wait so it doesnt launch the next installation until its done
            Start-Process -wait wusa -ArgumentList "/update $UpdateFilePath","/quiet","/norestart"
        }

#Restart-Computer -force
