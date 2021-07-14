#self-escalating lines

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process PowerShell -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
    exit;
}

##
start-transcript

$AllServers=Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' -Properties ServerName,Operatingsystem,OperatingSystemVersion,IPv4Address |Sort-Object -Property Operatingsystem |Select-Object -Property Name,Operatingsystem,IPv4Address

$printservers= @()
$spoolDisabled = @()
$rdshost = @()
$PatchServers =@()

$Credes = Get-Credential

foreach ($server in $allServers.name)
{
    If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $server -Quiet)
    {
        #$sharedprinters = get-printer -ComputerName $server | where {$_.shared}
        $sharedprinters = Get-CimInstance Win32_Printer -ComputerName $server | where {$_.ShareName -ne $Null}
        if ($sharedprinters.Count -ge 1)
        {
            Write-Host "$($Server) is a Print Server. Patch it."
            $printservers += $Server
        } elseif (get-windowsfeature -Name RDS-Connection-Broker -ComputerName $server|Where Installed) {
            Write-Host "$($Server) is a Terminal Server. Patch it."
            $rdshost += $Server
        } 
        
        else {
            try{
                Invoke-Command -ComputerName $server -ScriptBlock { (Stop-Service -Name Spooler -Force), (Set-Service -Name Spooler -StartupType Disabled) }
                Write-Host "$($Server) spooler service has been stopped and disabled."
                $spoolDisabled += $Server
            } catch {
                Write-Host "$($Server) Error: unable to stop/disable spooler service."
            }
        }   
        
    } else {
            write-host "$Server is Offline"
        }
}
#Listing Servers to be batched and Servers wehre spooler is disabled
write-host "Print Server(s). Patch these:-"  -ForegroundColor black -backgroundcolor yellow
foreach ($pserver in $printservers)
    {
        $OS = $(((gcim Win32_OperatingSystem -ComputerName $pserver).Name).split(‘|’)[0])
        Write-Host "$($pserver) OS: $($OS)" -ForegroundColor green -backgroundcolor black
     }
write-host "Terminal Server(s). Patch these:-" -ForegroundColor black -backgroundcolor yellow
foreach ($RDH in $rdshost)
    {
        $OS = $(((gcim Win32_OperatingSystem -ComputerName $RDH).Name).split(‘|’)[0])
        Write-Host "$($RDH) OS: $($OS)"  -ForegroundColor green -backgroundcolor black
     }
write-host "Print Spools Disabled: $($spoolDisabled)." -ForegroundColor black -backgroundcolor yellow

#read-host “Press ENTER to continue...to attempt to download patches to Servers automatically"
#Listing PrintServers and attempting to download patch fliles in c:\Temp\ folder at destination
$PatchServers = $printservers + $rdshost
 #echo $PatchServers
#attempting to download patch fliles in c:\Temp\ folder at destination
write-host "Attempting to download patch files to Servers that need to be patched"  -ForegroundColor green -backgroundcolor black
foreach ($pserver in $PatchServers)
    {
        $OS = $(((gcim Win32_OperatingSystem -ComputerName $pserver).Name).split(‘|’)[0])
            if ( $OS -like "*Microsoft Windows Server 2012 R2*" )
            {
                #PATCH1
                $source1 = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/04/windows8.1-kb5001403-x64_7f15c4b281f38d43475abb785a32dbaf0355bad5.msu'
                $destination1 = "\\$pserver\c$\temp\patches\windows8.1-kb5001403-x64_7f15c4b281f38d43475abb785a32dbaf0355bad5.msu"
                Invoke-WebRequest -Uri $source1 -OutFile (New-Item -path $destination1 -force)
                    If (Test-Path -Path $destination1 ) {
                         Write-Host "kb5001403 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5001403 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        }
                #PATCH2
                $source2 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows8.1-kb5004954-x64_691dc48f8697e7dd2d138d8c6ac2a92d27927467.msu'
                $destination2 = "\\$pserver\c$\temp\patches\windows8.1-kb5004954-x64_691dc48f8697e7dd2d138d8c6ac2a92d27927467.msu"
                Invoke-WebRequest -Uri $source2 -OutFile (New-Item -path $destination2 -force)
                    If (Test-Path -Path $destination2 ) {
                         Write-Host "kb5004954 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004954 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        }
                #PATCH3
                $source3 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows8.1-kb5004958-x64_8b73440b9c53bcea2660d9409b6ad3920f104cd2.msu'
                $destination3 = "\\$pserver\c$\temp\patches\windows8.1-kb5004958-x64_8b73440b9c53bcea2660d9409b6ad3920f104cd2.msu"
                
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
                $destination1 = "\\$pserver\c$\temp\patches\windows8-rt-kb5001401-x64_1027ae2c9888c2dfe0caadeafc506b3012789c56.msu"
                
                Invoke-WebRequest -Uri $source1 -OutFile (New-Item -path $destination1 -force)
                    If (Test-Path -Path $destination1 ) {
                         Write-Host "kb5001401 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5001401 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        }
                        #PATCH2
                $source2 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows8-rt-kb5004956-x64_3291813fb4b21c3308452ee0914e0dd06427d608.msu'
                $destination2 = "\\$pserver\c$\temp\patches\windows8-rt-kb5004956-x64_3291813fb4b21c3308452ee0914e0dd06427d608.msu"
                
                Invoke-WebRequest -Uri $source2 -OutFile (New-Item -path $destination2 -force)
                     If (Test-Path -Path $destination2 ) {
                         Write-Host "kb5004956 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004956 Please download manually" -ForegroundColor red -backgroundcolor black
                        }    
                        #PATCH3           
                $source3 = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/07/windows8-rt-kb5004960-x64_15b7362a1146198852b2be37c4997f81e16495b6.msu'
                $destination3 = "\\$pserver\c$\temp\patches\windows8-rt-kb5004960-x64_15b7362a1146198852b2be37c4997f81e16495b6.msu"
                
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
                $destination1 = "\\$pserver\c$\temp\patches\windows10.0-kb5001402-x64_0108fcc32c0594f8578c3787babb7d84e6363864.msu"
                
                Invoke-WebRequest -Uri $source1 -OutFile (New-Item -path $destination1 -force)
                     If (Test-Path -Path $destination1 ) {
                         Write-Host "kb5001402 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5001402 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        } 
                $source2 = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2021/07/windows10.0-kb5004948-x64_206b586ca8f1947fdace0008ecd7c9ca77fd6876.msu'
                $destination2 = "\\$pserver\c$\temp\patches\windows10.0-kb5004948-x64_206b586ca8f1947fdace0008ecd7c9ca77fd6876.msu"
                
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
                $destination1 = "\\$pserver\c$\temp\patches\windows10.0-kb5003711-x64_577dc9cfe2e84d23b193aae2678b12e777fc7e55.msu"
                
                Invoke-WebRequest -Uri $source1 -OutFile (New-Item -path $destination1 -force)
                      If (Test-Path -Path $destination1 ) {
                         Write-Host "kb5003711 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5003711 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        } 
                $source2 = 'http://download.windowsupdate.com/c/msdownload/update/software/secu/2021/07/windows10.0-kb5004947-x64_c00ea7cdbfc6c5c637873b3e5305e56fafc4c074.msu'
                $destination2 = "\\$pserver\c$\temp\patches\windows10.0-kb5004947-x64_c00ea7cdbfc6c5c637873b3e5305e56fafc4c074.msu"
                
                Invoke-WebRequest -Uri $source2 -OutFile (New-Item -path $destination2 -force)
                      If (Test-Path -Path $destination2 ) {
                         Write-Host "kb5004947 downloaded on $($pserver) C:\temp\patches" -ForegroundColor green -backgroundcolor black
                        }
                        Else {
                        Write-Host "unable to download kb5004947 on $($pserver) Please download manually" -ForegroundColor red -backgroundcolor black
                        } 
            }

        Write-Host "$($pserver) OS: $($OS)"
     }



#read-host “Press ENTER to continue...to attempt to install downloaded patches on Servers automatically"

#write-host "Attempting to install downloaded patches on Print and Terminal Servers"  -ForegroundColor green -backgroundcolor black

#get Domain Admin credentials
#$cred = (Get-Credential)

#read-host “Press ENTER to continue...to attempt to place patch install scripts on target servers"

write-host "Placing patchinstall scripts on serves that need to be batched"  -ForegroundColor green -backgroundcolor black

#get Domain Admin credentials
#$cred = (Get-Credential)

foreach ($pserver in $PatchServers)
    {

    Copy-Item -Path .\installpatches.ps1 -Destination \\$pserver\c$\temp\patches\installpatches.ps1
    Copy-Item -Path .\installpatches.bat -Destination \\$pserver\c$\temp\patches\installpatches.bat
    Write-Host "Run installpatches.bat as ADMINISTRAOTR on $($pserver) located under C:\temp\patches" -ForegroundColor green -backgroundcolor black    
    

        #Write-Host "$($pserver) OS: $($OS)"
     }


#Create a scheduled task to trigger msu install scripts on servers that need to be patched and triggering them. delete task after.
write-host "Attempting to install downloaded patches on Print and Terminal Servers. "  -ForegroundColor green -backgroundcolor black
write-host "!!!!TARGET SERVERS WILL BE REBOOTED!!!! Launch install scripts manually if you dont intend to autoreboot"  -ForegroundColor red -backgroundcolor black

#read-host “Press ENTER to continue..."

#$Credes = Get-Credential

foreach ($pserver in $PatchServers)
{
#$Credes = Get-Credential
Invoke-Command -ComputerName $pserver -Credential $Credes -Scriptblock {
                $command = { powershell.exe -ExecutionPolicy Bypass c:\temp\patches\installpatches.ps1 -RunType $true -Path c:\temp\patches;}

                $TaskName = "PrintNightmare"
                $User = [Security.Principal.WindowsIdentity]::GetCurrent()
                $Scheduler = New-Object -ComObject Schedule.Service

                $Task = $Scheduler.NewTask(0)

                $RegistrationInfo = $Task.RegistrationInfo
                $RegistrationInfo.Description = $TaskName
                $RegistrationInfo.Author = $User.Name

                $Settings = $Task.Settings
                $Settings.Enabled = $True
                $Settings.StartWhenAvailable = $True
                $Settings.Hidden = $False

                $Action = $Task.Actions.Create(0)
                $Action.Path = "powershell"
                $Action.Arguments = "-Command $command"

                $Task.Principal.RunLevel = 1

                $Scheduler.Connect()
                $RootFolder = $Scheduler.GetFolder("\")
                $RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, "SYSTEM", $Null, 1) | Out-Null
                $RootFolder.GetTask($TaskName).Run(0) | Out-Null
                sleep 2
                Do {
	                Write-Host "Waiting for scheduled task to finish"
	                sleep 360
                } 
                Until ($RootFolder.GetTask($TaskName).State -eq 3)
                $RootFolder.DeleteTask($TaskName,0)
                Restart-Computer -Force
}

}

Stop-transcript
