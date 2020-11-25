<#
Script to rebuild, repair, redownload, reinstall, and reregister windows updates 
By: Alan Newingham
Date: 11/25/2020

Parts taken from my other update scripts and following site: https://answers.microsoft.com/en-us/insider/forum/insider_wintp-insider_update-insiderplat_pc/windows-update-database-corruption/7d4a68e8-cad3-422b-a54b-d5c17e397319
PSWindowsUpdate module comes from the PSGallery and needs to be trusted: https://msdn.microsoft.com/en-us/powershell/gallery/psgallery/psgallery_gettingstarted

#>

function regServer32 {
	<#
	    .SYNOPSIS
               simple function to use regsvr32 
               assumes file path $env:systemroot\system32 so you only need the file name. 
    	.Example 
                regServer32 atl.dll
	#>
    [CmdletBinding()]
    param (
        [ValidateScript({ Test-Path -Path $_ -PathType 'Leaf' })]
        [string]$FilePath
    )
    process {
        try {
            Set-Location $env:systemroot\system32 
            $Result = Start-Process -FilePath 'regsvr32.exe' -Args "/s $FilePath" -Wait -NoNewWindow -PassThru
		} catch {
        	Write-Error $_.Exception.Message $false
		}
	}
} 

function registerDLL {
<#
    .SYNOPSIS
            simple function to register a lot of DLL's
            requires on regServer32 function
    .Example 
            registerDLL
#>
    process {
            try {
                regServer32 atl.dll 
                regServer32 urlmon.dll 
                regServer32 mshtml.dll 
                regServer32 shdocvw.dll 
                regServer32 browseui.dll 
                regServer32 jscript.dll 
                regServer32 vbscript.dll 
                regServer32 scrrun.dll 
                regServer32 msxml.dll 
                regServer32 msxml3.dll 
                regServer32 msxml6.dll 
                regServer32 actxprxy.dll 
                regServer32 softpub.dll 
                regServer32 wintrust.dll 
                regServer32 dssenh.dll 
                regServer32 rsaenh.dll 
                regServer32 gpkcsp.dll 
                regServer32 sccbase.dll 
                regServer32 slbcsp.dll 
                regServer32 cryptdlg.dll 
                regServer32 oleaut32.dll 
                regServer32 ole32.dll 
                regServer32 shell32.dll 
                regServer32 initpki.dll 
                regServer32 wuapi.dll 
                regServer32 wuaueng.dll 
                regServer32 wuaueng1.dll 
                regServer32 wucltui.dll 
                regServer32 wups.dll 
                regServer32 wups2.dll 
                regServer32 wuweb.dll 
                regServer32 qmgr.dll 
                regServer32 qmgrprxy.dll 
                regServer32 wucltux.dll 
                regServer32 muweb.dll 
                regServer32 wuwebv.dll
        } catch {
                Write-Error $_.Exception.Message $false
        }
    } 
}



function rebuildUpdates {
<#
	    .SYNOPSIS
               simple function to rebuild windows updates 
    	.Example 
                rebuildUpdates
#>

    process {
            try {
                $bit = Get-CIMinstance -Class Win32_Processor -ComputerName LocalHost | Select-Object AddressWidth

                Stop-Service -Name BITS 
                Stop-Service -Name wuauserv 
                Stop-Service -Name appidsvc 
                Stop-Service -Name cryptsvc 

                Remove-Item "$env:allusersprofile\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -ErrorAction SilentlyContinue 
                Rename-Item $env:systemroot\SoftwareDistribution SoftwareDistribution.bak -ErrorAction SilentlyContinue 
                Rename-Item $env:systemroot\System32\Catroot2 catroot2.bak -ErrorAction SilentlyContinue 
                Remove-Item $env:systemroot\WindowsUpdate.log -ErrorAction SilentlyContinue 

                sc.exe sdset bits 'D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)'
                sc.exe sdset wuauserv 'D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)'

                REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v AccountDomainSid /f 
                REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v PingID /f 
                REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientId /f 

                netsh winsock reset 
                netsh winhttp reset proxy 
    
                Get-BitsTransfer | Remove-BitsTransfer 

                registerDLL

                Start-Service -Name BITS 
                Start-Service -Name wuauserv 
                Start-Service -Name appidsvc 
                Start-Service -Name cryptsvc 
                
                wuauclt /resetauthorization /detectnow 
            } catch {
                Write-Error $_.Exception.Message $false
	   }
	}
}




function getUpdates {
<#
	    .SYNOPSIS
               simple function to rebuild, repair, redownload, reregister windows updates 

    	.Example 
                getUpdates
#>

    process {
            try {	
                
                #Rebuild Updates
                rebuildUpdates

                # Install Windows Update
                Install-PackageProvider NuGet -Force
                Import-PackageProvider NuGet -Force

                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

                # Now actually do the update and reboot if necessary
                Install-Module PSWindowsUpdate
                Get-Command –module PSWindowsUpdate
                Add-WUServiceManager -ServiceID 7971f918-a847-4430-9279-4a52d1efe18d -Confirm:$false
                Get-WUInstall –MicrosoftUpdate –AcceptAll –AutoReboot

            } catch {
                    Write-Error $_.Exception.Message $false
		}
	}
}

getUpdates

#run system checks
sfc /scannow 
DISM.exe /Online /Cleanup-image /scanhealth 
DISM.exe /Online /Cleanup-image /Restorehealth

