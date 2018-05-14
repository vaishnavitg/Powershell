$vmGuestCSV = "D:\Upload\ESXI\config\testguestvms.csv"
$esxiHost = ""
$arrayGuest = @{}
$skipGuest = @()

Connect-VIServer $esxiHost -User $esxihostadmin -Password $esxihostpwd
foreach ($vmGuest in $vmGuestCSV)
{
    if ($vmGuest.esxiHost -eq $esxiHost)
    {
        $selectedVm = Get-VM -Name $vmGuest.VmName

        if ($selectedVm -eq 0)
        {
            $skipGuest.Add("$($vmGuest.VmName) - $($vmGuest.VmNewName)")
            continue
        }

        $selectedVm | Set-VM -Name $vmGuest.VmNewName -Verbose -Confirm:$false
        Get-VM -Name $vmGuest.VmNewName | start-vm -Confirm:$false 
        $arrayGuest.Add($vmGuest.VmNewName, $vmGuest.ProdIpAddress, $vmGuest.RemoteIpAddress, $vmGuest.GateWay, $vmGuest.PriDns, $vmGuest.SecDns)
    }
}

get-vm -Location $esxiHost | where {($_.name -notlike "Appliance Images*") -and ($_.PowerState -match "PoweredOff")} | start-vm -Confirm:$false -RunAsync

foreach ($vmGuest in $arrayGuest)
{
    if ((Get-VM -Name $vmGuest[0]).extensiondata.Guest.ToolsStatus -ne "toolsNotRunning")
    {
        continue
    }

    Do 
    {
        $toolsStatus = (Get-VM -Name $vmGuest[0]).extensiondata.Guest.ToolsStatus
        sleep 5
    } 
    until ($toolsStatus -ne "toolsNotRunning")
}


# --- Requires declaration for every VM. Unique IP for both Unattended and Sysprep file
$sysprepPsTemplate = "D:\Upload\ESXI\config\vmguest\sysprep.ps1"

# Windows Server 2016 
$sysprepXmlTemplate = "D:\Upload\ESXI\config\vmguest\vmguesttemplate.xml"
$enableremoteps = "D:\Upload\ESXI\config\vmguest\enableremoting.ps1"


foreach($vmguest in $arrayGuest)
{
    if ($vmguest[0].StartsWith('S'))
    {	
        $sysprepXml = "D:\Upload\ESXI\config\vmguest\$($vmguest[1])$($vmguest[0]).xml"
        			
        $xmlReplaceValues = Get-Content $sysprepXmlTemplate
	    $xmlReplaceValues = $xmlReplaceValues.Replace("_newServerName_", $vmguest[0])
	    $xmlReplaceValues = $xmlReplaceValues.Replace("_prodIpAddress_", $vmguest[1])
	    $xmlReplaceValues = $xmlReplaceValues.Replace("_gateway_", $vmguest[3])
	    $xmlReplaceValues = $xmlReplaceValues.Replace("_remoteIpAddress_", $vmguest[2])
	    $xmlReplaceValues = $xmlReplaceValues.Replace("_priDns_", $vmguest[4])
	    $xmlReplaceValues = $xmlReplaceValues.Replace("_secDns_", $vmguest[5])
	    $xmlReplaceValues = $xmlReplaceValues.Replace("_priDns_", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2")
        $xmlReplaceValues| Set-Content "$($sysprepXml)"

        $sysprepPs = Get-Content $sysprepPsTemplate
		$sysprepPs = $sysprepPs.Replace("_location_","$($Destinationlocation)$($vmguest[1])$($vmguest[0]).xml") 
        $sysprepPs| Set-Content "$($vmunattendlocation).ps1" 

        Copy-VMGuestFile -Source "$($sysprepXml)" -Destination $($Destinationlocation)`
         -VM $($vmguest[0]) -LocalToGuest -GuestUser $emsadmin -GuestPassword $vmguestpwd -Force -Verbose

        Copy-VMGuestFile -Source "$($vmunattendlocation).ps1" -Destination $($Destinationlocation)`
         -VM $($vmguestname) -LocalToGuest -GuestUser $emsadmin -GuestPassword $vmguestpwd -Force -Verbose

		Copy-VMGuestFile -Source "$($enableremoteps)" -Destination $($Destinationlocation)`
         -VM $($vmguestname) -LocalToGuest -GuestUser $emsadmin -GuestPassword $vmguestpwd -Force -Verbose

		Copy-VMGuestFile -Source "$($miraCastViewsrc)" -Destination $($miraCastViewdst)`
         -VM $($vmguestname) -LocalToGuest -GuestUser $emsadmin -GuestPassword $vmguestpwd -Force -Verbose
    }
}     

$domainCredential = ""
$computerHostname = "" 
workflow ActiveDIrectory
{
    sequence
    {
        $currentMessage = "Sysprep Server - $($computerHostname)"
        InlineScript
        {
               
            return "success" 
        }
        -PSComputerName $computerHostname -PSCredential $domainCredential -PSProgressMessage $currentMessage
                
        Do 
        {
            $toolsStatus = (Get-VM -Name $computerHostname).extensiondata.Guest.ToolsStatus
            sleep 5
        } 
        until ($toolsStatus -ne "toolsNotRunning")

        $currentMessage = "Change Network Interface Name"
        InlineScript
        {
            Rename-NetAdapter -Name "Ethernet0" -NewName "Production" 
            Rename-NetAdapter -Name "Ethernet1" -NewName "Remote Admin"
            
            return "success" 
        }
        -PSComputerName $computerHostname -PSCredential $domainCredential -PSProgressMessage $currentMessage

        $currentMessage = "Join Server to $($domainAddress)"
        InlineScript
        {
            $portQueryResult = Test-NetConnection -Port 80 -InformationLevel "Quiet"
            if ($portQueryResult.Equals($true))
            {
                Add-Computer -DomainName $domainAddress -Credential $domainCredential -LocalCredential $localcredential -PassThru
                Invoke-GPUpdate -Force
                return "success"
            }
            return "failed"
        }
        -PSComputerName $computerHostname -PSCredential $domainCredential -PSProgressMessage $currentMessage
                
        Restart-Computer -Wait -PSPersist $true -PSComputerName $computerHostname -PSCredential $domainCredential
                
        $currentMessage = "Install Endpoint Protection System"
        InlineScript
        {
            Start-Process "msiexec.exe" -ArgumentList "" -NoNewWindow -Wait
            
            return "success" 
        }
        -PSComputerName $computerHostname -PSCredential $domainCredential -PSProgressMessage $currentMessage
        
        $currentMessage = "Install System Center Configuration Manager Client"
        InlineScript
        {
            Start-Process "msiexec.exe" -ArgumentList "" -NoNewWindow -Wait
            
            return "success" 
        }
        -PSComputerName $computerHostname -PSCredential $domainCredential -PSProgressMessage $currentMessage
        
        $currentMessage = "Install System Center Configuration Manager Client"
        InlineScript
        {
            Start-Process "msiexec.exe" -ArgumentList "" -NoNewWindow -Wait
            
            return "success" 
        }
        -PSComputerName $computerHostname -PSCredential $domainCredential -PSProgressMessage $currentMessage

        $currentMessage = ""
        InlineScript
        {
            
        }
        -PSComputerName $computerHostname -PSCredential $domainCredential -PSProgressMessage $currentMessage
    }
}
        # --- Declare Var for VMGuest in CSV --- #

        $vmguestname = $vmguest.name
        $vmguestnewname = $vmguest.newname
        $vmguesthostip = $vmguest.esxiip
            
        # --- END Declare for VMGuest in CSV --- #    

        # --- Display VMGuests Default Name and New Name from Host List
        Write-Host "VMGuests in VMGuest CSV List : ($($vmguestname)) ($($vmguestnewname))" -ForegroundColor Green n     

        # --- Verify the schoolHost IP in CSV with the IP in the VMGuest csv file.
        if ($($vmguesthostip) -eq $($esxiIp))
        {

            Write-Host "[INFO] --- ESXI Host Verification - VMGuest HOST : ($vmguesthostip) / ESXi Host : ($esxiIp) " -ForegroundColor Yellow 


            # Windows Server 2012 R2
            #$unattend2012templatelocation = "D:\Upload\ESXI\config\vmguest\vmguesttemplate_2012.xml"

            # --- Check the first letter to identify Windows Hosts
            if ($vmguestnewname.StartsWith($wincheck)) # --- Check for VMGuest name starts with letter 's'
            {

                Write-host "[INFO] --- Performing Sysprep Window OS Servers for: ($esxisch)" -ForegroundColor Yellow `n
                Write-Host "[INFO] --- Windows Host Verification Before Execution : ($vmguestnewname)" -ForegroundColor Yellow `n

				# --- Replacement of Values in the unattended.xml file for Server 2016
				$xmlReplaceValues = Get-Content $unattendtemplatelocation
				$xmlReplaceValues = $xmlReplaceValues.Replace("NewServerName","$($vmguest.newname)")
				$xmlReplaceValues = $xmlReplaceValues.Replace("PROD-IPAddress", "$($vmguest.ipaddress01)")
				$xmlReplaceValues = $xmlReplaceValues.Replace("PROD-Gateway", "$($vmguest.gateway01)")
				$xmlReplaceValues = $xmlReplaceValues.Replace("MGMT-IPAddress", "$($vmguest.ipaddress02)")
				$xmlReplaceValues = $xmlReplaceValues.Replace("PriDNSIP", "$($vmguest.dns01)")
				$xmlReplaceValues = $xmlReplaceValues.Replace("SecDNSIP", "$($vmguest.dns02)")    
				$xmlReplaceValues = $xmlReplaceValues.Replace("SecDNSIP", "$($vmguest.dns02)")               
#
                $BaseXMLTemplate = $xmlReplaceValues.Replace("Management", "Remote Admin")| Set-Content "$($vmunattendlocation).xml" # --- Added to replace Management NIC Name to Remote Admin
				
                # --- Replacement values for Sysprep in vmguest
                $ps1ReplaceValues = Get-Content $syspreppslocation
				$BasePSTemplate = $ps1ReplaceValues.Replace("_location_","$($Destinationlocation)$($vmguesthostip)$($vmguestnewname).xml") | Set-Content "$($vmunattendlocation).ps1" 
	
				# --- Define the new name of the sysprep powershell file at the Staging Server
				$sysprepps = "$($Destinationlocation)$($vmguesthostip)$($vmguestnewname).ps1"	

                # --- Define the Cyber Arc route
                $CyberArcIPDc3 = "10.166.89.0"
                $CyberArcIPDc4 = "10.168.89.0"
                $CyberArcSN = "255.255.255.128"
                
               
                If ($vmguestname -match $emscheck)  # --- Verify for EMS Server (NCSCSSV01)
                {

    				Write-Host "[INFO] --- EMS Windows Host Verification Within Execution : ($vmguestnewname)" -ForegroundColor Yellow `n
    				Write-Host "[INFO] --- Preparing Sysprep before Execution : ($vmguestname)" -ForegroundColor Yellow 
    				Write-Host "[INFO] --- Creating Snapshots before Sysprep : ($vmguestname)" -ForegroundColor Yellow 
                    
                    new-snapshot -Name "Before Win10 Sysprep Process" -Description "Perform Snapshot before Sysprep Process only for $($vmguestname)" -VM $($vmguestname) -Server $($vmguesthostip) -Confirm:$false -Verbose

                    $xmlReplaceValues = Get-Content "$($vmunattendlocation).xml" # --- Replace current Prod Key with valid Win 10 Prod Key
                    $BaseXMLTemplate = $xmlReplaceValues.Replace("B6VX8-NP23X-HK48Q-33F79-P38VR", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2")| Set-Content "$($vmunattendlocation).xml" 

                    # --- MiraCastViewer required for Syspre Win10
                    $miraCastViewsrc = "D:\Upload\VMGUESTS\Stage4_Windows_OS-Configuration\MiracastView\"
                    $miraCastViewdst = "C:\Windows\MiracastView\"


                    Write-Host "[INFO] --- Transferring Files to VM Guest : '$($vmguestname)' / '$($vmguestnewname)' $($Destinationlocation)" -ForegroundColor Yellow

				    # --- Transfer sysprep files to ems monitoring server (Windows 10)
				    Copy-VMGuestFile -Source "$($vmunattendlocation).xml" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $emsadmin -GuestPassword $vmguestpwd -Force -Verbose
				    Copy-VMGuestFile -Source "$($vmunattendlocation).ps1" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $emsadmin -GuestPassword $vmguestpwd -Force -Verbose
				    Copy-VMGuestFile -Source "$($enableremoteps)" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $emsadmin -GuestPassword $vmguestpwd -Force -Verbose
				    Copy-VMGuestFile -Source "$($miraCastViewsrc)" -Destination $($miraCastViewdst) -VM $($vmguestname) -LocalToGuest -GuestUser $emsadmin -GuestPassword $vmguestpwd -Force -Verbose

                    Write-Host "[INFO] --- Executing Sysprep in Windows Guest : $($vmguestname) from: $($vmguesthostip)" -ForegroundColor Yellow           
                    Write-Host "[INFO] --- Sysprep File : $($Destinationlocation)$($sysprepps)" -ForegroundColor Yellow `n
                                          
                    # --- Execute the customized sysprep script that was changed above
                    Invoke-VMScript -ScriptText $sysprepps -ScriptType Powershell -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $emsadmin -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose

                    # --- Delay 3 mins for VMGuest to boot up and ready for next stage 
                    Sleep 180

                    # --- Enable RDP
                    # --- Disable Windows Firewall
                    # --- Set VMGuest VHDD Online
                    Write-Host "[INFO] --- Enabling RDP in Windows Guest : '$($vmguestname)' / '$($vmguestnewname)'"  -ForegroundColor Yellow
                    Write-Host "[INFO] --- Disabling Windows Firewall in Windows Guest : '$($vmguestname)' / '$($vmguestnewname)'" -ForegroundColor Yellow `n
                    Invoke-VMScript -ScriptText {

                    Set-location $($Destinationlocation);.\$($enableremoteps);
                    Set-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0 -Verbose; 
                    Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False -Verbose ;
                    Get-Disk | Set-Disk -IsOffline $false -Verbose ;
                    $vmdriveletter = Get-WmiObject win32_volume | where -Property DriveType -eq 5;
                    $vmsetwmi = Set-WmiInstance -InputObject $vmdriveletter -Arguments @{DriveLetter = 'Z:'};

                    } -ScriptType Powershell -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $emsadmin -GuestPassword $vmguestpwd -ToolsWaitSecs 30 -Verbose

                    # --- Apply Windows 10 Power Plan Configurations
                                      
                    Invoke-VMScript -ScriptText "powercfg /change monitor-timeout-ac 0" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($emsadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
                    Invoke-VMScript -ScriptText "powercfg /change monitor-timeout-dc 0" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($emsadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
                    Invoke-VMScript -ScriptText "powercfg /change standby-timeout-ac 0" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($emsadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
                    Invoke-VMScript -ScriptText "powercfg /change standby-timeout-dc 0" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($emsadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
                     
                    # --- Execute the Cyber -ARC route in route table
                                        
                    Invoke-VMScript -ScriptText "Route add -p $($CyberArcIPDc3) mask $($CyberArcSN) $($vmguest.gateway02)" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($emsadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
                    Invoke-VMScript -ScriptText "Route add -p $($CyberArcIPDc4) mask $($CyberArcSN) $($vmguest.gateway02)" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($emsadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose

    				Write-Host "[COMPLETED] --- Sysprep Execution : '$($vmguestname)' / '$($vmguestnewname)'" -ForegroundColor Green `n 

                }

                # ---NON EMS Server and Avoid Sysprep on CA UIM IWMC 2012
                elseif (($vmguestname -notmatch $emscheck) -and ($vmguestname -notmatch $uim2012check))
                {

                    # --- Create Snapshot before sysprep
    				Write-Host "[INFO] --- Windows Host Verification Within Execution : ($vmguestnewname)" -ForegroundColor Yellow `n
    				Write-Host "[INFO] --- Preparing Sysprep before Execution : ($vmguestname)" -ForegroundColor Yellow 
    				Write-Host "[INFO] --- Creating Snapshots before Sysprep : ($vmguestname)" -ForegroundColor Yellow 
                    
                    new-snapshot -Name "Before Win Server 2016 Sysprep Process" -Description "Perform Snapshot before Sysprep Process only for Othe Win 2016" -VM $($vmguestname) -Server $($vmguesthostip) -Confirm:$false -Verbose

                    # --- Transfer sysprep files to vmguest (Windows 2016 Servers)
                    Copy-VMGuestFile -Source "$($vmunattendlocation).xml" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Force -Verbose
                    Copy-VMGuestFile -Source "$($vmunattendlocation).ps1" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Force -Verbose
                    Copy-VMGuestFile -Source "$($enableremoteps)" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Force -Verbose            

                    Write-Host "[INFO] --- Executing Sysprep in Windows Guest : $($vmguestname) from: $($vmguesthostip)" -ForegroundColor Yellow           
                    Write-Host "[INFO] --- Sysprep File : $($Destinationlocation)$($sysprepps)" -ForegroundColor Yellow `n
                             
                    # --- Execute the customized sysprep script that is changed above
                    Invoke-VMScript -ScriptText $sysprepps -ScriptType Powershell -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Verbose -ToolsWaitSecs 30
                    
                    # --- Delay 3 mins for VMGuest to boot up and ready for next stage 
                    Sleep 180
                   
                    # --- Enable RDP
                    # --- Disable Windows Firewall
                    # --- Set VMGuest VHDD Online
                    Write-Host "[INFO] --- Enabling RDP in Windows Guest : '$($vmguestname)' / '$($vmguestnewname)'"  -ForegroundColor Yellow
                    Write-Host "[INFO] --- Disabling Windows Firewall in Windows Guest : '$($vmguestname)' / '$($vmguestnewname)'" -ForegroundColor Yellow `n
                    Invoke-VMScript -ScriptText {

                    Set-location $($Destinationlocation);.\$($enableremoteps);
                    Set-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0 -Verbose; 
                    Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False -Verbose ;
                    Get-Disk | Set-Disk -IsOffline $false -Verbose ;
                    $vmdriveletter = Get-WmiObject win32_volume | where -Property DriveType -eq 5;
                    $vmsetwmi = Set-WmiInstance -InputObject $vmdriveletter -Arguments @{DriveLetter = 'Z:'};
                    
                    } -ScriptType Powershell -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Verbose -ToolsWaitSecs 30

                    Invoke-VMScript -ScriptText "Route add -p $($CyberArcIPDc3) mask $($CyberArcSN) $($vmguest.gateway02)" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
                    Invoke-VMScript -ScriptText "Route add -p $($CyberArcIPDc4) mask $($CyberArcSN) $($vmguest.gateway02)" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
    				Write-Host "[COMPLETED] --- Sysprep Execution : ($vmguestname)" -ForegroundColor Green `n 

                }

                # --- Sysprep on CA UIM IWMC 2012
                elseif ($vmguestcheck2 -match ($uim2012check))
                {

                    # --- Create Snapshot before sysprep
    				Write-Host "[INFO] --- Windows Host Verification Within Execution : ($vmguestnewname)" -ForegroundColor Yellow `n
    				Write-Host "[INFO] --- Preparing Sysprep before Execution : ($vmguestname)" -ForegroundColor Yellow 
    				Write-Host "[INFO] --- Creating Snapshots before Sysprep : ($vmguestname)" -ForegroundColor Yellow 
                    
                    new-snapshot -Name "Before Win Server 2012 Sysprep Process" -Description "Perform Snapshot before Sysprep Process only for UIM IWMC 2012" -VM $($vmguestname) -Server $($vmguesthostip) -Confirm:$false -Verbose
                    
                    # --- Added to replace current Prod Key with valid Win 2012 R2 Prod Key
                    $xmlReplaceValues = Get-Content "$($vmunattendlocation).xml" 
                    $BaseXMLTemplate = $xmlReplaceValues.Replace("B6VX8-NP23X-HK48Q-33F79-P38VR", "D2N9P-3P6X9-2R39C-7RTCD-MDVJX")| Set-Content "$($vmunattendlocation).xml" 

                    Write-Host "[INFO] --- Transferring Files to VM Guest : $($vmguestnewname) $($Destinationlocation)" -ForegroundColor Yellow `n
                    
                    # --- Transfer sysprep files to vmguest (Windows 2012 Servers)
                    Copy-VMGuestFile -Source "$($vmunattendlocation).xml" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Force -Verbose
                    Copy-VMGuestFile -Source "$($vmunattendlocation).ps1" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Force -Verbose
                    Copy-VMGuestFile -Source "$($enableremoteps)" -Destination $($Destinationlocation) -VM $($vmguestname) -LocalToGuest -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Force -Verbose            

                    Write-Host "[INFO] --- Executing Sysprep in Windows Guest : $($vmguestname) from: $($vmguesthostip)" -ForegroundColor Yellow           
                    Write-Host "[INFO] --- Sysprep File : $($Destinationlocation)$($sysprepps)" -ForegroundColor Yellow `n                    
                          
                    # --- Execute the customized sysprep script for CA UIM IWMC 2012 with added Win Server 2012 R2 Prod Key
                    Invoke-VMScript -ScriptText $sysprepps -ScriptType Powershell -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Verbose -ToolsWaitSecs 30
                    
                    # --- Delay 3 mins for VMGuest to boot up and ready for next stage 
                    sleep -Seconds 180

                    # --- Enable RDP
                    # --- Disable Windows Firewall
                    # --- Set VMGuest VHDD Online
                    Write-Host "[INFO] --- Enabling RDP in Windows Guest : '$($vmguestname)' / '$($vmguestnewname)'"  -ForegroundColor Yellow
                    Write-Host "[INFO] --- Disabling Windows Firewall in Windows Guest : '$($vmguestname)' / '$($vmguestnewname)'" -ForegroundColor Yellow `n
                    Invoke-VMScript -ScriptText {

                    Set-location $($Destinationlocation);.\$($enableremoteps);
                    Set-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0 -Verbose; 
                    Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False -Verbose ;
                    Get-Disk | Set-Disk -IsOffline $false -Verbose ;
                    $vmdriveletter = Get-WmiObject win32_volume | where -Property DriveType -eq 5;
                    $vmsetwmi = Set-WmiInstance -InputObject $vmdriveletter -Arguments @{DriveLetter = 'Z:'};
                    
                    } -ScriptType Powershell -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -Verbose -ToolsWaitSecs 30

                    Invoke-VMScript -ScriptText "Route add -p $($CyberArcIPDc3) mask $($CyberArcSN) $($vmguest.gateway02)" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
                    Invoke-VMScript -ScriptText "Route add -p $($CyberArcIPDc4) mask $($CyberArcSN) $($vmguest.gateway02)" -ScriptType Bat -VM $($vmguestname) -Server $($vmguesthostip) -GuestUser $($vmguestadmin) -GuestPassword $vmguestpwd -ToolsWaitSecs 30  -Verbose
    				Write-Host "[COMPLETED] --- Sysprep Execution : '$($vmguestname)' / '$($vmguestnewname)'" -ForegroundColor Green `n 
                }
            
            }
            ELSE # --- Skip all Linux VM from Windows Sysprep
            {

                Write-Host "[INFO] --- Skipping Linux Based OS for SYSPREP" -ForegroundColor Green
            }

        }
        ELSE # --- ESXi Host IPaddress not Tally between Schoolhost and Vmguest csv
        {
            Write-host "[INFO] --- $($vmguesthostip) / $($esxiIp)" -ForegroundColor Yellow
            Write-host "[INFO] --- ESXi Host not Tally!" -ForegroundColor Yellow `n
        }

    }

#endregion ------------ 3. END OF Rename all VM based on namin Convention -------------#
    # --- Disconnect all ViServer Sessions
    write-host "[INFO] --- Disconnecting to ESXiHost: $esxihostIP"  -foreground Yellow
    Disconnect-VIServer $esxihostIP -Force -Confirm:$false -Verbose
    write-host "[INFO] --- Disconnected from ESXiHost Server $esxihostIP" -foreground green `n