#Requires -RunAsAdministrator

[CmdletBinding()]
param(
     [Parameter(mandatory = $false, HelpMessage = "Specify where downloaded files and script log should be saved")]
     [string]$DefaultLogPath = "C:\Temp\VirtualTeamsHelper",
 
     [Parameter(mandatory = $false, HelpMessage = "Use this parameter to cleanup downloaded files")]
     [switch]$Cleanup,

     [Parameter(mandatory = $false, HelpMessage = "Use this parameter to only detect where Teams is installed")]
     [switch]$CheckVersions
 )

 

Function GenerateFolder($path) {
    $global:foldPath = $null
    foreach($foldername in $path.split("\")) {
        $global:foldPath += ($foldername+"\")
        if (!(Test-Path $global:foldPath)){
            New-Item -ItemType Directory -Path $global:foldPath
            
        }
    }
}

$TestPath = test-path $DefaultLogPath
if (!($TestPath)){GenerateFolder $DefaultLogPath}


If ($Cleanup -eq $true){

    Start-Transcript -Path "$DefaultLogPath\Teams_Logfile.txt" -Append
    
    Write-Output "Cleanup parameter has been set to $True"
    Write-Output "Removeing Downloaded files"
    Remove-Item -Path "$DefaultLogPath\vc_redist.x64.exe" -Force -ErrorAction Ignore
    Remove-Item -Path "$DefaultLogPath\MsRdcWebRTCSvc_x64.msi" -Force -ErrorAction Ignore
    Remove-Item -Path "$DefaultLogPath\Teams.exe" -Force -ErrorAction Ignore

    Write-Output "Downloaded files has been deleted"
    Write-Output "Ending script"
    Stop-Transcript
    break
}


#Start Transcript
Start-Transcript -Path "$DefaultLogPath\Teams_Logfile.txt" -Append


#Reset parameter
$MachinewideInstaller = $null
$InstalledLocalAppData = $null
$Windows365Environment = $null
$AVDEnvironment = $null


#Detect Virtual Platform
$IsVirtual = ((Get-WmiObject win32_computersystem).model -eq 'VMware Virtual Platform' -or ((Get-WmiObject win32_computersystem).model -eq 'Virtual Machine'))
If ($IsVirtual -eq $true){
    
    #Check if devices is Cloud PC or AVD environment
    $Computername = $env:COMPUTERNAME
    if ($Computername -like "CPC-*"){

        Write-Output "This virtual device is a Windows 365 Cloud PC"
        $Windows365Environment = $true

    }else{

        Write-Output "This virtual device is apart of a Azure virtual Desktop environment"
        $AVDEnvironment = $true
    }


}else {
    
    Write-Output "The PC you are running this script on is not a virtual device"
    Write-Output "This script only works on virtual device, ending script."
    Stop-Transcript
    break

}



#Check if Teams is installed in local app data
#Locate User
$CurrentUserProfile = get-itemproperty "REGISTRY::HKEY_USERS\*\Volatile Environment"
#Write-Output "This Cloud PC belongs to user: $($CurrentUserProfile.USERNAME)"
#Write-Output ""
$Username = $CurrentUserProfile.USERNAME

$exepath = "c:\users\" + $Username + "\AppData\Local\Microsoft\Teams\update.exe"
$deadpath = "c:\users\" + $Username + "\AppData\Local\Microsoft\Teams\.dead"
$currentpath = "c:\users\" + $Username + "\AppData\Local\Microsoft\Teams\current"      
            
        if (((test-path -Path $exepath) -eq $true) -and ((test-path -Path $currentpath) -eq $true) -and ((test-path -Path $deadpath) -eq $false)) {
             Write-Output "Microsoft Teams Is installed in Local Appdata"
             Write-Output ""
             $InstalledLocalAppData = $True

             $TeamsVersion = Get-Content "c:\users\$Username\AppData\Roaming\Microsoft\Teams\settings.json" | ConvertFrom-Json | Select Version, Ring, Environment -ErrorAction Ignore
             Write-Output "Microsoft Teams Version, deployment ring and environment is:"
             Write-Output "$TeamsVersion"
             Write-Output ""
        }


#Determine if Teams is installed in Program Files x86
$path = "C:\Program Files (x86)\Microsoft\Teams"
$exepath = "C:\Program Files (x86)\Microsoft\Teams\update.exe"
$deadpath = "C:\Program Files (x86)\Microsoft\Teams\.dead"
$currentpath = "C:\Program Files (x86)\Microsoft\Teams\current"     
if ((Test-Path -path $path) -eq $true) {
    if (((test-path -Path $exepath) -eq $true) -and ((test-path -Path $currentpath) -eq $true) -and ((test-path -Path $deadpath) -eq $false)) {
        write-Output "Microsoft Teams is installed in the Program Files x86 folder."
        Write-Output ""

        $MachinewideInstaller = $True

       		$TeamsVersion = Get-Content "c:\users\$Username\AppData\Roaming\Microsoft\Teams\settings.json" | ConvertFrom-Json | Select Version, Ring, Environment -ErrorAction Ignore
             Write-Output "Microsoft Teams Version, deployment ring and environment is:"
             Write-Output "$TeamsVersion"
             Write-Output ""
        
    }
    
}



#Determine if Teams is installed in Program Files
$path = "C:\Program Files\Microsoft\Teams"
$exepath = "C:\Program Files)\Microsoft\Teams\update.exe"
$deadpath = "C:\Program Files\Microsoft\Teams\.dead"
$currentpath = "C:\Program Files\Microsoft\Teams\current"     
if ((Test-Path -path $path) -eq $true) {
    if (((test-path -Path $exepath) -eq $true) -and ((test-path -Path $currentpath) -eq $true) -and ((test-path -Path $deadpath) -eq $false)) {
        write-Output "Microsoft Teams is installed in the Program Files folder."
        Write-Output ""

        $MachinewideInstaller = $True

        $TeamsVersion = Get-Content "c:\users\$Username\AppData\Roaming\Microsoft\Teams\settings.json" | ConvertFrom-Json | Select Version, Ring, Environment -ErrorAction Ignore
        Write-Output "Microsoft Teams Version, deployment ring and environment is:"
        Write-Output "$TeamsVersion"
        Write-Output ""
        
    }
    
}



#Detect installed version of WEBRTC
if ((test-path -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\AddIns\WebRTC Redirector\') -eq $true) {

$version = (Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\AddIns\WebRTC Redirector\')
Write-Output "The Installed version of WEBRTC is: $($version.currentversion)"

    $Servicestatus = Get-service -Name 'RDWebRTCSVC' -ErrorAction Ignore
    if($Servicestatus.status -eq "Stopped") {
   
        Write-Output "WEBRTC Service is stopped!"

    }


    
}else {
   Write-Output "Unable to detect version of WEBRTC Version"
            
}

#Detect installed version of Microsoft Visual C++
if ((test-path -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64\') -eq $true) {

    $version = (Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\X64\')
           
    Write-Output "The Installed version of Microsoft Visual C++ is: $($version.Version)"
    Write-Output ""
    
 }else {
    Write-Output "Unable to detect version of Microsoft Visual C++"
    Write-Output ""
}


#Detect if required regkey is present
reg add "HKLM\SOFTWARE\Microsoft\Teams" /v IsWVDEnvironment /t REG_DWORD /d 1 /f
if ((test-path -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Teams\') -eq $true) {

    $version = (Get-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Teams\' -Name "IsWVDEnvironment" -ErrorAction Ignore) 
    If($Version.IsWVDEnvironment -eq 1){

        Write-Output "Required Regkey 'IsWVDEnvironment' under 'HKLM\SOFTWARE\Microsoft\Teams' is set to '1'"
        Write-Output ""

    }else{

        Write-Output "Required Regkey 'IsWVDEnvironment' under 'HKLM\SOFTWARE\Microsoft\Teams' is set not set correctly"
        Write-Output ""

    }
        
}else {
  Write-Output "Unable to check the required regkey"
                
}


If ($CheckVersions -eq $true){
   
    Write-Output "CheckVersions parameter is set to $True"
    Write-Output "Will not download a new Teams version or other relevant updates"
   
    Write-Output "Ending script"
    Stop-Transcript
    break
}


#Downloading latets Microsoft WEBRTC
Write-Output "Downloading latests WEBRTC version..."
invoke-WebRequest -Uri https://aka.ms/msrdcwebrtcsvc/msi -OutFile "$DefaultLogPath\MsRdcWebRTCSvc_x64.msi"
Start-Sleep -Seconds 5
Write-Output "WEBRTC have been downloaded to folder: $DefaultLogPath"
Write-Output ""


#Downloading latets  C++ runtime
Write-Output "Downloading latests C++ Runtime version..."
invoke-WebRequest -Uri https://aka.ms/vs/16/release/vc_redist.x64.exe -OutFile "$DefaultLogPath\vc_redist.x64.exe"
Start-Sleep -Seconds 5
Write-Output "C++ runtime have been downloaded to folder: $DefaultLogPath"
Write-Output ""

#Installing latests C++ Runtime version
Write-Output "Installing latests C++ Runtime version..."
Write-Output ""
Start-Process -FilePath "$DefaultLogPath\vc_redist.x64.exe" -ArgumentList '/q', '/norestart'
Start-Sleep -s 10

#Installing latests WebRTC version
Write-Output "Installing latests WEBRTC version..."
Write-Output ""
msiexec /i $DefaultLogPath\MsRdcWebRTCSvc_x64.msi /q /n
Start-Sleep -s 10
$Servicestatus = Get-service -Name 'RDWebRTCSVC' -ErrorAction Ignore
    if($Servicestatus.status -eq "Stopped") {
   
        Write-Output "WEBRTC Service is stopped"
        Write-Output "Starting service..."
        Start-Service -Name RDWebRTCSVC
        Write-Output ""
    }


#Setting required Regkey
Write-Output "Adding required Regkey 'IsWVDEnvironment' under 'HKLM\SOFTWARE\Microsoft\Teams'"
Write-Output ""
reg add "HKLM\SOFTWARE\Microsoft\Teams" /v IsWVDEnvironment /t REG_DWORD /d 1 /f
    



If ($Windows365Environment){

        #IF Cloud PC run this part
        #Downloading latets Microsoft Teams to the Cloud PC
        Write-Output "Downloading latests microsoft Teams version..."
        invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?LinkID=2187327&clcid=0x409&culture=en-us&country=US" -OutFile "$DefaultLogPath\Teams.exe"
        Start-Sleep -Seconds 10
        Write-Output "Microsoft Teams have been downloaded to folder: $DefaultLogPath"
        Write-Output ""



        If ($MachinewideInstaller -eq $true){

        #Detect Machine Wide Teams
        $MachineWide = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "Teams Machine-Wide Installer"}
            if ($MachineWide.Name -eq "Teams Machine-Wide Installer"){
            write-output "Stopping Microsoft Team Processes"
            #Stop Microsoft Teams Processes
            if ((get-process | Where-Object ProcessName -eq "Teams").Count -gt 0) { Stop-Process -name teams -force }

            #Uninstall microsoft teams
            $MachineWide.Uninstall()
            start-sleep -Seconds 5

                #Check if Microsoft Teams is uninstalled
                $MachineWide = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "Teams Machine-Wide Installer"}
                if ($MachineWide.Name -eq $null){
                Write-Output "Teams has been uninstalled"
                Write-Output "Run the new installer as the user to install the correct Microsoft Teams version"
                Write-Output "The installer is loacted at $DefaultLogPath"
                Write-Output ""
            
            
                }



            }else{
            Write-Output "Unable to find the uninstallation of Microsoft Teams."
            Write-Output ""

            }

        

        }

        If ($InstalledLocalAppData -eq $true){
        Write-Output "Microsoft Teams is installed correctly on this Cloud PC Device."
        Write-Output "The latests Microsoft teams is downloaded to the folder $DefaultLogPath"
        Write-Output "Try to update the teams client by run the installer as the user or select 'Update inside Microsoft Teams'"
        Write-Output ""

        }

        If ($MachinewideInstaller -eq $null -and $InstalledLocalAppData -eq $null){

        Write-Output "There was not detected any Microsoft Teams on this Cloud PC."
        Write-Output "Install the latets Microsoft team as the user"
        Write-Output "The Teams installer is located at $DefaultLogPath"
        Write-Output ""


        } 

}elseif ($AVDEnvironment) {
    #IF AVD run this part
    #Downloading latets Microsoft Teams to the AVD Machine
    Write-Output "Downloading latests microsoft Teams version..."
    invoke-WebRequest -Uri "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true" -OutFile "$DefaultLogPath\Teams_MachineWide.msi"
    Start-Sleep -Seconds 10
    Write-Output "Microsoft Teams have been downloaded to folder: $DefaultLogPath"
    Write-Output ""


    If ($InstalledLocalAppData -eq $true){
        Installeret forkett! gør noget!

    }


    If ($MachinewideInstaller -eq $true){
        Write-Output "Microsoft Teams is installed correctly on this AVD machine."
        Write-Output "Installing the latest Microsoft Teams version on this AVD Machine"
        msiexec /i $DefaultLogPath\Teams_MachineWide.msi /l*v $DefaultLogPath\Teams_MachineWide_installLog.txt ALLUSERS=1 ALLUSER=1 /q
        Start-Sleep -Seconds 10
        
        Write-Output ""

    }


    If ($MachinewideInstaller -eq $null -and $InstalledLocalAppData -eq $null){

        Write-Output "There was not detected any Microsoft Teams on this Cloud PC."
        Write-Output "Installing the latest Microsoft Teams version on this AVD Machine"
        msiexec /i $DefaultLogPath\Teams_MachineWide.msi /l*v $DefaultLogPath\Teams_MachineWide_installLog.txt ALLUSERS=1 ALLUSER=1 /q
        Start-Sleep -Seconds 10
        
        Write-Output ""


    }



}else{

    Write-Output "Unable to detect environment"
    Write-Output "Ending script"

}




Stop-Transcript
