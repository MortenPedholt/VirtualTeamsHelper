# Disclaimer
Please use this script with care. Read the Examples before running the script.
It will uninstall Teams on the virtual client if executed without any parameters.

# Description of VirtualTeamsHelper
With VirtualTeamsHelper you can use to get information about the Microsoft Teams installation on a Virtual Device such as Windows 365 Cloud PC.
It will pull information about WEBRTC, Visual C++, version of Teams, Deployment ring of Teams and more.

# Examples of running the script

- .\VirtualTeamsHelper.ps1 : Detect versions, Download and install the latests WEBRTC and Visual C++. Uninstall Teams if it's installed in the wrong location. Download the latest Teams version.
- .\VirtualTeamsHelper.ps1 -CheckVersions : Use this parameter to only detect where Teams is installed and what versions the required components are running
- .\VirtualTeamsHelper.ps1 -DefaultLogPath : Specify where downloaded files and script log should be saved. Default logpath is: C:\Temp\VirtualTeamsHelper
- .\VirtualTeamsHelper.ps1 -Cleanup : Use this parameter to cleanup downloaded files.



# Requirements

The script must be executed as Administrator
