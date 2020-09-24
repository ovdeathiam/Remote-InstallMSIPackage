# Remote-InstallMSIPackage
## How does it work?
The script uses WSMan to establiesh a PSSession, copy the *Microsoft Installer Package* onto the target machine and perform the installation.

## Prerequisites
You need to be running Windows as the script relies on **WindowsInstaller.Installer** ComObject to read the source *.msi package data, and query.exe to check who is currently using the target computer.
* Running Microsoft Windows
* Windows Powershell 5.1 or Powershell Core
* You need to set your ExecutionPolicy to allow executing 3rd party scripts
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy Bypass
  ```
* Have sufficient rights to run **PSSession** and **msiexec** on the target machine.

### Usage
```powershell
> .\Remote-InstallMSIPackage.ps1 -ComputerName <Hostname> -Path <Path to *.msi>
Connection established to <Hostname> used by 'mkowalsky'.

Product Name : 7-Zip 19.00
Version      : 19.00.00.0
InstallDate  : 20200924

Should we proceed with the installation? [y/n]: y
Installed successfully.

```

### Examples
#### Installation
![alt text](https://github.com/ovdeathiam/Remote-InstallMSIPackage/raw/master/img/example-install.png "Installation example")
#### Reinstallation
![alt text](https://github.com/ovdeathiam/Remote-InstallMSIPackage/raw/master/img/example-reinstall.png "Reinstallation example")
