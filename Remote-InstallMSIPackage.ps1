[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True)]
    [String]$ComputerName,
    
    [Parameter(Mandatory=$True)]
    [ValidateScript({$_ | Test-Path })]
    [String]$Path
)

Begin {
    function Get-MsiProductInformation {
        # Fork of https://gist.github.com/jstangroome/913062
        [CmdletBinding()]
        param (
            [Parameter(Mandatory=$true)]
            [ValidateScript({$_ | Test-Path -PathType Leaf})]
            [string]
            $Path,

            [Parameter(Mandatory=$true)]
            [string]$Property
        )
        
        function Get-Property ($Object, $PropertyName, [object[]]$ArgumentList) {
            return $Object.GetType().InvokeMember($PropertyName, 'Public, Instance, GetProperty', $null, $Object, $ArgumentList)
        }
    
        function Invoke-Method ($Object, $MethodName, $ArgumentList) {
            return $Object.GetType().InvokeMember($MethodName, 'Public, Instance, InvokeMethod', $null, $Object, $ArgumentList)
        }
    
        $ErrorActionPreference = 'Stop'
        Set-StrictMode -Version Latest
    
        # http://msdn.microsoft.com/en-us/library/aa369432(v=vs.85).aspx
        $msiOpenDatabaseModeReadOnly = 0
        $Installer = New-Object -ComObject WindowsInstaller.Installer
    
        $Database = Invoke-Method $Installer OpenDatabase  @($Path, $msiOpenDatabaseModeReadOnly)
    
        $View = Invoke-Method $Database OpenView  @("SELECT Value FROM Property WHERE Property='$Property'")
    
        Invoke-Method $View Execute | Out-Null
    
        $Record = Invoke-Method $View Fetch
        if ($Record) {
            Write-Output (Get-Property $Record StringData 1)
        }
    
        Invoke-Method $View Close @() | Out-Null
        Remove-Variable -Name Record, View, Database, Installer
    
    }

    
    $MsiData = [PsCustomObject]@{
        FileName            = ($Path -Split '\\')[-1]
        ProductName     = Get-MsiProductInformation -Path $Path -Property ProductName
        ProductVersion  = Get-MsiProductInformation -Path $Path -Property ProductVersion
        Manufacturer    = Get-MsiProductInformation -Path $Path -Property Manufacturer
    }

    # If (Test-Connection -ComputerName $ComputerName <#-TcpPort 5985#>) {
    If (Test-WSMan -ComputerName $ComputerName -ErrorAction Stop) {
        If ($Session = New-PSSession -ComputerName $ComputerName -ErrorAction Stop) {
            $UserName = (Invoke-Command -ScriptBlock {query user} -Session $Session) -replace '\s{2,}',',' | ConvertFrom-Csv | Where-Object SESSIONNAME -eq 'console' | Select-Object -ExpandProperty USERNAME
            Write-Host "Connection established to $ComputerName used by '$UserName'."
        }# Else { Throw "Cant establish PSSession to host." }
    }# Else { Throw "Host unreachable." }
}

Process {
    Invoke-Command -OutVariable ProductData -Session $Session -ArgumentList $MsiData -ScriptBlock { Get-CimInstance -ClassName win32_product -filter ("name = '" + $args.ProductName + "'") } | Out-Null
    
    If ($ProductData) {
        Write-Host ("*** '{0}' is currently installed on the remote host." -f $MsiData.ProductName)
        [PSCustomObject]@{
            "Product Name"      = $MsiData.ProductName
            "Version"           = "{0} -> {1}" -f $ProductData.Version, $MsiData.ProductVersion
            "InstallDate"       = "{0} -> {1}" -f $ProductData.InstallDate, (Get-Date).ToString('yyyyMMdd')
        } | Format-List
    }
    Else {
        [PSCustomObject]@{
            "Product Name"      = $MsiData.ProductName
            "Version"           = $MsiData.ProductVersion
            "InstallDate"       = (Get-Date).ToString('yyyyMMdd')
        } | Format-List
    }

    If ((Read-Host "Should we proceed with the installation? [y/n]") -eq 'y') {
        Copy-Item -ToSession $Session -Path $Path -Destination ("C:\" + $MsiData.FileName)

        # Uninstall
        If ($ProductData.Name) {
            Invoke-Command -OutVariable UninstallStatus -Session $Session -ArgumentList $MsiData -ScriptBlock { Get-CimInstance -ClassName win32_product -filter ("name = '" + $args.ProductName + "'") | Invoke-CimMethod -MethodName Uninstall } | Out-Null
            If ($UninstallStatus.ReturnValue -eq 0) {
                Write-Host "Uninstalled successfully."
            } Else {
                Write-Host ("Uninstallation returned exitcode: {0}" -f $UninstallStatus.ReturnValue)
            }
        }

        # Install
        Invoke-Command -Session $Session -ArgumentList $MsiData -ScriptBlock { Start-Process -Wait -FilePath "msiexec" -ArgumentList @( '/i', ("C:\" + $args.FileName), '/qn') } | Out-Null

        # Query product
        Invoke-Command -OutVariable NewProductData -Session $Session -ArgumentList $MsiData -ScriptBlock { Get-CimInstance -ClassName win32_product -filter ("name = '" + $args.ProductName + "'") | Select-Object Name, Version, InstallDate } | Out-Null
        If ($NewProductData.Name) {
            Write-Host "Installed successfully."
            Invoke-Command -Session $Session -ArgumentList $MsiData -ScriptBlock { Remove-Item ("C:\" + $MsiData.FileName) } | Out-Null
        } Else {
            Write-Host "Installation failed."
        }
    }
}
End {
    Remove-PSSession $Session
}