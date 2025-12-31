
#Use this script for testing Umango issues

##set security when needed
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
#Set-ExecutionPolicy -ExecutionPolicy AllSigned
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine

## start searching for info

## Make directory c:\Umango
$DateString = (Get-Date).ToString('dd-MMM-yyyy') # Get the Current Date Formatted
$FolderPath = "c:\Umango\$DateString" # Frame the Folder Name

# Check Folder Exists
If (Test-Path -Path $FolderPath)
 {
    #$Folder = Get-Item -Path $FolderPath
    #Write-Host "Folder already exists." -f Yellow
 } Else
  {
    #Create a New Folder  
    $Folder = New-Item -ItemType Directory -Path $FolderPath
    New-Item -Path $FolderPath\Info.txt -ItemType File -Verbose
  }

# Check if service is running
$Status = Get-Service -DisplayName "*Umango*"
    if ($Status.Status -eq "Running")
     {
        $pcstatus = "Service is running"
     } else
      {
    $pcstatus = "Service is not running"
      }

# Check firewall
$FPD = if ((Get-NetFirewallProfile -Name Domain).enabled) {'Active'} else {'Disabled'}
$FPP = if ((Get-NetFirewallProfile -Name Public).enabled) {'Active'} else {'Disabled'}
$FPPR = if ((Get-NetFirewallProfile -Name Private).enabled) {'Active'} else {'Disabled'}

## Basic GUI
Add-Type -assembly System.Windows.Forms
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='ITS BeNeLux | Umango Pre-Check | V1.05'
$main_form.Width = 620
$main_form.Height = 700
$main_form.minimumSize = New-Object System.Drawing.Size(660,800) 
$main_form.maximumSize = New-Object System.Drawing.Size(660,800) 
$main_form.AutoSize = $false

#Check If Visual C++ is installed
$VC = Get-ItemPropertyValue -LiteralPath 'HKLM:SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\X64' -Name Version
Add-Content -Path $FolderPath\Info.txt -Value "Umango info V1.05"
Add-Content -Path $FolderPath\Info.txt -Value "Visual C++ - "
Add-Content -Path $FolderPath\Info.txt -Value $VC

# --- Check .NET Framework (registry) ---
$release = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release

$frameworkVersions = @{
    378389 = "4.5"
    378675 = "4.5.1"
    378758 = "4.5.1"
    379893 = "4.5.2"
    393295 = "4.6"
    393297 = "4.6"
    394254 = "4.6.1"
    394271 = "4.6.1"
    394802 = "4.6.2"
    394806 = "4.6.2"
    460798 = "4.7"
    460805 = "4.7"
    461308 = "4.7.1"
    461310 = "4.7.1"
    461808 = "4.7.2"
    461814 = "4.7.2"
    528040 = "4.8"
    528049 = "4.8"
    533320 = "4.8.1"
}

if ($release -and $frameworkVersions.ContainsKey($release)) {
    $framework = $frameworkVersions[$release]
} elseif ($release) {
    $framework = "Unknown Framework (Release=$release)"
} else {
    $framework = "No .NET Framework 4.x detected"
}

# --- Check .NET (Core/5+) via CLI ---
try {
    $runtimes = & dotnet --list-runtimes | Select-String "Microsoft.NETCore.App"
    # Extract only the version column (second token)
    $coreVersions = $runtimes | ForEach-Object { ($_ -split '\s+')[1] }
} catch {
    $coreVersions = @("None")
}

# --- Output clean values ---

$core =  $($coreVersions -join ', ')

#Check If SQLexpress is installed
#Check If SQLserver is installed
# Initialize variables
$mssqlserverStatus  = "Not Found"
$mssqlserverVersion = "Not Found"
$sqlexpressStatus   = "Not Found"
$sqlexpressVersion  = "Not Found"

# Function to check service and get version
function Get-SqlServiceInfo($svcName) {
    $service = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($null -ne $service -and $service.Status -eq "Running") {
        $status = "Running"
        try {
            $instance = if ($svcName -eq "MSSQLSERVER") { "." } else { ".\$svcName" }
            $connectionString = "Server=$instance;Database=master;Integrated Security=True;"

            $connection = New-Object System.Data.SqlClient.SqlConnection $connectionString
            $connection.Open()

            $command = $connection.CreateCommand()
            $command.CommandText = "SELECT SERVERPROPERTY('ProductVersion')"
            $version = $command.ExecuteScalar()

            $connection.Close()
        }
        catch {
            $version = "Not Found"
        }
    }
    elseif ($null -ne $service) {
        $status  = "Installed but not running"
        $version = "Not Found"
    }
    else {
        $status  = "Not Found"
        $version = "Not Found"
    }

    return @{ Status = $status; Version = $version }
}

# Check MSSQLSERVER
$mssqlInfo = Get-SqlServiceInfo "MSSQLSERVER"
$mssqlserverStatus  = $mssqlInfo.Status
$mssqlserverVersion = $mssqlInfo.Version

# Check SQLEXPRESS
$sqlexpressInfo = Get-SqlServiceInfo "SQLEXPRESS"
$sqlexpressStatus  = $sqlexpressInfo.Status
$sqlexpressVersion = $sqlexpressInfo.Version

# Output results
Write-Output "MSSQLSERVER Status : $mssqlserverStatus"
Write-Output "MSSQLSERVER Version: $mssqlserverVersion"
Write-Output "SQLEXPRESS Status  : $sqlexpressStatus"
Write-Output "SQLEXPRESS Version : $sqlexpressVersion"

Add-Content -Path $FolderPath\Info.txt -Value "SQLserver - "
Add-Content -Path $FolderPath\Info.txt -Value $mssqlserverStatus
Add-Content -Path $FolderPath\Info.txt -Value "SQLserver version - "
Add-Content -Path $FolderPath\Info.txt -Value $mssqlserverVersion
#
Add-Content -Path $FolderPath\Info.txt -Value "SQLexpress - "
Add-Content -Path $FolderPath\Info.txt -Value $sqlexpressStatus
Add-Content -Path $FolderPath\Info.txt -Value "SQLexpress version - "
Add-Content -Path $FolderPath\Info.txt -Value $sqlexpressVersion


#Check Computer name
$DeviceName = "$env:COMPUTERNAME"
Add-Content -Path $FolderPath\Info.txt -Value "DeviceName - "
Add-Content -Path $FolderPath\Info.txt -Value $DeviceName
if ($DeviceName.Length -gt 15) {
    Write-Host "Warning: Computername is longer than 15 characters!" -ForegroundColor Red
    Write-Host " "
}

#Check Windows version
$OSversion = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Caption, Version, BuildNumber
Add-Content -Path $FolderPath\Info.txt -Value "Windows OS - "
Add-Content -Path $FolderPath\Info.txt -Value $OSversion
$OSversion2 = (Get-WmiObject Win32_OperatingSystem).caption


# Run dsregcmd and capture AzureJoined
$dsreg = dsregcmd.exe /status

# Check for AzureADJoined status and store in variable
if ($dsreg | Select-String "AzureADJoined\s*:\s*YES") {
    $AzureJoined = "Yes"
} else {
    $AzureJoined = "No"
}

# Try to get UPN using whoami
$upn = whoami /upn 2>$null

# Output result or "none" if empty
if ([string]::IsNullOrWhiteSpace($upn)) {
    $upn = "None"
} else {
    $upn
}

# Run dsregcmd and capture MDM
$dsreg = dsregcmd.exe /status

# Intune Enrollment status
$IntuneEnrolled = if ($dsreg | Select-String "MDMEnrollment\s*:\s*YES") { "Yes" } else { "No" }

# Tenant Name ophalen
# $tenantLine = $dsreg | Select-String "TenantName\s*:\s*(.+)"
# $TenantName = if ($tenantLine) { $tenantLine.Matches[0].Groups[1].Value.Trim() } else { "unknown" }

# Get the network adapter that is currently connected
$adapter = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null }

if ($null -eq $adapter) {
    Write-Host "No active network adapter found." -ForegroundColor Yellow
} else {
    $dhcpEnabled = $adapter.IPv4Address[0].PrefixOrigin -eq 'Dhcp'

    if ($dhcpEnabled) {    
    Write-Host "❗ DHCP is been used for this connected networkcard!" -ForegroundColor Red
    $dhcp = "DCHP is active "

    } else {
        Write-Host "Static IP is configured for the connected network." -ForegroundColor Green
    }
}

# Add Dialog

$lbl_line34 = New-Object System.Windows.Forms.label
$lbl_line34.Text = "Networkcard"
$lbl_line34.Location  = New-Object System.Drawing.Point(50,10)
$lbl_line34.AutoSize = $true
$main_form.Controls.Add($lbl_line34)
$lbl_line34 = New-Object System.Windows.Forms.label
$lbl_line34.Text = "$dhcp"
$lbl_line34.Location  = New-Object System.Drawing.Point(400,10)
$lbl_line34.AutoSize = $true
$main_form.Controls.Add($lbl_line34)
#
$lbl_line35 = New-Object System.Windows.Forms.label
$lbl_line35.Text = "DeviceName"
$lbl_line35.Location  = New-Object System.Drawing.Point(50,30)
$lbl_line35.AutoSize = $true
$main_form.Controls.Add($lbl_line35)
$lbl_line36 = New-Object System.Windows.Forms.label
$lbl_line36.Text = "$DeviceName"
$lbl_line36.Location  = New-Object System.Drawing.Point(400,30)
$lbl_line36.AutoSize = $true
#
$main_form.Controls.Add($lbl_line36)
$lbl_line23 = New-Object System.Windows.Forms.label
$lbl_line23.Text = "Window version"
$lbl_line23.Location  = New-Object System.Drawing.Point(50,50)
$lbl_line23.AutoSize = $true
$main_form.Controls.Add($lbl_line23)
$lbl_line24 = New-Object System.Windows.Forms.label
$lbl_line24.Text = $OSversion2
$lbl_line24.Location  = New-Object System.Drawing.Point(400,50)
$lbl_line24.AutoSize = $true
$main_form.Controls.Add($lbl_line24)
#
$lbl_line28 = New-Object System.Windows.Forms.label
$lbl_line28.Text = "Device is Azure Joined"
$lbl_line28.Location  = New-Object System.Drawing.Point(50,70)
$lbl_line28.AutoSize = $true
$main_form.Controls.Add($lbl_line28)
$lbl_line30 = New-Object System.Windows.Forms.label
$lbl_line30.Text = "$AzureJoined"
$lbl_line30.Location  = New-Object System.Drawing.Point(400,70)
$lbl_line30.AutoSize = $true
$main_form.Controls.Add($lbl_line30)
#
$lbl_line29 = New-Object System.Windows.Forms.label
$lbl_line29.Text = "User /upn"
$lbl_line29.Location  = New-Object System.Drawing.Point(50,90)
$lbl_line29.AutoSize = $true
$main_form.Controls.Add($lbl_line29)
$lbl_line30 = New-Object System.Windows.Forms.label
$lbl_line30.Text = "$upn"
$lbl_line30.Location  = New-Object System.Drawing.Point(400,90)
$lbl_line30.AutoSize = $true
$main_form.Controls.Add($lbl_line30)
#
$lbl_line31 = New-Object System.Windows.Forms.label
$lbl_line31.Text = "MDMEnrollement"
$lbl_line31.Location  = New-Object System.Drawing.Point(50,110)
$lbl_line31.AutoSize = $true
$main_form.Controls.Add($lbl_line31)
$lbl_line33 = New-Object System.Windows.Forms.label
$lbl_line33.Text = "$IntuneEnrolled"
$lbl_line33.Location  = New-Object System.Drawing.Point(400,110)
$lbl_line33.AutoSize = $true
$main_form.Controls.Add($lbl_line33)
#
$lbl_line33 = New-Object System.Windows.Forms.label
$lbl_line33.Text = "-"
$lbl_line33.Location  = New-Object System.Drawing.Point(50,130)
$lbl_line33.AutoSize = $true
$main_form.Controls.Add($lbl_line33)
$lbl_line33 = New-Object System.Windows.Forms.label
$lbl_line33.Text = "-"
$lbl_line33.Location  = New-Object System.Drawing.Point(400,130)
$lbl_line33.AutoSize = $true
$main_form.Controls.Add($lbl_line33)
#
$lbl_line13 = New-Object System.Windows.Forms.label
$lbl_line13.Text = "Visual C++ version"
$lbl_line13.Location  = New-Object System.Drawing.Point(50,150)
$lbl_line13.AutoSize = $true
$main_form.Controls.Add($lbl_line13)
$lbl_line16 = New-Object System.Windows.Forms.label
$lbl_line16.Text = $VC 
$lbl_line16.Location  = New-Object System.Drawing.Point(400,150)
$lbl_line16.AutoSize = $true
$main_form.Controls.Add($lbl_line16)
#
$lbl_line21 = New-Object System.Windows.Forms.label
$lbl_line21.Text = "Umango service"
$lbl_line21.Location  = New-Object System.Drawing.Point(50,170)
$lbl_line21.AutoSize = $true
$main_form.Controls.Add($lbl_line21)
$lbl_line16 = New-Object System.Windows.Forms.label
$lbl_line16.Text = $pcstatus
$lbl_line16.Location  = New-Object System.Drawing.Point(400,170)
$lbl_line16.AutoSize = $true
$main_form.Controls.Add($lbl_line16)
#
$lbl_line18 = New-Object System.Windows.Forms.label
$lbl_line18.Text = "SQL express"
$lbl_line18.Location  = New-Object System.Drawing.Point(50,190)
$lbl_line18.AutoSize = $true
$main_form.Controls.Add($lbl_line18)
$lbl_line19 = New-Object System.Windows.Forms.label
$lbl_line19.Text = $sqlexpressStatus
$lbl_line19.Location  = New-Object System.Drawing.Point(400,190)
$lbl_line19.AutoSize = $true
$main_form.Controls.Add($lbl_line19)
#
$lbl_line20 = New-Object System.Windows.Forms.label
$lbl_line20.Text = "SQL server"
$lbl_line20.Location  = New-Object System.Drawing.Point(50,210)
$lbl_line20.AutoSize = $true
$main_form.Controls.Add($lbl_line20)
$lbl_line20 = New-Object System.Windows.Forms.label
$lbl_line20.Text = $mssqlserverStatus
$lbl_line20.Location  = New-Object System.Drawing.Point(400,210)
$lbl_line20.AutoSize = $true
$main_form.Controls.Add($lbl_line20)
#
$lbl_line24 = New-Object System.Windows.Forms.label
$lbl_line24.Text = ".Net Framework"
$lbl_line24.Location  = New-Object System.Drawing.Point(50,230)
$lbl_line24.AutoSize = $true
$main_form.Controls.Add($lbl_line24)
$lbl_line25 = New-Object System.Windows.Forms.label
$lbl_line25.Text = "$framework"
$lbl_line25.Location  = New-Object System.Drawing.Point(400,230)
$lbl_line25.AutoSize = $true
$main_form.Controls.Add($lbl_line25)
#
$lbl_line26 = New-Object System.Windows.Forms.label
$lbl_line26.Text = ".Net version('s)"
$lbl_line26.Location  = New-Object System.Drawing.Point(50,250)
$lbl_line26.AutoSize = $true
$main_form.Controls.Add($lbl_line26)
$lbl_line27 = New-Object System.Windows.Forms.label
$lbl_line27.Text = "$core"
$lbl_line27.Location  = New-Object System.Drawing.Point(400,250)
$lbl_line27.AutoSize = $true
$main_form.Controls.Add($lbl_line27)
#
$lbl_line14 = New-Object System.Windows.Forms.label
$lbl_line14.Text = "Create Custom view in Windows Eventviewer"
$lbl_line14.Location  = New-Object System.Drawing.Point(50,275)
$lbl_line14.AutoSize = $true
$main_form.Controls.Add($lbl_line14)
#
$main_form.Controls.Add($lbl_line18)
$lbl_line1 = New-Object System.Windows.Forms.label
$lbl_line1.Text = "Enable Server log"
$lbl_line1.Location  = New-Object System.Drawing.Point(50,300)
$lbl_line1.AutoSize = $true
$main_form.Controls.Add($lbl_line1)
$lbl1 = New-Object System.Windows.Forms.label
$lbl1.Text = "Windows Firewall"
$lbl1.Location  = New-Object System.Drawing.Point(50,330)
$lbl1.AutoSize = $true
$main_form.Controls.Add($lbl1)
$lbl2 = New-Object System.Windows.Forms.label
$lbl2.Text = " - Public"
$lbl2.Location  = New-Object System.Drawing.Point(50,360)
$lbl2.AutoSize = $true
$main_form.Controls.Add($lbl2)
$lbl3 = New-Object System.Windows.Forms.label
$lbl3.Text = " - Private"
$lbl3.Location  = New-Object System.Drawing.Point(50,380)
$lbl3.AutoSize = $true
$main_form.Controls.Add($lbl3)
$lbl4 = New-Object System.Windows.Forms.label
$lbl4.Text = "- Domain"
$lbl4.Location  = New-Object System.Drawing.Point(50,400)
$lbl4.AutoSize = $true
$main_form.Controls.Add($lbl4)
$lbl12 = New-Object System.Windows.Forms.label
$lbl12.Text = "Users / Group"
$lbl12.Location  = New-Object System.Drawing.Point(50,420)
$lbl12.AutoSize = $true
$main_form.Controls.Add($lbl12)


$lbl_line15 = New-Object System.Windows.Forms.label
$lbl_line15.Text = "Show running processes"
$lbl_line15.Location  = New-Object System.Drawing.Point(50,520)
$lbl_line15.AutoSize = $true
$main_form.Controls.Add($lbl_line15)
$lbl_16 = New-Object System.Windows.Forms.label
$lbl_16.Text = "Test open ports to Umango"
$lbl_16.Location  = New-Object System.Drawing.Point(50,550)
$lbl_16.AutoSize = $true
$main_form.Controls.Add($lbl_16)
$lbl_17 = New-Object System.Windows.Forms.label
$lbl_17.Text = "Test open ports to MFD or localhost"
$lbl_17.Location  = New-Object System.Drawing.Point(50,580)
$lbl_17.AutoSize = $true
$main_form.Controls.Add($lbl_17)

#
## section Cleanup
$lbl_line5 = New-Object System.Windows.Forms.label
$lbl_line5.Text = "-- Clean-up"
$lbl_line5.Location  = New-Object System.Drawing.Point(50,660)
$lbl_line5.AutoSize = $true
$main_form.Controls.Add($lbl_line5)
#add buttons

# Button Enable printservice in eventviewer
$Button_ena = New-Object System.Windows.Forms.Button
$Button_ena.Location = New-Object System.Drawing.Size(400,270)
$Button_ena.Size = New-Object System.Drawing.Size(120,26)
$Button_ena.Text = "Enable"
$main_form.Controls.Add($Button_ena)

# Button Set registration key for debug log
$Button_setkey = New-Object System.Windows.Forms.Button
$Button_setkey.Location = New-Object System.Drawing.Size(400,297)
$Button_setkey.Size = New-Object System.Drawing.Size(120,26)
$Button_setkey.Text = "Set Key"
$main_form.Controls.Add($Button_setkey)

# Button Show Firewall Ruls
$Button_fr = New-Object System.Windows.Forms.Button
$Button_fr.Location = New-Object System.Drawing.Size(400,327)
$Button_fr.Size = New-Object System.Drawing.Size(120,26)
$Button_fr.Text = "Get Rules"
$main_form.Controls.Add($Button_fr)

# Button Show Users & Groups
$Button_ug = New-Object System.Windows.Forms.Button
$Button_ug.Location = New-Object System.Drawing.Size(400,417)
$Button_ug.Size = New-Object System.Drawing.Size(120,26)
$Button_ug.Text = "Get User('s)"
$main_form.Controls.Add($Button_ug)

# Button Open explorer
$Button_ep = New-Object System.Windows.Forms.Button
$Button_ep.Location = New-Object System.Drawing.Size(400,657)
$Button_ep.Size = New-Object System.Drawing.Size(120,26)
$Button_ep.Text = "Show zip file"
$main_form.Controls.Add($Button_ep)

# Button Clear logdirectory
$Button_del = New-Object System.Windows.Forms.Button
$Button_del.Location = New-Object System.Drawing.Size(400,687)
$Button_del.Size = New-Object System.Drawing.Size(120,26)
$Button_del.Text = "Delete Logs"
$main_form.Controls.Add($Button_del)

#Button open new screen and show running process
$Button_rp = New-Object System.Windows.Forms.Button
$Button_rp.Location = New-Object System.Drawing.Size(400,517)
$Button_rp.Size = New-Object System.Drawing.Size(120,26)
$Button_rp.Text = "Show"
$main_form.Controls.Add($Button_rp)

#Button open new screen and show open ports to papercut
$Button_op = New-Object System.Windows.Forms.Button
$Button_op.Location = New-Object System.Drawing.Size(400,547)
$Button_op.Size = New-Object System.Drawing.Size(120,26)
$Button_op.Text = "Show"
$main_form.Controls.Add($Button_op)

#Button open new screen and show open ports to MFD
$Button_opm = New-Object System.Windows.Forms.Button
$Button_opm.Location = New-Object System.Drawing.Size(400,577)
$Button_opm.Size = New-Object System.Drawing.Size(120,26)
$Button_opm.Text = "Show"
$main_form.Controls.Add($Button_opm)

# Actie

$lbl9 = New-Object System.Windows.Forms.label
$lbl9.Text = $FPP
$lbl9.Location  = New-Object System.Drawing.Point(400,360)
$lbl9.AutoSize = $true
$main_form.Controls.Add($lbl9)
$lbl10 = New-Object System.Windows.Forms.label
$lbl10.Text = $FPPR
$lbl10.Location  = New-Object System.Drawing.Point(400,380)
$lbl10.AutoSize = $true
$main_form.Controls.Add($lbl10)
$lbl11 = New-Object System.Windows.Forms.label
$lbl11.Text = $FPD
$lbl11.Location  = New-Object System.Drawing.Point(400,400)
$lbl11.AutoSize = $true
$main_form.Controls.Add($lbl11)


# set register key
$Button_setkey.Add_Click({
# Define the registry path
$regPath = "HKLM:\SOFTWARE\Umango"

# Ensure the key exists
if (-not (Test-Path $regPath)) {
    New-Item -Path "HKLM:\SOFTWARE" -Name "Umango" -Force | Out-Null
}
Add-Content -Path $FolderPath\Info.txt -Value "Windows registry adjustment"
Add-Content -Path $FolderPath\Info.txt -Value "HKLM:\SOFTWARE\Umango is set"
})

# Show all users
$Button_ug.Add_Click({
    $Date = (Get-Date).ToString('dd-MMM-yyyy')
    $path = "c:\Umango\$Date\user.txt"
    $rawscript = net localgroup Administrators | Out-String -Width 600 | Set-Content -Path $path
    Start-Process powershell.exe -ArgumentList "-NoProfile -Command $rawscript"
})

# Show all firewall rules
$Button_fr.Add_Click({
    $Date = (Get-Date).ToString('dd-MMM-yyyy')
    $path = "c:\Umango\$Date\firewall.txt"
    $rawscript = Get-NetFirewallRule  | Out-String -Width 600 | Set-Content -Path $path
    Start-Process powershell.exe -ArgumentList "-NoProfile -Command $rawscript"
})

# Remove all info
$Button_del.Add_Click
    ({ 
        if (Test-Path 'c:\Umango')
        {
            Try {
                Remove-Item -Path 'c:\Umango' -Recurse -Force -ErrorAction Stop
                Write-Host "Folder removed succesfully."
                }
            catch {
                Write-Host "Could not remove folder: $($_.Exception.Message)"
                  }    
     }
    })
    #Get-ChildItem -Path 'c:\umango' -Force

# Enable eventviewer
$Button_ena.Add_Click({ 
$viewsPath = "$env:ProgramData\Microsoft\Event Viewer\Views"
if (-not (Test-Path $viewsPath)) {
    New-Item -ItemType Directory -Force -Path $viewsPath | Out-Null
}

#$guid = [guid]::NewGuid().ToString()
$customViewXml = @"
<ViewerConfig>
  <QueryConfig>
    <QueryParams>
      <Simple>
        <Channel>Application</Channel>
        <RelativeTimeInfo>0</RelativeTimeInfo>
        <Level>1,2,3,4,0,5</Level>
        <BySource>False</BySource>
      </Simple>
    </QueryParams>
    <QueryNode>
      <Name>Umango-test</Name>
      <QueryList>
        <Query Id="0">
          <Select Path="Application">*[System[Provider[@Name='Umango']]]</Select>
        </Query>
      </QueryList>
    </QueryNode>
  </QueryConfig>
</ViewerConfig>

"@

$viewFile = Join-Path $viewsPath "Umango.xml"
$customViewXml | Out-File -FilePath $viewFile -Encoding utf8
Write-Output "Custom Event Viewer view 'Umango' created at $viewFile"
})

## Open Windows explorer and make zip file
$Button_ep.Add_Click({ 
    wevtutil epl Application $FolderPath\Umango.evtx /ow:true
    Compress-Archive -Path C:\Umango -DestinationPath C:\Umango\Umangologs.zip -Force
    Invoke-Item C:\Umango\ 2>$null
})


## open new window with running processes
$Button_rp.Add_Click({
$Date = (Get-Date).ToString('dd-MMM-yyyy')
    $path = "c:\Umango\$Date\running-process.txt"
$rawScript = @"
`$connections = Get-NetTCPConnection | Select-Object LocalPort, OwningProcess
`$processTable = @()
foreach (`$proc in Get-Process) {
    `$procId = `$proc.Id
    `$ports = (`$connections | Where-Object { `$_.OwningProcess -eq `$procId }).LocalPort
    if (`$ports.Count -gt 0) {
        `$obj = [PSCustomObject]@{
            Id = `$procId
            ProcessName = `$proc.ProcessName
            DisplayName = `$proc.Description
            Ports = `$ports -join ', '
        }
        `$processTable += `$obj
    }
}
`$processTable | Out-String -Width 600 | Set-Content $path
`$processTable | Out-String -Width 500
Read-Host 'Press Enter to exit'
"@

# Encode the cleaned-up script
$bytes = [System.Text.Encoding]::Unicode.GetBytes($rawScript)
$encodedCommand = [Convert]::ToBase64String($bytes)

# Launch it in a new PowerShell window

Start-Process powershell.exe -ArgumentList "-NoExit", "-EncodedCommand $encodedCommand"
})

# Open new window with port check to Umango

$Button_op.Add_Click({
    $scriptContent = @'
   $logfilepath = "C:\umango\Open-ports-Papercut.txt"

# Create the directory if it doesn't exist
if (-not (Test-Path -Path (Split-Path $logfilepath))) {
    New-Item -Path (Split-Path $logfilepath) -ItemType Directory -Force
}

#Get-Date).ToString() | Out-File -FilePath $logfilepath -Encoding UTF8 -Append

    # Port tests
        $hosts = @(
        "www.umango.com",
        "www.abbyy.com",
        "login.microsoftonline.com",
        "graph.microsoft.com"
         )

        $maxAttempts = 2
        $delaySeconds = 3

        foreach ($targetHost in $hosts)
        {
            Write-Host "`nTesting $targetHost..." -ForegroundColor Cyan
            $success = $false

            for ($i = 1; $i -le $maxAttempts; $i++)
            {
                $result = Test-NetConnection $targetHost -Port 443

                    if ($result.TcpTestSucceeded)
                    {
                    Write-Host "✔ SUCCESS: Port 443 open on $targetHost (attempt $i)" -ForegroundColor Green
                    $success = $true
                    break
                    }
                    else
                    {
                    Write-Host "⚠ Attempt $i failed for $targetHost" -ForegroundColor Yellow
                    Start-Sleep -Seconds $delaySeconds
                    }
            }

        if (-not $success)
        {
        Write-Host "✖ FAIL: Port 443 NOT reachable on $targetHost after $maxAttempts attempts" -ForegroundColor Red
        }

    # Log everything regardless
    #Add-Content -Path $logfilepath -Value "`n--- Testing $targetHost ---"


    # Log minimal result
        if ($result.TcpTestSucceeded)
         {
            Add-Content -Path $logfilepath -Value "Result: SUCCESS"
         }
        else
         {
            Add-Content -Path $logfilepath -Value "Result: FAILURE"
         }

    #Add-Content -Path $logfilepath -Value "Remote Address: $($result.RemoteAddress)"

    #"------------------------------`n" | Out-File -FilePath $logfilepath -Append

    }



Read-Host 'Press Enter to exit'
'@

# Save to temp script file
$tempScriptPath = "$env:TEMP\run-port-check.ps1"
$scriptContent | Set-Content -Path $tempScriptPath -Encoding UTF8

# Execute in new PowerShell window
start-Process powershell.exe -ArgumentList "-NoExit", "-ExecutionPolicy Bypass", "-File `"$tempScriptPath`""

})


# Button close form
$close = New-Object 'System.Windows.Forms.Button'
#$close.Font = 'Calibri, 12.25pt'
$close.DialogResult = 'OK'
$close.Location = '530,720'
$close.Margin = '5, 5, 5, 5'
$close.Name = 'buttonOK'
$close.Size = '100, 25'
$close.BackColor ="white"
$close.ForeColor ="black"
$close.Text = 'Exit'
$close.UseCompatibleTextRendering = $True
$close.UseVisualStyleBackColor = $False
$close.Add_Click({$main_form.Close()})    
$close.Show()#$button.Hide()
$main_form.Controls.Add($close)

# Disable other types of close/exit
#$form.add_FormClosing({$_.Cancel=$true})
#[void] $form.ShowDialog()

$main_form.ShowDialog()

#Close if Papercut Hive is not running
#Else {
#powershell -WindowStyle hidden -Command "& {[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); [System.Windows.Forms.MessageBox]::Show('Papercut Hive EdgeNode service is not running on this device !!','WARNING')}"
#}