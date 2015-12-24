$TimeBegin = Get-Date
$InputPath = ".\ComputerList.csv"
$ComputerList = Get-Content -Path $InputPath 
$Time = Get-Date -Format 'MMddyyyy_hhmmss'
$OutputDir = "Output_$Time"
mkdir $OutputDir
$OutputPath = ".\$OutputDir\Output.csv" 
$Results = @();

#Grabs info about monitor size and make - calls monMake
function getMonitorSize($pc) {
    $oWmi = Get-WmiObject -ComputerName $pc -Namespace 'root\wmi' -Query "SELECT MaxHorizontalImageSize,MaxVerticalImageSize,InstanceName FROM WmiMonitorBasicDisplayParams" -ErrorAction SilentlyContinue 
    $sizes = @();
    $make = @();
    $output = ""
    $count = $oWmi.Count

    #Compute monitor size based on Max horizontal and vertical image size
    if ($count -gt 1) {
        foreach ($i in $oWmi) {
            $x = [Math]::Pow($i.MaxHorizontalImageSize/2.54,2)
            $y = [Math]::Pow($i.MaxVerticalImageSize/2.54,2)
            $sizes += [Math]::Round([Math]::Sqrt($x + $y),0)
            $make += monMake($i.InstanceName)
        }
    } 
    else {
        $count = 1
        $x = [Math]::Pow($oWmi.MaxHorizontalImageSize/2.54,2)
        $y = [Math]::Pow($oWmi.MaxVerticalImageSize/2.54,2)
        $sizes += [Math]::Round([Math]::Sqrt($x + $y),0)
        $make += monMake($oWmi.InstanceName)
    }
    #This will improve output for multiple monitors of the same make and size
    $umake = $make | Select -Unique
    $usizes = $sizes | Select -Unique

    if ($count -eq 0 -or $usizes -lt 17) {
        $monitor = "N/A"
    }
    elseif ($usizes.Count -eq 1 -and $umake.Count -eq 0) {
        $monitor = "$count" + " - $usizes" + '" ' + "Unknown" 
    }
    elseif ($usizes.Count -eq 1 -and $umake.Count -eq 1) {
        $monitor = "$count - $usizes" + '" ' + "$umake"
    } 
    else {
        for ($i = 0; $i -le $count - 1; $i++) {
            $monitor += $sizes[$i].ToString() + '" ' + $make[$i] + " "
        }
    }
    return $monitor
}
#Formats make of monitor - called by the above function
function monMake($iName) {
    $maker = $null
    if ($iName -ne $null) {
        $str = Out-String -InputObject $iName
        $str = $str.Split("\")[1].Substring(0,3)
        if ($str -eq "SAM") {
            $maker = "Samsung"
        }
        elseif ($str -eq "HWP") {
            $maker = "HP"
        }
        elseif ($str -eq "DEL") {
            $maker = "Dell"
        }     
    }
    return $maker
}
#This function was changed from using GWMI win32_Product class but that was very slow so I just tested the office paths
#I am hoping to improve this so it is more dynamic and can catch situations where only one or two office apps are installed
#this function works for 99% of computers out there
function getOffice($pc) {
    #$office = Get-WmiObject -Class Win32_Product -ComputerName $pc | select Name | where { $_.Name -match “Office”}
    $oVer = "N/A"
    $full32_07 = "\\$pc\c$\Program Files\Microsoft Office\Office12\WINWORD.EXE"
    $full64_07 = "\\$pc\c$\Program Files (x86)\Microsoft Office\Office12\WINWORD.EXE"
    $view32_07 = "\\$pc\c$\Program Files\Microsoft Office\Office12\XLVIEW.EXE"
    $view64_07 = "\\$pc\c$\Program Files (x86)\Microsoft Office\Office12\XLVIEW.EXE"
    $full32_10 = "\\$pc\c$\Program Files\Microsoft Office\Office14\WINWORD.EXE"
    $full64_10 = "\\$pc\c$\Program Files (x86)\Microsoft Office\Office14\WINWORD.EXE"
    $view32_10 = "\\$pc\c$\Program Files\Microsoft Office\Office14\XLVIEW.EXE"
    $view64_10 = "\\$pc\c$\Program Files (x86)\Microsoft Office\Office14\XLVIEW.EXE"

    if ((Test-Path -Path $full32_07) -or (Test-Path -Path $full64_07)) {
        $oVer = "Full 2007"
    }
    elseif ((Test-Path $view32_07) -or (Test-Path -Path $view64_07)) {
        $oVer = "Viewers 2007"
    }
    if ((Test-Path -Path $full32_10) -or (Test-Path -Path $full64_10)) {
        $oVer = "Full 2010"
    }
    #Not sure of 2010 even has viewers... just in case 
    elseif ((Test-Path $view32_10) -or (Test-Path -Path $view64_10)) {
        $oVer = "Viewers 2010"
    }
    return $oVer 
}  
#Grab wireless and LAN MAC addresses
function getNetwork($pc) {
    $adapters = Get-WmiObject -ComputerName $pc win32_NetworkAdapter
    $lanMac, $wanMac = $null 

    foreach ($a in $adapters) {
        
        if ($a.NetConnectionID -eq "Local Area Connection") {
            $lanMac = $a.MacAddress
        }
        elseif ($a.NetConnectionID -eq "Wireless Network Connection") {
            $wanMac = $a.MacAddress
        }
    }
    if ($lanMac -eq $null) {
        $lanMac = "N/A"
    }
    if ($wanMac -eq $null) {
        $wanMac = "N/A"
    }
    $network = New-Object psobject
    $network | Add-Member -MemberType NoteProperty -Value $lanMac -Name EthernetMac
    $network | Add-Member -MemberType NoteProperty -Value $wanMac -Name WirelessMAC
    return $network
}
#query for and format hardware and OS related info
function getHardware($pc) {
    #query for hardware/OS info
    $serial = Get-WmiObject -Class Win32_bios -ComputerName $pc | select -Property SerialNumber
    $Win32_Computer = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $pc | select -Property Model,Manufacturer,TotalPhysicalMemory,UserName
    $rawOS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $pc | select -Property Name,Description,OSArchitecture
    $Model = $Win32_Computer.Model
    $cpu = Get-WmiObject -Class Win32_Processor -ComputerName $pc | select -Property Name
    
    #format ram output
    $dc = $Win32_Computer.TotalPhysicalMemory/1024
    $dc = $dc/1024
    $dc = $dc/1024
    $RAM = [math]::Ceiling($dc)
    
    #format username output
    if ($Win32_Computer.UserName -ne $null) {
        $un = $Win32_Computer.UserName.split("\")[1] 
    }
    else {
        $un = "N/A"
    }

    #format OS
    $os = $rawOS.Name.Split("|")[0]

    if ($os -match "Microsoft Windows 7 Enterprise") {
        $os = "WIN 7 ENT"
    }
    elseif ($os -match "Microsoft Windows 7 Pro") {
        $os = "WIN 7 PRO"
    }

    #format cpu output

    #construct hardware PSObject with formatted info
    $hardware = New-Object psobject
    $hardware | Add-Member -MemberType NoteProperty -Value $serial.SerialNumber -Name Serial
    $hardware | Add-Member -MemberType NoteProperty -Value $rawOS.Description -Name Description
    $hardware | Add-Member -MemberType NoteProperty -Value "$Model".Trim() -Name Model
    $hardware | Add-Member -MemberType NoteProperty -Value "$RAM GB" -Name RAM
    $hardware | Add-Member -MemberType NoteProperty -Value $os -Name OS
    $hardware | Add-Member -MemberType NoteProperty -Value $un -Name Username
    $hardware | Add-Member -MemberType NoteProperty -Value $cpu.Name -Name CPU
    $hardware | Add-Member -MemberType NoteProperty -Value $rawOS.OSArchitecture -Name Bit
    return $hardware
}
#Query AD for the distinguished name of the passed in hostname and format it to immediate OU
function getOU($pc) {   
    $searchBase = "OU=Sites CCHS,DC=Centracare,DC=Com"
    $domCom = Get-ADComputer -SearchBase $searchBase -Filter {Name -eq $pc}
    if ($domCom -ne $null) {
        if ($domCom -is [System.Array]) {
            foreach ($d in $domCom) {
                $Out = $d.ToString().Substring($pc.Length + 4).Split(',')[0].Substring(3)
                if ($Out -notmatch "Intel AMT") {
                    $OU = $Out
                }
            }
        }
        else { 
            $OU = $domCom.ToString().Substring($pc.Length + 4).Split(',')[0].Substring(3)
        }
    }
    else {
        $OU = "N/A"
    }
    return $OU
}
#Grabs user's full name
function getFullName ($un) {
    $searchBase = "OU=Users,OU=Sites CCHS,DC=Centracare,DC=Com"
    $adUN = Get-ADUser -SearchBase $searchBase -Filter {sAMAccountName -eq $un}
    $name = $adUN.GivenName + " " + $adUN.Surname
    return $name
}
#this contructs our psobject that contains all of our data
#doing it this way allows me to handle an empty row for pcs off the network
function getObject($pc, $ou, $desc, $un, $mod, $ser, $wty, $ram, $cpu, $lan, $wan, $jackID, $os, $mon, $off, $bit) {
    $InventoryItem = New-Object psobject 
    #Hostname, DHCP, Description, username, Model, Serial, wty, RAM, CPU, Lan MAC, WAN MAC, OS, monitors, and office
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $pc -Name Hostname
    $InventoryItem | Add-Member -MemberType NoteProperty -Value "DHCP" -Name IP
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $desc -Name Description
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $un -Name Username
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $ou -Name OU
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $mod -Name Model
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $ser -Name Serial
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $wty -Name WTY
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $ram -Name RAM
    #$InventoryItem | Add-Member -MemberType NoteProperty -Value $cpu -Name CPU
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $lan -Name EthernetMac
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $wan -Name WirelessMac
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $jackID -Name JackID
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $os -Name OS
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $mon -Name Monitor
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $off -Name Office   
    $InventoryItem | Add-Member -MemberType NoteProperty -Value $bit -Name Architecture

    return $InventoryItem  
}

#this loops over the hn list and calls the functions that query the pcs for info
foreach ($PC in $ComputerList) {
    #This test connection to PC if it fails it will write just the HN in csv
    if (Test-Connection -ComputerName $PC -Quiet -Count 1) {
        Write-Verbose "Connection to $PC succeeded..." -Verbose
        $monitor = getMonitorSize($PC)
        $office = getOffice($PC) 
        $net = getNetwork($PC)
        $hd = getHardware($PC)
        $ou = getOU($PC)
        #Place holders
        $wty = " "
        $jack = " "
       
        if ($ou -eq "Private Citrix" -or $ou -eq "EMR Citrix" -or $ou -eq "infosys" -or $ou -eq "Non-EMR" -and $hd.Username -ne $null) {
            $hd.Username = getFullName($hd.Username) 
        }
        #costruct the object that contains the data and add it to the results array
        $Results += getObject $PC $ou $hd.Description $hd.Username $hd.Model $hd.Serial $wty $hd.RAM $hd.CPU $net.EthernetMac $net.WirelessMac $jack $hd.OS $monitor $office $hd.Bit
    }
    else {
        Write-Verbose "Connection to $PC failed..." -Verbose
        $Results += getObject($PC)
    }
}
#write the results array to the output csv
$Results | Export-Csv -NoTypeInformation -Path $OutputPath
$duration = $(Get-Date) - $TimeBegin
Write-Host "This script ran for $duration" 