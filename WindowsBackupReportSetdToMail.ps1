Param(
    [Parameter(Mandatory)][string] $Root,
    [switch] $txt=$false,
    [switch] $mail=$false,
    [string] $SmtpSrv,
    [int] $Port=25,
    [string] $To,
    [string] $Fm,
    [string] $Pwd,
    [string] $Sub
)

#------------ FILE backup handling routines -------------------

function Get-MachineListFB ($WinBackupRoot) {
$MachineRoots = @{}
    Get-ChildItem -Path $WinBackupRoot -Recurse -Depth 1 -File -Filter "MediaID.bin" | 
    Where-Object {$_.Directory.FullName -ne $WinBackupRoot} |
    foreach {
            $MachineRoots[$_.Directory.Name]=$_.Directory.FullName
    }
return $MachineRoots
}


function Get-LastFileBackupSet ($MachineRoot) {
$objLB = $null
try {
    $LastWbcat = Get-ChildItem -Path $MachineRoot -Recurse -Depth 2 -File -Filter "GlobalCatalog.wbcat" |
        Where-Object {$_.Directory.Name -eq "Catalogs"} |
        Sort-Object LastWriteTime |
        Select-Object -Last 1 LastWriteTime, Directory, `
            @{Name="Machine";Expression={$_.Directory.Parent.Parent.Name.ToUpper()}}, `
            @{Name="Path";Expression={$_.Directory.Parent.FullName}}, `
            @{Name="Name";Expression={$_.Directory.Parent.Name}}

    $ZIPFiles = (
        Get-ChildItem $LastWbcat.Directory.Parent.FullName -Recurse -Depth 2 -File  -Filter "*.zip" | 
        Select-Object Length | 
        Measure-Object -Property Length -Sum
        )
}
finally
{
    if (($ZIPFiles) -and ($LastWbcat)) {
        $objLB = @{
            Machine = $LastWbcat.Machine
            Path = $LastWbcat.Path
            Name = $LastWbcat.Name
            Updated = $LastWbcat.LastWriteTime
            Size = $ZIPFiles.Sum
        }
    }
}
  return $objLB
}


function Get-LastFileBackupSetSummaryTXTForm ($MachineRoot) {
$ret = ""
$Now = Get-Date
$LastBT = Get-LastFileBackupSet ($MachineRoot)
if ($LastBT) {
    $ret =  "Machine: " + $LastBT.Machine + "`r`nSet name: " + $LastBT.Name + "`r`nLast performed: " + ($Now - $LastBT.Updated).Days + " day(s) ago. (" + $LastBT.Updated + ")`r`nSize: " + ('{0:0} {1}' -f ($LastBT.Size/1024/1024), "MB")
  }
else{
    $ret = "Data retrieving error."
}
return $ret
}


function Get-LastFileBackupSetSummaryTXTRow ($MachineRoot) {
$ret = ""
$Now = Get-Date
$LastBT = Get-LastFileBackupSet ($MachineRoot)
if ($LastBT) {
    $ret = "Machine: " + $LastBT.Machine + "; Set name: " + $LastBT.Name + "; Last performed: " + ($Now - $LastBT.Updated).Days + " day(s) ago. (" + $LastBT.Updated + "); Size: " + ('{0:0} {1}' -f ($LastBT.Size/1024/1024), "MB")
  }
else{
    $ret = "Data retrieving error."
}
return $ret
}


function Get-LastFileBackupSetSummaryHTMLRow ($MachineRoot) {
$ret = ""
$Now = Get-Date
$LastBT = Get-LastFileBackupSet ($MachineRoot)
if ($LastBT) {
    $DaysDiff = ($Now - $LastBT.Updated).Days
    if ($DaysDiff -eq 0) {$ColorCode=""}
    if (($DaysDiff -eq 1) -or ($DaysDiff -eq 2)) {$ColorCode=' style="background-color: #FFFFDD"'}
    if ($DaysDiff -gt 2) {$ColorCode=' style="background-color: #FFAAAA"'}
    $ret = "<tr><td>" + $LastBT.Machine + "</td><td>" + $LastBT.Name + "</td><td align=right" + $ColorCode + ">" + $DaysDiff + " day(s)</td><td>" + $LastBT.Updated + "</td><td align=right>" + ('{0:0} {1}' -f ($LastBT.Size/1024/1024), "MB") + "</td></tr>"
  }
else{
    $ret = "<tr><td>Data retrieving error.</td></tr>"
}
return $ret
}


function Get-LastFileBackupDatesTXT ($WinBackupRoot) {
$table="File backup:`r`n"
$MachineRoots = Get-MachineListFB ($WinBackupRoot)
foreach ($machine in $MachineRoots.Keys) {
    $line = Get-LastFileBackupSetSummaryTXTRow($MachineRoots[$machine])
    $table+= -join($line, "`r`n")
    }
return $table
}


function Get-LastFileBackupDatesHTML ($WinBackupRoot) {
$table+='<h4>File backup:'+"</h4>`n"
$table+="<table border=1>`n"
$table+='<tr><td><h4>Machine</h4></td><td><h4>Latest backup set dir</h4></td><td><h4>Age</h4></td><td><h4>Updated</h4></td><td><h4>Size</h4></td></tr>'+"`n"
$MachineRoots = Get-MachineListFB ($WinBackupRoot)
foreach ($machine in $MachineRoots.Keys) {
    $line = Get-LastFileBackupSetSummaryHTMLRow($MachineRoots[$machine])
    $table+= -join($line, "`n")
    }
$table+= "</table>`n"
return $table
}


#------------ IMAGE backup handling routines -------------------

function Get-MachineListIB ($WinBackupRoot) {
$MachineRoots = @{}
$ImageBackupRoot = $WinBackupRoot+"\WindowsImageBackup"
    Get-ChildItem -Path $ImageBackupRoot -Recurse -Depth 1 -File -Filter "MediaID" | 
    foreach {
        $MachineRoots[$_.Directory.Name]=$_.Directory.FullName
        }
return $MachineRoots
}


function Get-LastImageBackupSet ($MachineRoot) {
$objLB = $null
try {
    $LastBSXML = Get-ChildItem -Path $MachineRoot -Recurse -Depth 1 -File -Filter "BackupSpecs.xml" | 
        Sort-Object LastWriteTime |
        Select-Object -Last 1 LastWriteTime, Directory, `
            @{Name="Machine";Expression={$_.Directory.Parent.Name.ToUpper()}}, `
            @{Name="Path";Expression={$_.Directory.FullName}}, `
            @{Name="Name";Expression={$_.Directory.Name}}
    $VHDFiles = (
        Get-ChildItem $LastBSXML.Directory.FullName -File -Filter "*.vhd" | 
        Select-Object Length | 
        Measure-Object -property length -sum -ErrorAction SilentlyContinue)
}
finally
{
  if (($VHDFiles) -and ($LastBSXML)) {
      $objLB = @{
          Machine = $LastBSXML.Machine
          Path = $LastBSXML.Path
          Name = $LastBSXML.Name
          Updated = $LastBSXML.LastWriteTime
          Size = $VHDFiles.Sum
      }
  }
}
  return $objLB
}


function Get-LastImageBackupSetSummaryTXTForm ($MachineRoot) {
$ret = ""
$Now = Get-Date
$LastBT = Get-LastImageBackupSet ($MachineRoot)
if ($LastBT) {
    $ret =  "Machine: " + $LastBT.Machine + "`nSet name: " + $LastBT.Name + "`nLast performed: " + ($Now - $LastBT.Updated).Days + " day(s) ago. (" + $LastBT.Updated + ")`nSize: " + ('{0:0} {1}' -f ($LastBT.Size/1024/1024), "MB`r`n")
  }
else{
    $ret = "Data retrieving error."
}
return $ret
}


function Get-LastImageBackupSetSummaryTXTRow ($MachineRoot) {
$ret = ""
$Now = Get-Date
$LastBT = Get-LastImageBackupSet ($MachineRoot)
if ($LastBT) {
    $ret = "Machine: " + $LastBT.Machine + "; Set name: " + $LastBT.Name + "; Last performed: " + ($Now - $LastBT.Updated).Days + " day(s) ago. (" + $LastBT.Updated + "); Size: " + ('{0:0} {1}' -f ($LastBT.Size/1024/1024), "MB")
  }
else{
    $ret = "Data retrieving error."
}
return $ret
}


function Get-LastImageBackupSetSummaryHTMLRow ($MachineRoot) {
$ret = ""
$Now = Get-Date
$LastBT = Get-LastImageBackupSet ($MachineRoot)
if ($LastBT) {
    $DaysDiff = ($Now - $LastBT.Updated).Days
    if ($DaysDiff -eq 0) {$ColorCode=""}
    if (($DaysDiff -eq 1) -or ($DaysDiff -eq 2)) {$ColorCode=' style="background-color: #FFFFDD"'}
    if ($DaysDiff -gt 2) {$ColorCode=' style="background-color: #FFAAAA"'}
    $ret = "<tr><td>" + $LastBT.Machine + "</td><td>" + $LastBT.Name + "</td><td align=right" + $ColorCode + ">" + $DaysDiff + " day(s)</td><td>" + $LastBT.Updated + "</td><td align=right>" + ('{0:0} {1}' -f ($LastBT.Size/1024/1024), "MB") + "</td></tr>"
  }
else{
    $ret = "<tr><td>Data retrieving error.</td></tr>"
}
return $ret
}


function Get-LastImageBackupDatesTXT ($WinBackupRoot) {
$table="Image backup:`r`n"
$MachineRoots = Get-MachineListIB ($WinBackupRoot)
foreach ($machine in $MachineRoots.Keys) {
    $line = Get-LastImageBackupSetSummaryTXTRow($MachineRoots[$machine])
    $table+= -join($line, "`r`n")
    }
return $table
}


function Get-LastImageBackupDatesHTML ($WinBackupRoot) {
$table+='<h4>Volume image backup:'+"</h4>`n"

$table+="<table border=1>`n"
$table+='<tr><td><h4>Machine</h4></td><td><h4>Latest backup set dir</h4></td><td><h4>Age</h4></td><td><h4>Updated</h4></td><td><h4>Size</h4></td></tr>'+"`n"
$MachineRoots = Get-MachineListIB ($WinBackupRoot)
foreach ($machine in $MachineRoots.Keys) {
    $line = Get-LastImageBackupSetSummaryHTMLRow($MachineRoots[$machine])
    $table+= -join($line, "`n")
    }
$table+= "</table>`n"

return $table
}


#------------ Common routines -------------------

function Get-LastWinBackupDatesHTML ($WinBackupRoot) {
$Now = Get-Date
$table='<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'+"`n"
$table+='<html xmlns="http://www.w3.org/1999/xhtml"><head><title>Backup volumes at "'+$WinBackupRoot+'"</title></head><body>'+"`n"
$table+='<h3>Backup volumes at "'+$WinBackupRoot+'"'+"</h3>`n"
$table+='<h3>Generated: '+$Now+"</h3>`n"
$table+= Get-LastFileBackupDatesHTML ($WinBackupRoot)
$table+= Get-LastImageBackupDatesHTML ($WinBackupRoot)
$table+= "</body>`n</html>"
return $table
}

function Get-LastWinBackupDatesTXT ($WinBackupRoot) {
$Now = Get-Date
$table='Backup volumes at "'+$WinBackupRoot+'"'+"`r`n"
$table+='Generated: '+$Now+"`r`n`r`n"
$table+= Get-LastFileBackupDatesTXT ($WinBackupRoot)
$table+= "`r`n"
$table+= Get-LastImageBackupDatesTXT ($WinBackupRoot)
return $table
}


function Send-Email-Report ($SendInfo, $TXTonly=$false){
$Credential = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $SendInfo.Username, $SendInfo.SecurePassword
if ($TXTonly) {
    Send-MailMessage -To $SendInfo.To -From $SendInfo.From -Subject $SendInfo.Subject -Body $SendInfo.MsgBody -SmtpServer $SendInfo.SmtpServer -Credential $Credential -Port $SendInfo.SmtpPort -Encoding UTF8
    }
else
    {
    Send-MailMessage -To $SendInfo.To -From $SendInfo.From -Subject $SendInfo.Subject -BodyAsHtml $SendInfo.MsgBody -SmtpServer $SendInfo.SmtpServer -Credential $Credential -Port $SendInfo.SmtpPort
    }
}


#------------ MAIN ROUTINE -------------------

if ($txt) {
    $Ret = Get-LastWinBackupDatesTXT ($Root)
    }
else {
    $Ret = Get-LastWinBackupDatesHTML ($Root)
    }

if ($mail) {
    if (($SmtpSrv) -and ($Port) -and ($To) -and ($Fm) -and ($Pwd) -and ($Sub) -and ($Ret)) {
            $objArgs = @{
                SmtpServer = $SmtpSrv
                SmtpPort = $Port
                To = $To.Split(',')
                From = $Fm
                Username = $Fm
                SecurePassword = ConvertTo-SecureString -String $Pwd -AsPlainText -Force
                Subject = $Sub
                MsgBody = $Ret
            }
            Send-Email-Report ($objArgs, $txt)
        }
    else {
        Write-Host $Ret
        }
    }
else {
    Write-Host $Ret
    }

