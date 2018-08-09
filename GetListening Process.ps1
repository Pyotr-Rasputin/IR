Write-Host "Gathering the list of listening ports."
$NS = netstat -nao |Select-String "Listening"
$InfoList = @()

Write-Host "Parsing the list of listening ports, to gather additional information."
foreach ($Record in $NS){ 
    
    $RecordSplit = $Record -split '\s\s*'    
    $Port = "" 
   
    if($r[2] -match '\['){
        $Port = $RecordSplit[2] -split "\]:"
    }
    else{        
        $Port = $RecordSplit[2] -split ":"
    }  
  
    
    $ParentName = Get-Process | Where-Object -Property Id -EQ $RecordSplit[5]  
    $processInfo = Get-WmiObject Win32_Process | Select ProcessId,Name,CommandLine | Where-Object {$_.ProcessId -eq $RecordSplit[5]}
  
    $GenericObject = New-Object -TypeName PSObject
    $GenericObject | Add-Member -MemberType NoteProperty -Name 'Protocol' -Value $RecordSplit[1]
    $GenericObject | Add-Member -MemberType NoteProperty -Name 'Port' -Value $Port[1]
    $GenericObject | Add-Member -MemberType NoteProperty -Name 'ProcessID' -Value $RecordSplit[5]   
    $GenericObject | Add-Member -MemberType NoteProperty -Name 'ProcessName' -Value $ParentName.ProcessName  
    $GenericObject | Add-Member -MemberType NoteProperty -Name 'Executable' -Value $processInfo.Name  
    $GenericObject | Add-Member -MemberType NoteProperty -Name 'Command' -Value $processInfo.CommandLine
    $InfoList += $GenericObject
}

Write-Host "Here is the list of listening ports, the process that launched them, and the command line that did so; where available."
Foreach($InfoListElement in $InfoList){
    $InfoListElement
}
