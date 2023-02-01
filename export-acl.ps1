param(
    [string]$out,
    [Int]$depth = 1,
    [string]$scan,
    [switch]$help,
    [switch]$q,
    [string]$style="Light13",
    [switch]$split,
    [switch]$noninherited,
    [switch]$onlyusers,
    [switch]$fullnames,
    [switch]$nobuiltin,
    [switch]$nosystem
)

<#
.Description
Formats the excel document passed onto the function, with arbitrary values.
#>
function format{
    param(
        $file
    )
    
    foreach($worksheet in $file.Workbook.Worksheets){
        $row = 1
        $column = 1
        $exit = $false
        #iterates on every cell with a value
        while(!$exit){
            $cell = $worksheet.Cells.Item($row,$column)
            if($null -eq $cell.Value){
                $exit = $true
            }else{
                Set-ExcelColumn -ExcelPackage $file -Worksheetname $worksheet -Column $column -Width 50
            }
        $column += 1
        }
        deleteEmptyColumns -ws $worksheet -maxcols $column -maxrows (countRows -ws $worksheet)
    }
    #format cells for backlines and text on top
    Set-Format -Address $worksheet.Cells -WrapText -VerticalAlignment Top 
}

<#
.Description
Counts number of rows within the given worksheet
#>
function countRows{
    param(
        $ws
    )
    $row = 1
    $column = 1
    $end = $false
    while(!$end){
        $value = $ws.Cells.Item($row,$column).Value
        if($null -eq $value){$end = $true}else{$row += 1}
    }
    return $row
}

<#
.Description
Deletes empty columns within the given worksheet (ie no users are within this column, because they have been removed from the report with -nobuiltin or nosystem)
#>
function deleteEmptyColumns{
    param(
        $ws,
        $maxcols,
        $maxrows
    )
    for($curcol = 2;$curcol -le $maxcols;$curcol++){ # for each column
        $isempty = $true
        for($currow = 2;$currow -le $maxrows;$currow++){ # for each row in this column
            $value = $ws.Cells.Item($currow,$curcol).Value
            if($null -ne $value){
                $isempty = $false
            }
        }
        if($isempty){$ws.DeleteColumn($curcol)}
    }
}

<#
.Description
Returns all the child folders of a given folder, with the depth passed as a parameter
#>
function Get-Child-Recurse{
    param(
        [string]$working_dir,
        $depth
    )
    #get parent folder if available, for the first line of the report
    $result = @()
    if((Get-Item $working_dir)){
        $result += ((Get-Item $working_dir))
    }
    if($depth -eq -1){
        $result += ((Get-ChildItem -Directory -Path $working_dir -recurse -Force -ErrorAction silentlycontinue) | Sort-Object -Property FullName)
    }else{
        $result +=  ((Get-ChildItem -Directory -Path $working_dir -depth $depth -Force -ErrorAction silentlycontinue) | Sort-Object -Property FullName)
    }
    return $result
}

<#
.Description
Core function of this script. It takes all the rights and usernames and puts it in a variable to be exported to an excel sheet, thanks to Import-Excel module.
#>
function Export{
    param(
        $childFolders,
        $root,
        $dest
    )
    $Report = @()
    $allRights,$allacls = getAllRights -childFolders $childFolders #gets all rights given to determine the number of columns

    $current = 0 #working var to print the progression
    #iterates to append to each line 
    Foreach ($Folder in $childFolders) {

        if(!($allacls[$folder.Fullname])){#do not append a line if for some reason the acl could not be fetched. 
            continue
        }
        $Acl = $allacls[$folder.Fullname] #gets ACL for current folder

        if(!$q){#write progress bar
            $current += 1
            $percentage = [math]::Round((($current / $childFolders.Length) * 100))
            Write-Progress -Activity "Exporting ACLs for $root..." -Status "$percentage% Complete:" -PercentComplete $percentage
        }

        #if -noninherited is called, check if any of the rights is noninherited before processing
        if($noninherited){
            $break = $true
            foreach($access in $Acl.Access){
                #if it is not system or builtin account we are talking about (messed up my previous reports)
                if(!($access.IdentityReference.Value.split("\")[0].Contains("NT")) -and !($access.IdentityReference.Value.split("\")[0].Contains("BUILTIN"))){
                    if($access.IsInherited -eq $false){
                        $break = $false             
                    }
                }                
            }
            if($break){
                continue
            }
        }
        $rightsAndNames = getRightsAndMembers -acl $Acl 

        $rootParenthesis = "(" + $root + ")" # allows folder name to be printed even if it has the same name as the root folder
        $path = ($Folder.FullName -split $rootParenthesis,2)[-1]
        if($path -eq "\"){$path = "../$root"}#if this is the root folder report line
        #creates fields to export
        $Properties = [ordered]@{'Path'=$path}
        #adds a column per access type and creates fields
        foreach($right in $allRights){
            if($Properties.Contains($right)){
                $Properties[$right] = $rightsAndNames[$right]
            }else{
                $Properties.add($right,$rightsAndNames[$right])
            }
        }
        $Report += New-Object -TypeName PSObject -Property $Properties     
    }

    #exports to different files if -split is invoked
    if($split){
        if($dest[-1] -ne "\"){ # if the path given by the user lacks a final "\"
            $dest += "\"
        }
        $filename = $dest + $root + ".xlsx"
    }else{#-out parameter is given without formatting
        $filename = $dest
    }
    
    $file = $Report | Export-Excel $filename -WorksheetName $root -PassThru -TableStyle $style
    format -file $file
    Close-ExcelPackage $file
}

<#
.Description
Returns a hashtable with the Right type as key (ie FullControl), and the users associated as value.
#>
function getRightsAndMembers{
    param(
        $acl
    )
    $rightsAndNames =@{} # hash table with right as a key, and users as value
    foreach($access in $acl.Access){
        $namesAndMembers = "" 
        $name = $access.IdentityReference

        #exclude groups if $nobuiltin or $nosystem are specified
        if($nobuiltin){if($name.Value.split("\")[0].Contains("BUILTIN")){continue}}
        if($nosystem){if($name.Value.split("\")[0].Contains("NT")){continue}}

        if(isADGroup -name $name){#if the name is an AD Group
            $ADGroup = $name.Value.split("\")[-1]#strips the "domainname\ before username"
            if($onlyusers){
                if(getMembers -groupName $ADGroup){$namesAndMembers += (getMembers -groupName $ADGroup) + ", "}
            }else{
                $namesAndMembers += $name.Value + "{" + (getMembers -groupName $ADGroup) + "}`n"
            }            
        }else{#if its a username
            if($name.Value.Contains("\")){
                $namesAndMembers += $name.Value.split("\")[-1] + ", "
            }else{
                $namesAndMembers += $name.Value + ", "
            }
        }

        #get-acl cmdlet returns readandexecute even if the permission is only "list folder", so this check is for this
        if($access.inheritanceflags.tostring() -eq 'ContainerInherit'){#if permission is indeed List Folder Contents
            $filesystemrights = "List Folder Contents"
        }else{
            $filesystemrights = $access.FileSystemRights
        }
        if($rightsAndNames.ContainsKey($filesystemrights)){
            $rightsAndNames[$filesystemrights] += $namesAndMembers
        }else{
            $rightsAndNames.add($filesystemrights, $namesAndMembers)
        }
    }
    return $rightsAndNames
}

<#
.Description
Determines if the name "ie domain\g_rw_securitygroup or domain\user1 is an AD group or not"
#>
function isADGroup{
    param(
        $name
    )
    if($name.Value.Contains("\")){$splitted = $name.Value.split("\")[-1]}else{$splitted = $name.Value}
    $query = (Get-ADObject -Filter 'objectClass -eq "group" -and Name -eq $splitted')
    if($null -eq $query -or "" -eq $query){return $false}else{return $true}
}
<#
.Description
Return a string listing all the members of an AD group
#>
function getMembers{
    param(
        $groupName
    )
    $arrayMembers = Get-ADGroupMember -identity $groupName -recursive
    $stringMembers=""
    foreach($key in $arrayMembers){
        if($fullnames){
            $member = (Get-ADUser -identity $key).givenname + " " + (Get-ADUser -identity $key).surname
        }else{
            $member = $key.samaccountname
        }
        $stringMembers += $member + ", "
    }
    if($stringMembers.Length -gt 2 ){$stringMembers = $stringMembers.Substring(0,$stringMembers.Length-2)}
    return $stringMembers
}

<#
.Description
Return an array containing the different rights given throughout the folders, allowing to determine the exact number of columns for the sheet
Also returns a hashmap containing the ACLs for each and every folder (the folder fullname acts as the key, acl as the value)
#>
function getAllRights{
    param(
        $childFolders
    )
    $rightsArray=@()
    $allacls = [ordered]@{}
    $current = 0
    foreach($folder in $childFolders){
        if(!$q){#write progress bar
            $current += 1
            $percentage = [math]::Round((($current / $childFolders.Length) * 100))
            Write-Progress -Activity "Fetching ACLs for $root..." -Status "$percentage% Complete:" -PercentComplete $percentage
        }
        if(!(get-item -literalpath $folder.FullName -ErrorAction silentlycontinue)){
            if(!$q){write-host "W: Could not fetch ACLs for " $folder.Fullname -ForegroundColor Yellow}
            continue
        }else{
            $Acl = (get-item -literalpath $folder.FullName -ErrorAction silentlycontinue).getaccesscontrol()
        }
        
        
        if($allacls[$folder.Fullname]){
            if(!$q){Write-Host "W: " $folder.Fullname " already exists. Skipping..." -ForegroundColor Yellow}
            continue
        }else{
            $allacls.Add($folder.Fullname, $Acl)
        }
        
        if($Acl){
            foreach($accessType in $Acl.Access){
                if($false -eq ($rightsArray -contains $accessType)){
                    if($accessType.inheritanceflags.tostring() -eq 'ContainerInherit'){#if permission is List Folder Contents
                        $filesystemrights = "List Folder Contents"
                    }else{
                        $filesystemrights = $accessType.FileSystemRights
                    }
                    $rightsArray += $filesystemrights
                }
            }
        }elseif(!$q){
            Write-Host "W: Could not fetch ACLs for "$folder.Fullname -ForegroundColor Yellow
        }
    }
    return $rightsArray, $allacls
}

<#
.Description
Returns an array containing all the paths to scan, wether it comes from -scan or from a text file.
#>
function getPaths{
    param(
        $userinput
    )
    $array = @()
    if(isDirectory -userinput $userinput){$array += $userinput}else{$array = get-content $userinput}
    return $array
}

<#
.Description
Return true if the path given is a directory, false otherwise
#>
function isDirectory{
    param(
        $userinput
    )
    if((Test-Path $userinput) -eq $true){
        if((Get-Item $userinput) -is [System.IO.DirectoryInfo]){
            return $true        
        }
    }else{
        return $false
    }
}

<#
.Description
Returns the parent name of the given folder. Ex : C:\Parent\son\grandson will return Parent.
#>
function getRoot{
    param(
        $path
    )
    $currentDrive = Split-Path -qualifier $path
    $logicalDisk = Get-WmiObject Win32_LogicalDisk -filter "DriveType = 4 AND DeviceID = '$currentDrive'"
    $uncPath = $path.Replace($currentDrive, $logicalDisk.ProviderName)
    if($uncPath.split("\")[-1] -eq ""){
        $root = $uncPath.split("\")[-2]
    }else{
        $root = $uncPath.split("\")[-1]
    }
    return $root
}

<#
.Description
Returns true if the table style given by the user checks with the available table styles, false otherwise.
#>
function checkStyles{
    param(
        [string]$style
    )
    $allStyles = @()
    for($i = 1; $i -le 21; $i += 1){$allStyles += ("Light" + $i)}
    for($i = 1; $i -le 28; $i += 1){$allStyles += ("Medium" + $i)}
    for($i = 1; $i -le 11; $i += 1){$allStyles += ("Dark" + $i)}
    if($allStyles -contains $style){return $true}else{return $false}
}

<#
.Description
Check dependencies, and writes on host if parameters are missing or incorrect. Returns true if everything's okay.
#>
function checkRequirementsAndInput{
    $ok = $false
    #Installs Import-Excel module if not present
    if($null -eq (Get-InstalledModule | Select-String "ImportExcel")){
        Install-Module -Name ImportExcel
    }
    #checks parameters
    if($help){
        Write-Host (Get-Content -Raw -Encoding utf8 help.txt)
    }elseif($null -eq $out -or "" -eq $out -or $null -eq $scan -or "" -eq $scan){
        Write-Host "ERROR : Please specify -out and -scan parameters. `nUse -help for more details." -ForegroundColor Red
    }elseif($null -eq $style -or $false -eq (checkStyles($style))){
        Write-Host "ERROR : Please specify a valid style name.`nUse -help to see possibilities." -ForegroundColor Red
    }elseif($split -and ($false -eq (isDirectory -userinput $out))){
        Write-Host "ERROR : Please specify an existing directory if you used the -split parameter as several files will be saved." -ForegroundColor Red
    }elseif(!$split -and ($true -eq (isDirectory -userinput $out))){
        Write-Host "ERROR : Please specify a valid filename, with .xlsx extension." -ForegroundColor Red
    }elseif(!$split -and $true -eq (Test-Path $out)){
        Write-Host "ERROR : File already exists." -ForegroundColor Red
    }elseif(!$split -and "xlsx" -ne ($out.Split(".")[-1])){
        Write-Host "ERROR : Please specify a valid file extension : .xlsx" -ForegroundColor Red
    }else{$ok = $true}
    return $ok
}

#------------------------ MAIN ------------------------#
$ok = checkRequirementsAndInput

if($ok){
    foreach($dir in getPaths -userinput $scan){
        $root = getRoot -path $dir
        if(!$q){Write-Host "Fetching child folders for $root..." -NoNewline}
        $childs = Get-Child-Recurse -depth $depth -working_dir $dir
        if(!$q){Write-Host "Done."}
        Export -childFolders $childs -dest $out -root $root
    }
}else{
    Exit
}