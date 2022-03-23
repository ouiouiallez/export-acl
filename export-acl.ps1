param(
    [string]$out,
    [Int]$depth = 1,
    [string]$scan,
    [switch]$help,
    [switch]$q,
    [string]$style="Light13"
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
        while($exit -eq $false){
            $cell = $worksheet.Cells.Item($row,$column)
            if($null -eq $cell.Value){
                $exit = $true
            }else{
                Set-ExcelColumn -ExcelPackage $file -Worksheetname $worksheet -Column $column -Width 50
            }
                $column += 1
            }
        }
        #format cells for backlines and text on top
        Set-Format -Address $worksheet.Cells -WrapText -VerticalAlignment Top 
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
    return ((Get-ChildItem -Directory -Path $working_dir -depth $depth -Force) | Sort-Object -Property FullName)
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
    $allRights = getAllRights -childFolders $childFolders #gets all rights given to determine the number of columns

    #iterates to append to each line 
    Foreach ($Folder in $childFolders) {
        $Acl = Get-Acl -Path $Folder.FullName #gets ACL for current folder
        $rightsAndNames = getRightsAndMembers -acl $Acl 

        $rootParenthesis = "(" + $root + ")" # allows folder name to be printed even if it has the same name as the root folder
        $path = ($Folder.FullName -split $rootParenthesis,2)[-1]
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

    $file = $Report | Export-Excel $dest -WorksheetName $root -PassThru -TableStyle $style
    format -file $file
    Close-ExcelPackage $file
}

<#
.Description
Returns a hashtable with the Right type as key (ie FullControl), and the users asociated as value.
#>
function getRightsAndMembers{
    param(
        $acl
    )
    $rightsAndNames =@{} # hash table recensant droits et utilisateurs / groupes associés à ceux ci
    foreach($access in $acl.Access){
        $namesAndMembers = "" 
        $name = $access.IdentityReference
        if(isADGroup -name $name){#if the name is an AD Group
            $ADGroup = $name.Value.split("\")[-1]#strips the "domainname\ before username"
            $namesAndMembers += $name.Value + "{" + (getMembers -groupName $ADGroup) + "}, "
        }else{#if its a username
            $namesAndMembers += $name.Value + " "
        }
        if($rightsAndNames.ContainsKey($access.FileSystemRights)){
            $rightsAndNames[$access.FileSystemRights] += $namesAndMembers
        }else{
            $rightsAndNames.add($access.FileSystemRights, $namesAndMembers)
        }
    }
    return $rightsAndNames
}

<#
.Description
Customizable function to determine if the given name is a user or a AD Group, you decide how to run the tests. 
#>
function isADGroup{
    param(
        $name
    )
    $splitted = $name.Value.split("\")[-1]
    $query = (Get-ADObject -Filter 'objectClass -eq "group" -and Name -eq $splitted')
    if($null -eq $query -or "" -eq $query){
        return $false
    }else{
        return $true
    }
}
<#
.Description
Return a string listing all the members of an AD group
#>
function getMembers{
    param(
        $groupName
    )
    $arrayMembers = Get-ADGroupMember -identity $groupName -recursive | Select-Object SamAccountName
    $stringMembers=""
    foreach($key in $arrayMembers){
        $member = ((out-string -inputobject $key).Split([Environment]::NewLine)[6]).Trim()#isolates name
        $stringMembers += " " + $member     
    }
    return $stringMembers
}

<#
.Description
Return an arrray containing the different rights given throughout the folders, allowing to determine the exact number of columns for the sheet
#>
function getAllRights{
    param(
        $childFolders
    )
    $rightsArray=@()
    foreach($folder in $childFolders){
        $Acl = Get-Acl -Path $folder.FullName
        foreach($accessType in $Acl.Access){
            if($false -eq ($rightsArray -contains $accessType)){
                $rightsArray += $accessType.FileSystemRights
            }
        }
    }
    return $rightsArray
}

function getPaths{
    param(
        $userinput
    )
    $array = @()
    if(isDirectory -userinput $userinput){
        $array += $userinput
    }else{
        $array = get-content $userinput
    }
    return $array

}

function isDirectory{
    param(
        $userinput
    )
    if((Get-Item $userinput) -is [System.IO.DirectoryInfo]){
        return $true
    }else{
        return $false
    }
}

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

function checkStyles{
    param(
        [string]$style
    )
    $allStyles = @()
    for($i = 1; $i -le 21; $i += 1){$allStyles += ("Light" + $i)}
    for($i = 1; $i -le 28; $i += 1){$allStyles += ("Medium" + $i)}
    for($i = 1; $i -le 11; $i += 1){$allStyles += ("Dark" + $i)}
    if($allStyles -contains $style){
        return $true
    }else{
        return $false
    }
}

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
        Write-Host "Please specify -out and -scan parameters. `nUse -help for more details."
    }elseif($null -eq $style -or $false -eq (checkStyles($style))){
        Write-Host "Please specify a valid style name.`nType .\export-acl -help to see possibilities."
    }else{
        $ok = $true
    }
    return $ok
}

#------------------------ MAIN ------------------------#
$ok = checkRequirementsAndInput

if($ok -eq $true){
    foreach($dir in getPaths -userinput $scan){
        $root = getRoot -path $dir
        if($q -ne $true){write-host -nonewline "Scanning $root..."}
        Export -childFolders (Get-Child-Recurse -depth $depth -working_dir $dir) -dest $out -root $root
        if($q -ne $true){Write-Host "Done"}
    }
}else{
    Exit
}
    

