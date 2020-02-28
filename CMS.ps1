[CmdletBinding()]
Param(
[Int]$threads=20,
[Parameter(Mandatory=$true)][String]$Term = $null,
[Parameter(Mandatory=$false)][switch]$CLSIDs = $null,
[Parameter(Mandatory=$false)][switch]$ProgIDs = $null,
[Parameter(Mandatory=$true)][String]$Depth = $null)  

$ErrorActionPreference = 'silentlyContinue'

If($ProgIDs)
{
$z = Get-ChildItem registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID -Recurse | Where-Object Name -match ProgID | ForEach GetValue ''
$y = $z | Get-Unique
$w = ForEach($x in $y){If($x -notmatch '\.1'){$x}}
If($Depth -eq 1)
{
    ForEach($v in $w)
    {
        $com = [activator]::CreateInstance([type]::GetTypeFromProgID("$v")) #Instantiate the COM Object
        $members = $com | Get-Member #Store all the members into a variable
        ForEach($member in $members)
        { 
                If($member -notmatch 'Application' -and $member -notmatch 'Parent' -and $member -notmatch 'Cell' -and $member -notmatch 'Columns' -and $member -notmatch 'Rows')
                {If($member.TypeName -match '__')
                {If($member -match $Term)
                {
                 Write-Host "$v.$member"}}}
        }
    }
}

If($Depth -eq 2)
{
    ForEach($v in $w)
    {

        $com = [activator]::CreateInstance([type]::GetTypeFromProgID("$v")) #Instantiate the COM Object
        $members = $com | Get-Member #Store all the members into a variable
        ForEach($member in $members)
        {
            $membernames = $member.Name 
            If($member -notmatch 'Application' -and $member -notmatch 'Parent' -and $member -notmatch 'Cell' -and $member -notmatch 'Columns' -and $member -notmatch 'Rows')
            {
                If($member.TypeName -match '__')
                {
                    If($member -match $Term)
                    {
                        Write-Host "$v.$member"
                    }
                    ForEach($membername in $membernames)
                    {
                        $m1member = $com.$membername | gm
                        If($m1member -notmatch 'Application' -and $m1member -notmatch 'Parent' -and $m1member -notmatch 'Cell' -and $m1member -notmatch 'Columns' -and $m1member -notmatch 'Rows')
                        {
                            If($m1member.TypeName -match '__')
                            {  
                                If($m1member -match $Term)
                                 {
                                    Write-Host "$v.$membername.$m1member"
                                 } 
                            }
                        }
                    }
                }
            }
        }
    } 
}

If($Depth -eq 3)
{
    ForEach($v in $w)
    {

        $com = [activator]::CreateInstance([type]::GetTypeFromProgID("$v")) #Instantiate the COM Object
        $members = $com | Get-Member #Store all the members into a variable
        ForEach($member in $members)
        {
            
            If($member -notmatch 'Application' -and $member -notmatch 'Parent' -and $member -notmatch 'Cell' -and $member -notmatch 'Columns' -and $member -notmatch 'Rows')
            {
                If($member.TypeName -match '__' -and $member.TypeName -notmatch '57da806104b8')
                {
                    If($member -match $Term)
                    {
                        Write-Host "$v.$member"
                    } 
                    $membernames = $member.Name       
                    ForEach($membername in $membernames)
                    {
                        $m1members = $com.$membername | gm   
                        ForEach($m1member in $m1members)
                        {
                            If($m1member -notmatch 'Application' -and $m1member -notmatch 'Parent' -and $m1member -notmatch 'Cell' -and $m1member -notmatch 'Columns' -and $m1member -notmatch 'Rows')
                            {
                                If($m1member.TypeName -match '__')
                                { 
                                    If($m1member -match $Term)
                                     {
                                     Write-Host "$v.$membername.$m1member"
                                     } 
                                    $m1membernames = $m1member.Name 
                                    ForEach($m1membername in $m1membernames)
                                    {
                                        $m2members = $com.$membername.$m1membername | gm
                                        ForEach($m2member in $m2members)
                                        {
                                            If($m2member -notmatch 'Application' -and $m2member -notmatch 'Parent' -and $m2member -notmatch 'Cell' -and $m2member -notmatch 'Columns' -and $m2member -notmatch 'Rows')
                                            {
                                                If($m2member.TypeName -match '__')
                                                {
                                                    If($m2member -match $Term)
                                                    {
                                                    Write-Host "$v.$membername.$m1membername.$m2member"
                                                    } 
                                                }
                                            }   
                                        }              
                                    }  
                                }
                            }
                        }
                    }
                }
            }
        }
    } 
}
}

If($CLSIDs)
{
$z = Get-ChildItem registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID
$y = $z | Get-Unique

If($Depth -eq 0)
{
     ForEach($v in $y)
    {#Start-Sleep -Milliseconds 100
    $v}
}

If($Depth -eq 1)
{
    ForEach($v in $y)
    {
        $key = $v.Name
        $CLSID = $key.Split("{}")[1] | Get-Unique
        $t = Get-ChildItem "registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{$CLSID}"
        if($t -match "LocalServer32")
        {
            $o = Get-ChildItem "registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{$CLSID}" | Where-Object Name -match ProgID | ForEach GetValue ''
            $w = ForEach($x in $y){If($x -notmatch '\.1'){$x}}
            $com = [activator]::CreateInstance([type]::GetTypeFromCLSID("$CLSID")) #Instantiate the COM Object
            $members = $com | Get-Member #Store all the members into a variable
            ForEach($member in $members)
            {
                If($member.TypeName -match '__')
                {
                    If($member -match $Term)
                    {
                        Write-Host "ProgID: " $o
                        $CLSID.$member
                    }
                }
            }
        }
    }
}

If($Depth -eq 2)
{
    ForEach($v in $y)
    {
        $key = $v.Name
        $CLSID = $key.Split("{}")[1] | Get-Unique
        $o = Get-ChildItem "registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{$CLSID}" | Where-Object Name -match ProgID | ForEach GetValue ''
        $t = Get-ChildItem "registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{$CLSID}"
        if($t -match "LocalServer32")
        {
            $com = [activator]::CreateInstance([type]::GetTypeFromCLSID("$CLSID")) #Instantiate the COM Object
            $members = $com | Get-Member #Store all the members into a variable
            ForEach($member in $members)
            {
                $membernames = $member.Name 
                If($member -notmatch 'Application' -and $member -notmatch 'Parent' -and $member -notmatch 'Cell' -and $member -notmatch 'Columns' -and $member -notmatch 'Rows')
                {
                    If($member.TypeName -match '__')
                    {
                        If($member -match $Term)
                        {
                            Write-Host "ProgID: $o"
                            "$CLSID.$member"
                        }
                        ForEach($membername in $membernames)
                        {
                            $m1member = $com.$membername | gm
                            If($m1member -notmatch 'Application' -and $m1member -notmatch 'Parent' -and $m1member -notmatch 'Cell' -and $m1member -notmatch 'Columns' -and $m1member -notmatch 'Rows')
                            {
                                If($m1member.TypeName -match '__')
                                {  
                                    If($m1member -match $Term)
                                     {
                                        Write-Host "ProgID: $o"
                                        "$CLSID.$member.$m1member"

                                     } 
                                }
                            }
                        }
                    }
                }
            }
        }
    } 
}

If($Depth -eq 3)
{
    ForEach($v in $y)
    {
        $key = $v.Name
        $CLSID = $key.Split("{}")[1] | Get-Unique      
        $t = Get-ChildItem "registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{$CLSID}"
        if($t -match "LocalServer32")
        {
            $o = Get-ChildItem "registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{$CLSID}" | Where-Object Name -match ProgID | ForEach GetValue ''
            $com = [activator]::CreateInstance([type]::GetTypeFromCLSID("$CLSID")) #Instantiate the COM Object
            $members = $com | Get-Member #Store all the members into a variable
            ForEach($member in $members)
            {
                If($member -notmatch 'Parent' -and $member -notmatch 'Cell' -and $member -notmatch 'Columns' -and $member -notmatch 'Rows')
                {
                    If($member.TypeName -match '__' -and $member.TypeName -notmatch '57da806104b8')
                    {
                        If($member -match $Term)
                        {
                            Write-Host "ProgID: $o"
                            "$CLSID.$member"
                        } 
                        $membernames = $member.Name       
                        ForEach($membername in $membernames)
                        {
                            $m1members = $com.$membername | gm   
                            ForEach($m1member in $m1members)
                            {
                                If($m1member -notmatch 'Parent' -and $m1member -notmatch 'Cell' -and $m1member -notmatch 'Columns' -and $m1member -notmatch 'Rows')
                                {
                                    If($m1member.TypeName -match '__')
                                    { 
                                        If($m1member -match $Term)
                                         {
                                           Write-Host "ProgID: $o"
                                          "$CLSID.$member.$m1member"
                                         } 
                                        $m1membernames = $m1member.Name 
                                        ForEach($m1membername in $m1membernames)
                                        {
                                            $m2members = $com.$membername.$m1membername | gm
                                            ForEach($m2member in $m2members)
                                            {
                                                If($m2member -notmatch 'Parent' -and $m2member -notmatch 'Cell' -and $m2member -notmatch 'Columns' -and $m2member -notmatch 'Rows')
                                                {
                                                    If($m2member.TypeName -match '__')
                                                    {
                                                        If($m2member -match $Term)
                                                        {
                                                              Write-Host "ProgID: $o"
                                                             "$CLSID.$member.$m1member.$m2member"
                                                        } 
                                                    }
                                                }   
                                            }              
                                        }  
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    } 
}
}
