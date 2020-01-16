# Globals
# This script assumes that the NETBIOS name of your intended FQDN is the subdomain of the domain below.
# When I read your current NETBIOS\USERNAME format, it will be re-written to USERNAME@NETBIOS.<$GLOB_DESIRED_DOMAIN>
$GLOB_DESIRED_DOMAIN = "domain.edu"

# NETBIOS Name Expression for subdomains you want to match (maybe your email and upn domains are different). Wildcard accepted.
$GLOB_DESIRED_DOMAIN_SHORT = "NETBIOS*"

# Search-Registry Function from TechNet
function Search-Registry { 
        [CmdletBinding()] 
        param( 
            [Parameter(Mandatory, Position=0, ValueFromPipelineByPropertyName)] 
            [Alias("PsPath")] 
            [string[]] $Path, 
            [switch] $Recurse, 
            [Parameter(ParameterSetName="SingleSearchString", Mandatory)] 
            [string] $SearchRegex, 
            [Parameter(ParameterSetName="SingleSearchString")] 
            [switch] $KeyName, 
            [Parameter(ParameterSetName="SingleSearchString")] 
            [switch] $ValueName, 
            [Parameter(ParameterSetName="SingleSearchString")] 
            [switch] $ValueData, 
            [Parameter(ParameterSetName="MultipleSearchStrings")] 
            [string] $KeyNameRegex, 
            [Parameter(ParameterSetName="MultipleSearchStrings")] 
            [string] $ValueNameRegex, 
            [Parameter(ParameterSetName="MultipleSearchStrings")] 
            [string] $ValueDataRegex 
        ) 
     
        begin { 
            switch ($PSCmdlet.ParameterSetName) { 
                SingleSearchString { 
                    $NoSwitchesSpecified = -not ($PSBoundParameters.ContainsKey("KeyName") -or $PSBoundParameters.ContainsKey("ValueName") -or $PSBoundParameters.ContainsKey("ValueData")) 
                    if ($KeyName -or $NoSwitchesSpecified) { $KeyNameRegex = $SearchRegex } 
                    if ($ValueName -or $NoSwitchesSpecified) { $ValueNameRegex = $SearchRegex } 
                    if ($ValueData -or $NoSwitchesSpecified) { $ValueDataRegex = $SearchRegex } 
                } 
                MultipleSearchStrings { 
                } 
            } 
        } 
     
        process { 
            foreach ($CurrentPath in $Path) { 
                Get-ChildItem $CurrentPath -Recurse:$Recurse |  
                    ForEach-Object { 
                        $Key = $_ 
     
                        if ($KeyNameRegex) {  
                            Write-Verbose ("{0}: Checking KeyNamesRegex" -f $Key.Name)  
             
                            if ($Key.PSChildName -match $KeyNameRegex) {  
                                Write-Verbose "  -> Match found!" 
                                return [PSCustomObject] @{ 
                                    Key = $Key 
                                    Reason = "KeyName" 
                                } 
                            }  
                        } 
             
                        if ($ValueNameRegex) {  
                            Write-Verbose ("{0}: Checking ValueNamesRegex" -f $Key.Name) 
                 
                            if ($Key.GetValueNames() -match $ValueNameRegex) {  
                                Write-Verbose "  -> Match found!" 
                                return [PSCustomObject] @{ 
                                    Key = $Key 
                                    Reason = "ValueName" 
                                } 
                            }  
                        } 
             
                        if ($ValueDataRegex) {  
                            Write-Verbose ("{0}: Checking ValueDataRegex" -f $Key.Name) 
                 
                            if (($Key.GetValueNames() | % { $Key.GetValue($_) }) -match $ValueDataRegex) {  
                                Write-Verbose "  -> Match!" 
                                return [PSCustomObject] @{ 
                                    Key = $Key 
                                    Reason = "ValueData" 
                                } 
                            } 
                        } 
                    } 
            } 
        } 
    } 


# Lets circle through each key and check to see if the principal links us to a key that contains a value called "001f3d16"
$results = Search-Registry -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" -Recurse -SearchRegex "001f3d16" -ValueName

clear

Foreach ( $foundreg in $results ) {

    Write-Host "Evaulating $($foundreg.key)"

    $foundreg.key = $foundreg.key -replace "HKEY_CURRENT_USER","HKCU:"

    # Once we have our magic Provider GUID, lets read the value from binary into a string.
    $string_before = [System.Text.Encoding]::Unicode.GetString((Get-ItemProperty -Path $foundreg.key -Name "001f3d16")."001f3d16")
    write-host "$($string_before) was found in $($foundreg.key)"

    # If that string contains an Symbol, no work to do - exit.
    if ($string_before -like '*@*') {
    
            ####### UNCOMMENT THESE LINES ONLY FOR TESTING - SETS FIELD TO BAD DATA TO TRIGGER SCRIPT ACTION #####
            #$unicode = [System.Text.Encoding]::Unicode.GetBytes('domain\user')|%{[System.Convert]::ToString($_,16)}
            #$hex_ready = $unicode.Split(',') | % { "0x$_"}
            #Write-Host "Intentionally Breaking user_hint to NETBIOS format so script can re-run." -ForegroundColor red
            #Set-ItemProperty -Path "$($foundreg.key)" -Name "001f3d16" -Value ([byte[]]($hex_ready))
            ######


    Write-Host "Found @ in value which indicates FQDN: $($string_before), Exiting." -ForegroundColor yellow


    # If that string contains domain\, split() the string on the \ and extract username into $extract_username
    } elseif ($string_before -like "$($GLOB_DESIRED_DOMAIN_SHORT)\*") {

        # Get the username from string after backslash
        $netbios = $string_before.Split('\\')[0].trim([char]$null)
        $username = $string_before.Split('\\')[1].trim([char]$null)
        Write-Host "Converting $($string_before) to UPN: $($username)@$($netbios).$($GLOB_DESIRED_DOMAIN)" -ForegroundColor green

        # Convert to Unicode Stream Object
        $unicode = [System.Text.Encoding]::Unicode.GetBytes("$($username)@$($netbios).$($GLOB_DESIRED_DOMAIN)")|%{[System.Convert]::ToString($_,16)}

        # Convert Unicode Stream into comma seperated HEX (as expected by registry).
        $hex_ready = $unicode.Split(',') | % { "0x$_"}

        # Re-insert a concat'ed $extract_username@$domainstring in HEX
        Set-ItemProperty -Path "$($foundreg.key)" -Name "001f3d16" -Value ([byte[]]($hex_ready))


    } else {

        Write-Host "The value does not match expected format! Exiting." -ForegroundColor red

    }

Write-Host "Completed an Evaluation Cycle"
Write-Host "`n"
}
