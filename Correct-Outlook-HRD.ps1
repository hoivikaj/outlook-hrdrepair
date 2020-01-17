
# Globals
# This script assumes that the NETBIOS name of your intended FQDN is the subdomain of the domain below.
# When I read your current NETBIOS\USERNAME format, it will be re-written to USERNAME@NETBIOS.<$GLOB_DESIRED_DOMAIN>
$GLOB_DESIRED_DOMAIN = "domain.edu"

# NETBIOS Name Regular Expression Match. Wildcard is appended to include prod and test domains in script scope.
$GLOB_DESIRED_DOMAIN_SHORT = "NETBIOS*"

# Stole this Search-Registry Function.
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


# Lets circle through each key and check to see if the principal links us to a key that contains a value called "001e6750" - This value appears in the key that we need to manage.
$searchresults = Search-Registry -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" -Recurse -SearchRegex "001e6750" -ValueName

clear

Foreach ( $iterative_reg in $searchresults ) {

    Write-Host "Current Iterative Key: "`n`n"$($iterative_reg.key)`n"
    $iterative_reg.key = $iterative_reg.key -replace "HKEY_CURRENT_USER","HKCU:"

    $keyobject = Get-ItemProperty -Path $iterative_reg.key

    # Check if the target value exists. If it does, check to make sure value is correct.
    if ( [bool]($keyobject.PSobject.Properties.name -match "001f3d16") ) {

        # Since we already have the value, it was most likely created with a legacy workflow. Lets check it for known formats and correct if neccessary.
       
        # Once we have our magic Provider GUID, lets read the value from binary into a string.
        $001f3d16_raw = [System.Text.Encoding]::Unicode.GetString((Get-ItemProperty -Path $iterative_reg.key -Name "001f3d16")."001f3d16")
        write-host "$($001f3d16_raw.trim([char]$null)) was found`n"

        # If that string contains an Symbol, no work to do - exit.
        if ($001f3d16_raw -like '*@*') {
    
            # <<<<<<<This is the desired state and is the end state on repeated runs.>>>>>>>
            Write-Host "Found @ in value which indicates DESIRED STATE: $($001f3d16_raw.trim([char]$null)), Exiting.`n" -ForegroundColor yellow
            

            ################# UNCOMMENT THESE LINES ONLY FOR TESTING - SETS FIELD TO BAD DATA TO TRIGGER SCRIPT ACTION #####
            #$001f3d16_unicode = [System.Text.Encoding]::Unicode.GetBytes('domain\username')|%{[System.Convert]::ToString($_,16)}
            #$hex_ready = $001f3d16_unicode.Split(',') | % { "0x$_"}
            #write-Host "Intentionally Breaking user_hint to NETBIOS format so script can re-run. `n" -ForegroundColor red
            #Set-ItemProperty -Path "$($iterative_reg.key)" -Name "001f3d16" -Value ([byte[]]($hex_ready))
            ################


        }

        # If that string contains UNCH\, split() the string on the \ and extract username into $extract_username
        elseif ($001f3d16_raw -like "$($GLOB_DESIRED_DOMAIN_SHORT)\*") {

            # Get the username from string after backslash
            $netbios = $001f3d16_raw.Split('\\')[0].trim([char]$null)
            $username = $001f3d16_raw.Split('\\')[1].trim([char]$null)
            Write-Host "Converting $($001f3d16_raw) to UPN: $($username)@$($netbios).$($GLOB_DESIRED_DOMAIN)`n" -ForegroundColor green

            # Convert to Unicode Stream Object
            $001f3d16_unicode = [System.Text.Encoding]::Unicode.GetBytes("$($username)@$($netbios).$($GLOB_DESIRED_DOMAIN)")|%{[System.Convert]::ToString($_,16)}

            # Convert Unicode Stream into comma seperated HEX (as expected by registry).
            $001f3d16_unicode = $001f3d16_unicode.Split(',') | % { "0x$_"}

            # Re-insert a concat'ed $extract_username@$domainstring in HEX
            Set-ItemProperty -Path "$($iterative_reg.key)" -Name "001f3d16" -Value ([byte[]]($001f3d16_unicode))

        } 
    
        # The value did not contain a UPN indicator OR a NETBIOS indicator.
        else {

            Write-Host "The value exists but does not match expected format! Exiting.`n" -ForegroundColor red

        }

    } 
    
    # The value does NOT Exist, so lets create it by using the value copied from "001f3001"
    else {

        # Get the value of the email key to use as our user_hint.
        $001f3001_source = Get-ItemProperty -Path $iterative_reg.key -Name "001f3001"
        New-ItemProperty -Path $iterative_reg.key -Name "001f3d16" -PropertyType Binary -Value $001f3001_source."001f3001" | out-null
        Write-Host "Duplicated email value to user_hint value`n" -ForegroundColor green

    }
    
    Write-Host "Completed an Evaluation Cycle`n"
}
