function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter()]
        [System.String[]]
        $SenderDomainIs,

        [Parameter()]
        [System.String]
        $RejectMessageEnhancedStatusCode,

        [Parameter()]
        [System.String]
        $RejectMessageReasonText,

        [Parameter()]
        [System.String[]]
        $RecipientAddressContainsWords,
        
        [Parameter()]
        [System.String[]]
        $Comments,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure ="Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword
    )

    Write-Verbose -Message "Getting configuration of Transport Rule for $Name"
    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace("MSFT_", "")
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $ResourceName)
    $data.Add("Method", $MyInvocation.MyCommand)
    $data.Add("Principal", $GlobalAdminAccount.UserName)
    $data.Add("TenantId", $TenantId)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $ConnectionMode = New-M365DSCConnection -Platform 'ExchangeOnline' `
        -InboundParameters $PSBoundParameters

    $nullReturn = $PSBoundParameters
    $nullReturn.Ensure = "Absent"

    try
    {
        if ($null -eq (Get-Command 'Get-TransportRule' -ErrorAction SilentlyContinue))
        {
            return $nullReturn
        }
        $TransportRules = Get-TransportRule -ErrorAction Stop
        $TransportRule = $TransportRules | Where-Object -FilterScript { $_.Name -eq $Name }

        if ($null -eq $TransportRule)
        {
            Write-Verbose -Message "Transport Rule $($Name) does not exist."
            return $nullReturn
        }
        else
        {
            $result = @{
                Name                                 = $Name
                SenderDomainIs                       = $TransportRule.SenderDomainIs
                RejectMessageEnhancedStatusCode      = $TransportRule.RejectMessageEnhancedStatusCode
                RejectMessageReasonText              = $TransportRule.RejectMessageReasonText
                RecipientAddressContainsWords        = $TransportRule.RecipientAddressContainsWords
                Comments                             = $TransportRule.Comments
                Ensure                               = 'Present'
                GlobalAdminAccount                   = $GlobalAdminAccount
            }

            Write-Verbose -Message "Found Transport Rule $($Name)"
            Write-Verbose -Message "Get-TargetResource Result: `n $(Convert-M365DscHashtableToString -Hashtable $result)"
            return $result
        }
    }
    catch
    {
        Write-Verbose -Message $_
        Add-M365DSCEvent -Message $_ -EntryType 'Error' `
            -EventID 1 -Source $($MyInvocation.MyCommand.Source)
        return $nullReturn
    }
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter()]
        [System.String[]]
        $SenderDomainIs,

        [Parameter()]
        [System.String]
        $RejectMessageEnhancedStatusCode,

        [Parameter()]
        [System.String]
        $RejectMessageReasonText,

        [Parameter()]
        [System.String[]]
        $RecipientAddressContainsWords,
        
        [Parameter()]
        [System.String[]]
        $Comments,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure ="Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword
    )

    Write-Verbose -Message "Setting Transport Rule configuration for $Name"

    $currentTransportRuleConfig = Get-TargetResource @PSBoundParameters

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace("MSFT_", "")
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $ResourceName)
    $data.Add("Method", $MyInvocation.MyCommand)
    $data.Add("Principal", $GlobalAdminAccount.UserName)
    $data.Add("TenantId", $TenantId)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    if ($Global:CurrentModeIsExport)
    {
        $ConnectionMode = New-M365DSCConnection -Platform 'ExchangeOnline' `
            -InboundParameters $PSBoundParameters `
            -SkipModuleReload $true
    }
    else
    {
        $ConnectionMode = New-M365DSCConnection -Platform 'ExchangeOnline' `
            -InboundParameters $PSBoundParameters
    }


        $NewTransportRuleParams = @{
            Name                                 = $Name
            SenderDomainIs                       = $SenderDomainIs
            RejectMessageEnhancedStatusCode      = $RejectMessageEnhancedStatusCode
            RejectMessageReasonText              = $RejectMessageReasonText
            RecipientAddressContainsWords        = $RecipientAddressContainsWords
            Comments                             = $Comments
            Ensure                               = $Ensure
        }

   
    #Transport Rule doesn't exist but it should
    if ($Ensure -eq "Present" -and $currentTransportRuleConfig.Ensure -eq "Absent")
    {
        Write-Verbose -Message "Transport Rule '$($Name)' does not exist but it should. Creating Transport Rule."
        New-TransportRule @NewTransportRuleParams
    }
    #Transport Rule exists but shouldn't
    elseif ($Ensure -eq "Absent" -and $currentTransportRuleConfig.Ensure -eq "Present")
    {
        Write-Verbose -Message "Transport Rule '$($Name)' exists but shouldn't. Removing this Transport Rule."
        Remove-TransportRule -Identity $Name -Confirm:$false
    }
    elseif ($Ensure -eq "Present" -and $currentTransportRuleConfig.Ensure -eq "Present")
    {
        Write-Verbose -Message "Transport Rule '$($Name)' already exists. Updating settings"
        Write-Verbose -Message "Setting Transport Rule '$($Name)' with values: $(Convert-M365DscHashtableToString -Hashtable $NewTransportRuleParams)"
        Set-TransportRule $Name @NewTransportRuleParams
    }

}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Name,

        [Parameter()]
        [System.String[]]
        $SenderDomainIs,

        [Parameter()]
        [System.String]
        $RejectMessageEnhancedStatusCode,

        [Parameter()]
        [System.String]
        $RejectMessageReasonText,

        [Parameter()]
        [System.String[]]
        $RecipientAddressContainsWords,
        
        [Parameter()]
        [System.String[]]
        $Comments,

        [Parameter()]
        [ValidateSet("Present", "Absent")]
        [System.String]
        $Ensure ="Present",

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword
    )
    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace("MSFT_", "")
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $ResourceName)
    $data.Add("Method", $MyInvocation.MyCommand)
    $data.Add("Principal", $GlobalAdminAccount.UserName)
    $data.Add("TenantId", $TenantId)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    Write-Verbose -Message "Testing Transport Rule configuration for $Name"

    $CurrentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-M365DscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-M365DscHashtableToString -Hashtable $PSBoundParameters)"

    $ValuesToCheck = $PSBoundParameters
    $ValuesToCheck.Remove('GlobalAdminAccount') | Out-Null

    $TestResult = Test-M365DSCParameterState -CurrentValues $CurrentValues `
        -Source $($MyInvocation.MyCommand.Source) `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck $ValuesToCheck.Keys

    Write-Verbose -Message "Test-TargetResource returned $TestResult"

    return $TestResult
}

function Export-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $GlobalAdminAccount,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword
    )

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace("MSFT_", "")
    $data = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
    $data.Add("Resource", $ResourceName)
    $data.Add("Method", $MyInvocation.MyCommand)
    $data.Add("Principal", $GlobalAdminAccount.UserName)
    $data.Add("TenantId", $TenantId)
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $ConnectionMode = New-M365DSCConnection -Platform 'ExchangeOnline' `
        -InboundParameters $PSBoundParameters `
        -SkipModuleReload $true
    try
    {
        $dscContent = ""
        [array]$transportRules = Get-TransportRule -ErrorAction Stop
        if ($transportRules.Length -eq 0)
        {
            Write-Host $Global:M365DSCEmojiGreenCheckMark
        }
        else
        {
            Write-Host "`r`n" -NoNewLine
        }
        $i = 1

        foreach ($transportRule in $transportRules)
        {
            Write-Host "    |---[$i/$($transportRules.Count)] $($transportRule.Name)" -NoNewLine
            $params = @{
                Name                  = $transportRule.Name
                GlobalAdminAccount    = $GlobalAdminAccount
                ApplicationId         = $ApplicationId
                TenantId              = $TenantId
                CertificateThumbprint = $CertificateThumbprint
                CertificatePassword   = $CertificatePassword
                CertificatePath       = $CertificatePath
            }
            $Results = Get-TargetResource @Params
            $Results = Update-M365DSCExportAuthenticationResults -ConnectionMode $ConnectionMode `
                -Results $Results
            $dscContent += Get-M365DSCExportContentForResource -ResourceName $ResourceName `
                -ConnectionMode $ConnectionMode `
                -ModulePath $PSScriptRoot `
                -Results $Results `
                -GlobalAdminAccount $GlobalAdminAccount
            Write-Host $Global:M365DSCEmojiGreenCheckMark
            $i ++
        }
        return $dscContent
    }
    catch
    {
        Write-Verbose -Message $_
        Add-M365DSCEvent -Message $_ -EntryType 'Error' `
            -EventID 1 -Source $($MyInvocation.MyCommand.Source)
        return ""
    }
}

Export-ModuleMember -Function *-TargetResource
