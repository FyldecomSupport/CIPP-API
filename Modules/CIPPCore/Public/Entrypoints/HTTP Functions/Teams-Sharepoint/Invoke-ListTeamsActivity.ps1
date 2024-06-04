using namespace System.Net

Function Invoke-ListTeamsActivity {
    <#
    .FUNCTIONALITY
        Entrypoint
    .ROLE
        Teams.Activity.Read
    #>
    [CmdletBinding()]
    param($Request, $TriggerMetadata)


    # Write to the Azure Functions log stream.
    Write-Host 'PowerShell HTTP trigger function processed a request.'

    # Interact with query parameters or the body of the request.
    $TenantFilter = $Request.Query.TenantFilter
    $type = $request.query.Type
    $GraphRequest = New-GraphGetRequest -uri "https://graph.microsoft.com/beta/reports/get$($type)Detail(period='D7')" -tenantid $TenantFilter | ConvertFrom-Csv | Select-Object @{ Name = 'UPN'; Expression = { $_.'User Principal Name' } },
    @{ Name = 'LastActive'; Expression = { $_.'Last Activity Date' } },
    @{ Name = 'TeamsChat'; Expression = { $_.'Team Chat Message Count' } },
    @{ Name = 'CallCount'; Expression = { $_.'Call Count' } },
    @{ Name = 'MeetingCount'; Expression = { $_.'Meeting Count' } },
    @{ Name = 'AudioCallDuration'; Expression = { $_.'Audio Duration' } },
    @{ Name = 'VideoCallDuration'; Expression = { $_.'Video Duration' } },
    @{ Name = 'ScreenShareDuration'; Expression = { $_.'Screen Share Duration' } },
    @{ Name = 'PrivateChat'; Expression = { $_.'Private Chat Message Count' } }
                         
                        

    # Associate values to output bindings by calling 'Push-OutputBinding'.
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = @($GraphRequest)
        })

}
