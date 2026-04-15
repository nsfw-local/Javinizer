function Get-MgstageUrl {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [String]$Id,

        [Parameter()]
        [Switch]$AllResults
    )

    begin {
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

        $cookie = New-Object System.Net.Cookie
        $cookie.Name = 'adc'
        $cookie.Value = '1'
        $cookie.Domain = '.mgstage.com'
        $session.Cookies.Add($cookie)
    }

    process {
        $searchUrl = "https://www.mgstage.com/search/cSearch.php?search_word=$Id"

        try {
            Write-JVLog -Write:$script:JVLogWrite -LogPath $script:JVLogPath -WriteLevel $script:JVLogWriteLevel -Level Debug -Message "[$Id] [$($MyInvocation.MyCommand.Name)] Performing [GET] on URL [$searchUrl]"
            $webRequest = Invoke-WebRequest -Uri $searchUrl -Method Get -WebSession $session -Verbose:$false
        } catch {
            try {
                Start-Sleep -Seconds 3
                $webRequest = Invoke-WebRequest -Uri $searchUrl -Method Get -WebSession $session -Verbose:$false
            } catch {
                Write-JVLog -Write:$script:JVLogWrite -LogPath $script:JVLogPath -WriteLevel $script:JVLogWriteLevel -Level Error -Message "[$Id] [$($MyInvocation.MyCommand.Name)] Error occured on [GET] on URL [$searchUrl]: $PSItem" -Action 'Continue'
            }
        }

        try {
            # FIX: MGStage changed HTML structure, no longer uses <p class="tag">
            # Now parsing directly from product detail links
            $resultObject = [regex]::Matches($webRequest.Content, '<a href="(/product/product_detail/([^/]+)/)"') |
                ForEach-Object {
                    [PSCustomObject]@{
                        Id  = $_.Groups[2].Value
                        Url = "https://www.mgstage.com" + $_.Groups[1].Value
                    }
                } | Sort-Object Id -Unique
        } catch {
            # Do nothing
        }

        if ($Id -in $resultObject.Id) {
            $matchedResult = $resultObject | Where-Object { $Id -eq $_.Id }

            if ($matchedResult.Count -gt 1 -and !($AllResults)) {
                $matchedResult = $matchedResult[0]
            }

            $urlObject = foreach ($entry in $matchedResult) {
                [PSCustomObject]@{
                    Ja    = if ($entry.Url[-1] -ne '/') { $entry.Url + '/' } else { $entry.Url }
                    Id    = $entry.Id
                    Title = $entry.Title
                }
            }

            Write-Output $urlObject
        } else {
            Write-JVLog -Write:$script:JVLogWrite -LogPath $script:JVLogPath -WriteLevel $script:JVLogWriteLevel -Level Warning -Message "[$Id] [$($MyInvocation.MyCommand.Name)] not matched on Mgstage"
            return
        }
    }
}