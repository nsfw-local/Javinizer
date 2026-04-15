function Get-MgstageId {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        $id = (((($Webrequest.Content -split '<th>品番：<\/th>')[1] -split '<\/td>')[0]) -split '<td>')[1]

        if ($id -eq '') {
            $id = $null
        }

        Write-Output $Id
    }
}

function Get-MgstageTitle {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        $title = (($Webrequest.Content -split '<title>')[1] -split '<\/title>')[0]
        $title = Convert-HtmlCharacter -String $title

        if ($title -eq '') {
            $title = $null
        }

        Write-Output $Title
    }
}

function Get-MgstageDescription {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        if ($Webrequest.Content -match '<p class="txt introduction">') {
            $description = (($Webrequest.Content -split '<p class="txt introduction">')[1] -split '<\/p>')[0]
            $description = Convert-HtmlCharacter -String $description
        } else {
            $description = $null
        }

        Write-Output $description
    }
}

function Get-MgstageReleaseDate {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        $releaseDate = (((($Webrequest.Content -split '<th>配信開始日：<\/th>')[1] -split '<\/td>')[0]) -split '<td>')[1]
        $releaseDate = Get-Date $releaseDate -Format "yyyy-MM-dd"

        Write-Output $releaseDate
    }
}

function Get-MgstageReleaseYear {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        $releaseYear = Get-MgstageReleaseDate -WebRequest $Webrequest
        $releaseYear = ($releaseYear -split '-')[0]

        Write-Output $releaseYear
    }
}

function Get-MgstageRuntime {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        $length = ((($Webrequest.Content -split '<th>収録時間：<\/th>')[1] -split '<\/td>')[0] -split '<td>')[1]
        $length = ($length -replace 'min').Trim()

        Write-Output $length
    }
}

function Get-MgstageMaker {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        $maker = (((($Webrequest.Content -split '<th>メーカー：<\/th>')[1] -split '<\/td>')[0] -split '>')[2] -split '<\/a')[0]
        $maker = Convert-HtmlCharacter -String $maker

        if ($maker -eq '') {
            $maker = $null
        }

        Write-Output $maker
    }
}

function Get-MgstageLabel {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest,

        [Parameter()]
        [Object]$Replace
    )

    process {
        $label = (((($Webrequest.Content -split '<th>レーベル：<\/th>')[1] -split '<\/td>')[0] -split '>')[2] -split '<\/a')[0]
        $label = Convert-HtmlCharacter -String $label

        if ($label -eq '') {
            $label = $null
        }

        Write-Output $label
    }
}

function Get-MgstageSeries {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest,

        [Parameter()]
        [Object]$Replace
    )

    process {
        $series = (((($Webrequest.Content -split '<th>シリーズ：<\/th>')[1] -split '<\/td>')[0] -split '>')[2] -split '<\/a')[0]
        $series = Convert-HtmlCharacter -String $series

        if ($series -eq '') {
            $series = $null
        }

        Write-Output $series
    }
}

function Get-MgstageRating {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        try {
            $rating = ($Webrequest.Content | Select-String -Pattern '<span class="star_.*"><\/span>(.*)').Matches.Groups[1].Value
            $ratingCount = (($Webrequest.Content | Select-String -Pattern '\((\d*) 件\)').Matches.Groups[1].Value).ToString()
        } catch {
            return
        }

        # Multiply the rating value by 2 to conform to 1-10 rating standard
        $newRating = [Decimal]$rating * 2
        $newRating = [Math]::Round($newRating, 2)

        if ($newRating -eq 0) {
            $rating = $null
        } else {
            $rating = $newRating.ToString()
        }

        if ($ratingCount -eq 0) {
            $ratingObject = $null
        } else {
            $ratingObject = [PSCustomObject]@{
                Rating = $rating
                Votes  = $ratingCount
            }
        }

        Write-Output $ratingObject
    }
}

function Get-MgstageGenre {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        $genreArray = @()
        $genreHtml = ((($Webrequest.Content -split '<th>ジャンル：<\/th>')[1] -split '<\/td>')[0]) -split '<a href="\/search\/csearch\.php\?genre\[\]=.*">' | ForEach-Object { ($_ -replace '<td>' -replace '<\/a>').Trim() } | Where-Object { $_ -ne '' }

        foreach ($genre in $genreHtml) {
            $genre = Convert-HtmlCharacter -String $genre
            $genreArray += $genre
        }

        if ($genreArray.Count -eq 0) {
            $genreArray = $null
        }

        Write-Output $genreArray
    }
}

function Get-MgstageActress {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest,

        [Parameter(Mandatory = $false)]
        [String]$Id = ''
    )

    process {
        $movieActressObject = @()

        if ($Id -ne '') {
            try {
                # Step 1: Search shiroutoname.com
                $searchUrl = "https://shiroutoname.com/?s=$Id"
                Write-JVLog -Write:$script:JVLogWrite -LogPath $script:JVLogPath -WriteLevel $script:JVLogWriteLevel -Level Debug -Message "[$Id] [$($MyInvocation.MyCommand.Name)] Fetching actress from [$searchUrl]"
                $searchRequest = Invoke-WebRequest -Uri $searchUrl -Method Get -Verbose:$false

                # Step 2: Restrict to FIRST article only to avoid false positives from related articles
                $firstArticle = [regex]::Match($searchRequest.Content,
                    '<article[^>]*class="flex[^"]*"[^>]*>(.*?)</article>',
                    [System.Text.RegularExpressions.RegexOptions]::Singleline)

                if ($firstArticle.Success) {
                    $articleBlock = $firstArticle.Groups[1].Value

                    # Get detaillink href to follow to detail page
                    $detailHref = ([regex]::Match($articleBlock, '<a[^>]*class="detaillink"[^>]*href="([^"]+)"')).Groups[1].Value

                    if ($detailHref) {
                        # Step 3: Detail page has actress inside <blockquote class="details"> - no false positives
                        Write-JVLog -Write:$script:JVLogWrite -LogPath $script:JVLogPath -WriteLevel $script:JVLogWriteLevel -Level Debug -Message "[$Id] [$($MyInvocation.MyCommand.Name)] Following detaillink [$detailHref]"
                        $detailRequest = Invoke-WebRequest -Uri $detailHref -Method Get -Verbose:$false

                        $detailsBlock = ([regex]::Match($detailRequest.Content,
                            '<blockquote[^>]*class="[^"]*details[^"]*"[^>]*>(.*?)</blockquote>',
                            [System.Text.RegularExpressions.RegexOptions]::Singleline)).Groups[1].Value

                        $actressSource = $detailsBlock
                    } else {
                        # No detaillink - parse actress from first article directly
                        $actressSource = $articleBlock
                    }

                    # Parse mlink (MGS links) only, skip purchase buttons via class="mlink tag sbuy"
                    $actressMatches = [regex]::Matches($actressSource,
                        '<a\s+[^>]*class="mlink"[^>]*>([^<]+)<')

                    $actressNames = $actressMatches |
                        ForEach-Object { $_.Groups[1].Value.Trim() } |
                        Where-Object { $_ -ne '' -and $_ -notmatch '^[A-Za-z]+$' -or ($_ -match '[\u3040-\u309f\u30a0-\u30ff\u4e00-\u9faf]') } |
                        Where-Object { $_ -notmatch 'MGS|FANZA' } |
                        Select-Object -Unique

                    Write-JVLog -Write:$script:JVLogWrite -LogPath $script:JVLogPath -WriteLevel $script:JVLogWriteLevel -Level Debug -Message "[$Id] [$($MyInvocation.MyCommand.Name)] shiroutoname results: [$($actressNames -join ', ')]"

                    foreach ($actressName in $actressNames) {
                        $movieActressObject += [PSCustomObject]@{
                            LastName     = $null
                            FirstName    = $null
                            JapaneseName = Convert-JVCleanString -String $actressName
                            ThumbUrl     = $null
                        }
                    }
                }
            } catch {
                Write-JVLog -Write:$script:JVLogWrite -LogPath $script:JVLogPath -WriteLevel $script:JVLogWriteLevel -Level Warning -Message "[$Id] [$($MyInvocation.MyCommand.Name)] Error fetching from shiroutoname: $PSItem" -Action 'Continue'
            }
        }

        # Fallback: parse actress directly from MGStage page
        if ($movieActressObject.Count -eq 0) {
            Write-JVLog -Write:$script:JVLogWrite -LogPath $script:JVLogPath -WriteLevel $script:JVLogWriteLevel -Level Debug -Message "[$Id] [$($MyInvocation.MyCommand.Name)] Falling back to MGStage actress parsing"
            $movieActress = (((($Webrequest.Content -split '<th>出演：<\/th>')[1] -split '<\/td>')[0]) -replace '<td>' -replace '<\/a>' -replace '<a href="\/search\/csearch\.php\?actor\[\]=.*">') -split '\n' `
            | ForEach-Object { ($_).Trim() } | Where-Object { $_ -ne '' }

            foreach ($actress in $movieActress) {
                if ($actress -match '[\u3040-\u309f]|[\u30a0-\u30ff]|[\uff66-\uff9f]|[\u4e00-\u9faf]') {
                    $movieActressObject += [PSCustomObject]@{
                        LastName     = $null
                        FirstName    = $null
                        JapaneseName = Convert-JVCleanString -String $actress
                        ThumbUrl     = $null
                    }
                } else {
                    $movieActressObject += [PSCustomObject]@{
                        LastName     = ($actress -split ' ')[1] -replace '\\', ''
                        FirstName    = ($actress -split ' ')[0] -replace '\\', ''
                        JapaneseName = $null
                        ThumbUrl     = $null
                    }
                }
            }
        }

        if ($movieActressObject.Count -eq 0) {
            $movieActressObject = $null
        }

        Write-Output $movieActressObject
    }
}

function Get-MgstageCoverUrl {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        try {
            #$coverUrl = ($Webrequest.Content | Select-String -Pattern '<img src="(.*)" width=".*" height=".*" class="enlarge_image"').Matches.Groups[1].Value
            $coverUrl = ($Webrequest.Content | Select-String -Pattern 'class="link_magnify" href="(.*\.jpg)"').Matches.Groups[1].Value
        } catch {
            return
        }

        Write-Output $coverUrl
    }
}

function Get-MgstageScreenshotUrl {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        try {
            $screenshotUrl = ( $Webrequest.Content | Select-String -Pattern 'class="sample_image" href="(.*.jpg)"' -AllMatches ).Matches | ForEach-Object { $_.Groups[1].Value }
        } catch {
            return
        }

        Write-Output $screenshotUrl
    }
}

function Get-MgstageTrailerUrl {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Object]$Webrequest
    )

    process {
        try {
            $trailerID = ($Webrequest.Content | Select-String -Pattern '\/sampleplayer\/sampleplayer.html\/([^"]+)"').Matches.Groups[1].Value
            $traileriFrameUrl = 'https://www.mgstage.com/sampleplayer/sampleRespons.php?pid=' + $trailerID
            $trailerUrl = ((Invoke-WebRequest -Uri $traileriFrameUrl -WebSession $session -Verbose:$false).Content | Select-String -Pattern '(https.+.ism)\\/').Matches.Groups[1].Value -replace '\\', '' -replace 'ism', 'mp4'
        } catch {
            return
        }
        Write-Output $trailerUrl
    }
}
