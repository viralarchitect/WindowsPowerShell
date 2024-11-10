[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false)]
    [string]$SubReddit = 'news',

    [Parameter(Mandatory = $false)]
    [string]$SortBy = 'hot',

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$Limit = 100
)

# Function to remove duplicate entries from the resulting subreddit Excel files.
Function Remove-DuplicateRowsFromExcel {
    param(
        [Parameter(Mandatory = $false)]
        [string]$fileDir = '.\',

        [Parameter(Mandatory = $false)]
        [string]$fileName = 'r_[CONSOLIDATED].xlsx'
    )

    $filePath = Join-Path $fileDir $fileName

    if (-not (Test-Path $filePath)) {
        Write-Error "$filePath could not be found"
        return $null
    } else {
        Write-Debug "$filePath exists"
    }

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($filePath)

        $changesMade = $false

        foreach ($worksheet in $workbook.Worksheets) {
            # Process each worksheet

            $usedRange = $worksheet.UsedRange
            $lastRow = $usedRange.Rows.Count
            $lastColumn = $usedRange.Columns.Count

            if ($lastRow -le 1) {
                # Only headers, nothing to process
                continue
            }

            # Read data including headers
            $dataRange = $worksheet.Range("A1", $worksheet.Cells.Item($lastRow, $lastColumn))
            $dataValues = $dataRange.Value2

            # Read headers
            $headers = @()
            for ($j = 1; $j -le $dataValues.GetUpperBound(1); $j++) {
                $headers += $dataValues[1,$j]
            }

            # Read data rows
            $dataObjects = @()

            for ($i = 2; $i -le $dataValues.GetUpperBound(0); $i++) {
                $obj = [PSCustomObject]@{}

                for ($j = 1; $j -le $dataValues.GetUpperBound(1); $j++) {
                    $header = $headers[$j - 1]
                    $value = $dataValues[$i,$j]
                    $obj | Add-Member -NotePropertyName $header -NotePropertyValue $value
                }
                $dataObjects += $obj
            }

            $originalCount = $dataObjects.Count

            # Group data by Title
            $groups = $dataObjects | Group-Object -Property Title

            $uniqueDataObjects = @()

            foreach ($group in $groups) {
                $items = $group.Group

                # Convert Posted to DateTime and Upvotes to Int32
                foreach ($item in $items) {
                    if (-not ($item.Posted -is [DateTime])) {
                        $item.Posted = [DateTime]::Parse($item.Posted)
                    }
                    if (-not ($item.Upvotes -is [Int32])) {
                        $item.Upvotes = [int]$item.Upvotes
                    }
                }

                if ($items.Count -eq 1) {
                    $uniqueDataObjects += $items
                } else {
                    # Find maximum Posted date
                    $maxPostedDate = ($items | Measure-Object -Property Posted -Maximum).Maximum

                    $itemsWithMaxPostedDate = $items | Where-Object { $_.Posted -eq $maxPostedDate }

                    if ($itemsWithMaxPostedDate.Count -eq 1) {
                        $uniqueDataObjects += $itemsWithMaxPostedDate
                    } else {
                        # Find maximum Upvotes
                        $maxUpvotes = ($itemsWithMaxPostedDate | Measure-Object -Property Upvotes -Maximum).Maximum

                        $itemsWithMaxUpvotes = $itemsWithMaxPostedDate | Where-Object { $_.Upvotes -eq $maxUpvotes }

                        if ($itemsWithMaxUpvotes.Count -eq 1) {
                            $uniqueDataObjects += $itemsWithMaxUpvotes
                        } else {
                            # Sort by Subreddit ascending and take first one
                            $sortedItems = $itemsWithMaxUpvotes | Sort-Object -Property Subreddit
                            $selectedItem = $sortedItems[0]
                            $uniqueDataObjects += $selectedItem
                        }
                    }
                }
            }

            $uniqueCount = $uniqueDataObjects.Count

            if ($uniqueCount -lt $originalCount) {
                $changesMade = $true
            }

            # Clear existing data (excluding headers)
            $worksheet.Range("A2", $worksheet.Cells.Item($lastRow, $lastColumn)).ClearContents()

            # Write unique data back to worksheet
            $rows = $uniqueDataObjects.Count
            $cols = $lastColumn

            # Create multidimensional array
            $multidimArray = [object[,] ]::new($rows, $cols)

            for ($i = 0; $i -lt $rows; $i++) {
                $item = $uniqueDataObjects[$i]
                for ($j = 0; $j -lt $cols; $j++) {
                    $header = $headers[$j]
                    $multidimArray[$i,$j] = $item.$header
                }
            }

            # Write data to worksheet
            $writeRange = $worksheet.Range("A2").Resize($rows, $cols)
            $writeRange.Value2 = $multidimArray
        }

        $workbook.Save()
        $workbook.Close()
        $excel.Quit()

        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        if ($changesMade) {
            return $true
        } else {
            return $false
        }

    } catch {
        Write-Error "An error occurred: $_"
        return $null
    }
}

# Get the directory of the current script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Path to the JSON config file
$configFile = Join-Path $scriptDir 'reddit.json'

# Read and parse the JSON file
$config = Get-Content $configFile | ConvertFrom-Json

# Assign the values to variables
$clientID = $config.clientID
$clientSecret = $config.clientSecret
$userAgent = 'powershell-scripts/1.0 by viralarchitect'

# Encode the credentials for Basic Authentication
$authInfo = "$clientID`:$clientSecret"
$authBytes = [System.Text.Encoding]::UTF8.GetBytes($authInfo)
$authEncoded = [System.Convert]::ToBase64String($authBytes)

# Obtain an access token from Reddit
$AuthProgressId = 1
Write-Progress -Id $AuthProgressId -Activity "Authenticating with Reddit" -Status "Requesting access token" -PercentComplete 0

$tokenResponse = Invoke-RestMethod -Uri 'https://www.reddit.com/api/v1/access_token' `
    -Method POST `
    -Headers @{
        'Authorization' = "Basic $authEncoded"
        'User-Agent'    = $userAgent
    } `
    -Body @{
        'grant_type' = 'client_credentials'
    }

$accessToken = $tokenResponse.access_token

Write-Progress -Id $AuthProgressId -Activity "Authenticating with Reddit" -Completed

# Function to validate the existence of the subreddit
function Validate-Subreddit {
    param(
        [string]$subReddit,
        [string]$accessToken,
        [string]$userAgent
    )

    $uri = "https://oauth.reddit.com/r/$subReddit/about"
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers @{
            'Authorization' = "Bearer $accessToken"
            'User-Agent'    = $userAgent
        }
        return $true
    }
    catch {
        $errorResponse = $_.ErrorDetails.Message | ConvertFrom-Json
        if ($errorResponse.error -eq 404) {
            Write-Error "The subreddit r/$subReddit does not exist."
        } elseif ($errorResponse.error -eq 403) {
            Write-Error "Access to r/$subReddit is forbidden."
        } else {
            Write-Error "Failed to validate subreddit r/$subReddit. Error: $($_.Exception.Message)"
        }
        return $false
    }
}

# Validate the subreddit
$IsValidSubreddit = Validate-Subreddit -subReddit $SubReddit -accessToken $accessToken -userAgent $userAgent
if (-not $IsValidSubreddit) {
    exit
}

# Function to fetch and format top comments
function Get-TopComments {
    param (
        [string]$permalink,
        [string]$accessToken,
        [string]$userAgent,
        [int]$limit = 10,
        [int]$ParentProgressId = 0
    )

    $commentsUri = "https://oauth.reddit.com$($permalink)?limit=$limit&sort=top"

    # Wait to respect rate limits
    Start-Sleep -Milliseconds 1000

    $commentsProgressId = $ParentProgressId + 1

    # Initialize progress for fetching comments
    Write-Progress -Id $commentsProgressId -ParentId $ParentProgressId -Activity "Fetching comments" -Status "Fetching comments for post at $permalink" -PercentComplete 0

    try {
        $commentsResponse = Invoke-RestMethod -Uri $commentsUri `
            -Headers @{
                'Authorization' = "Bearer $accessToken"
                'User-Agent'    = $userAgent
            }
    }
    catch {
        Write-Warning "Failed to fetch comments for post at $permalink. Error: $($_.Exception.Message)"
        Write-Progress -Id $commentsProgressId -ParentId $ParentProgressId -Activity "Fetching comments" -Completed
        return ""
    }

    $comments = $commentsResponse[1].data.children

    $topComments = $comments | Where-Object { $_.kind -eq 't1' } | Select-Object -First $limit

    $commentsArray = @()
    $i = 1
    $TotalComments = $topComments.Count

    foreach ($comment in $topComments) {
        $percentComplete = ($i / $TotalComments) * 100
        Write-Progress -Id $commentsProgressId -ParentId $ParentProgressId -Activity "Processing comments" -Status "Processing comment $i of $TotalComments" -PercentComplete $percentComplete

        $commentData = $comment.data
        $author = $commentData.author
        $body = $commentData.body

        # Format each comment
        $commentsArray += "Comment $($i): by $author - $body"
        $i++
    }

    # Join the comments with two newlines
    $commentsText = $commentsArray -join "`n`n"

    # Clear the progress bar for comments
    Write-Progress -Id $commentsProgressId -ParentId $ParentProgressId -Activity "Processing comments" -Completed

    return $commentsText
}

# Fetch $Limit of the $sortBy posts from r/$SubReddit
$FetchPostsProgressId = 2
Write-Progress -Id $FetchPostsProgressId -Activity "Fetching posts from r/$SubReddit" -Status "Fetching posts" -PercentComplete 0

$Uri = 'https://oauth.reddit.com/r/' + $SubReddit + '/' + $sortBy + '?limit=' + $Limit
try {
    $response = Invoke-RestMethod -Uri $Uri `
        -Headers @{
            'Authorization' = "Bearer $accessToken"
            'User-Agent'    = $userAgent
        }
}
catch {
    Write-Error "Failed to fetch posts from r/$SubReddit. Error: $($_.Exception.Message)"
    Write-Progress -Id $FetchPostsProgressId -Activity "Fetching posts from r/$SubReddit" -Completed
    exit
}

Write-Progress -Id $FetchPostsProgressId -Activity "Fetching posts from r/$SubReddit" -Completed

# Prepare the data for Excel
$TotalPosts = $response.data.children.Count
$PostIndex = 0
$ProcessPostsProgressId = 3

$data = $response.data.children | ForEach-Object {
    $PostIndex++
    $percentComplete = ($PostIndex / $TotalPosts) * 100

    $post = $_.data
    Write-Progress -Id $ProcessPostsProgressId -Activity "Processing posts" -Status "Processing post $PostIndex of $TotalPosts" -PercentComplete $percentComplete

    # Debugging output
    Write-Host "Processing Post ID: $($post.id), Permalink: $($post.permalink)"

    # Fetch and format the top 10 comments for this post
    $commentsText = Get-TopComments -permalink $post.permalink -accessToken $accessToken -userAgent $userAgent -limit 10 -ParentProgressId $ProcessPostsProgressId

    [PSCustomObject]@{
        'Posted'              = [DateTimeOffset]::FromUnixTimeSeconds($post.created_utc).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
        'Subreddit'           = $post.subreddit_name_prefixed
        'Title'               = $post.title
        'Upvotes'             = $post.ups
        'External URL'        = if ($post.url -match '^/r/') { "https://www.reddit.com$($post.url)" } else { $post.url }
        'Reddit Comments URL' = "https://www.reddit.com$($post.permalink)"
        'Comments'            = $commentsText  # Add the comments here
    }
}

# Clear the progress bar for posts
Write-Progress -Id $ProcessPostsProgressId -Activity "Processing posts" -Completed

# Export the data to an Excel file
$ExportProgressId = 4
Write-Progress -Id $ExportProgressId -Activity "Exporting data to Excel" -Status "Exporting data" -PercentComplete 0

$data | Export-Excel -Path "r_$SubReddit.xlsx" -WorksheetName $sortBy -AutoSize -Append
$data | Export-Excel -Path "r_[CONSOLIDATED].xlsx" -WorksheetName $sortBy -AutoSize -Append

Write-Progress -Id $ExportProgressId -Activity "Exporting data to Excel" -Completed

Write-Host "Data successfully exported to r_$SubReddit.xlsx"

$result = Remove-DuplicateRowsFromExcel
if ($result -eq $true) {
    Write-Host "Cleanup Consolidated - Duplicates were found and removed."
} elseif ($result -eq $false) {
    Write-Host "No duplicates were found. No changes were made."
} elseif ($result -eq $null) {
    Write-Host "The file could not be found or accessed."
} else {
    Write-Host "An unexpected result was returned: $result"
}