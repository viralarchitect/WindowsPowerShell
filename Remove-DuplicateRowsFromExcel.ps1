Function Remove-DuplicateRowsFromExcel {
    [CmdletBinding()]
    param()

    $fileName = "r_[CONSOLIDATED].xlsx"
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $filePath = Join-Path $scriptPath $fileName

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
