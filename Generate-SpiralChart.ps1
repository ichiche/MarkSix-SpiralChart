#Global Variable
$Global:RegressionDrawNumberAndYearArray = @()
$Global:SpiralChartArray = New-Object Object[] 49
$Global:RegressionNumberArray = @()
[int]$Global:MiddleNumber = 0
[int]$Global:AddNumberCount = 0

#Color
$Red_List = @(1,2,7,8,12,13,18,19,23,24,29,30,34,35,40,45,46)
$Blue_List = @(3,4,9,10,14,15,20,25,26,31,36,37,41,42,47,48)
$Green_List = @(5,6,11,16,17,21,22,27,28,32,33,38,39,43,44,49)

#Number
$NumberOfDrawIn2019 = 144
$NumberOfDrawIn2018 = 149
$NumberOfDrawIn2017 = 153
$NumberOfDrawIn2016 = 151
$NumberOfDrawIn2015 = 152
$NumberOfDrawIn2014 = 152
$NumberOfDrawIn2013 = 152
$NumberOfDrawIn2012 = 152

$NumberOfDraw = @{
"2019" = $NumberOfDrawIn2019;
"2018" = $NumberOfDrawIn2018;
"2017" = $NumberOfDrawIn2017;
"2016" = $NumberOfDrawIn2016;
"2015" = $NumberOfDrawIn2015;
"2014" = $NumberOfDrawIn2014;
"2013" = $NumberOfDrawIn2013;
"2012" = $NumberOfDrawIn2012
}

#Initialize Array
for ($i = 0; $i -lt $Global:SpiralChartArray.Count; $i++) {
    $Global:SpiralChartArray[$i] = 0;
}

function Reference-MiddleNumber {
param(
    [string]$DrawYear,
    [string]$ShortDrawNumber
)  
    #Local Variable
    $FullDrawResult = $Global:DrawResultMultipleYears | sort DrawNumber

    foreach ($i in $FullDrawResult) {
        if ($i.DrawYear -eq $DrawYear -and $i.ShortDrawNumber -eq $ShortDrawNumber) {
            $Global:PastDrawResult = @([int]$i.Number1,[int]$i.Number2,[int]$i.Number3,[int]$i.Number4,[int]$i.Number5,[int]$i.Number6,[int]$i.SpecialNumber)

            [int]$FirstNumber = $i.Number1
            [int]$LastNumber = $i.Number6
            [double]$SumOfFirstAndLastNumber = $FirstNumber + $LastNumber
            $Global:MiddleNumber = [math]::ceiling($SumOfFirstAndLastNumber / 2)
        }
    }
}

function Add-Number {
    if ($Global:AddNumberCount + 1 -le 49) {
        $Global:AddNumberCount++
        return $Global:AddNumberCount
    } else {
        $Global:AddNumberCount = 1
        return $Global:AddNumberCount
    }
}

function Generate-SpiralChartArray {
    $Global:AddNumberCount = $Global:MiddleNumber

    $Global:SpiralChartArray[24] = $Global:MiddleNumber
    $Global:SpiralChartArray[25] = Add-Number
    $Global:SpiralChartArray[18] = Add-Number
    $Global:SpiralChartArray[17] = Add-Number
    $Global:SpiralChartArray[16] = Add-Number
    $Global:SpiralChartArray[23] = Add-Number
    $Global:SpiralChartArray[30] = Add-Number
    $Global:SpiralChartArray[31] = Add-Number
    $Global:SpiralChartArray[32] = Add-Number
    $Global:SpiralChartArray[33] = Add-Number
    $Global:SpiralChartArray[26] = Add-Number
    $Global:SpiralChartArray[19] = Add-Number
    $Global:SpiralChartArray[12] = Add-Number
    $Global:SpiralChartArray[11] = Add-Number
    $Global:SpiralChartArray[10] = Add-Number
    $Global:SpiralChartArray[9] = Add-Number
    $Global:SpiralChartArray[8] = Add-Number
    $Global:SpiralChartArray[15] = Add-Number
    $Global:SpiralChartArray[22] = Add-Number
    $Global:SpiralChartArray[29] = Add-Number
    $Global:SpiralChartArray[36] = Add-Number
    $Global:SpiralChartArray[37] = Add-Number
    $Global:SpiralChartArray[38] = Add-Number
    $Global:SpiralChartArray[39] = Add-Number
    $Global:SpiralChartArray[40] = Add-Number
    $Global:SpiralChartArray[41] = Add-Number
    $Global:SpiralChartArray[34] = Add-Number
    $Global:SpiralChartArray[27] = Add-Number
    $Global:SpiralChartArray[20] = Add-Number
    $Global:SpiralChartArray[13] = Add-Number
    $Global:SpiralChartArray[6] = Add-Number
    $Global:SpiralChartArray[5] = Add-Number
    $Global:SpiralChartArray[4] = Add-Number
    $Global:SpiralChartArray[3] = Add-Number
    $Global:SpiralChartArray[2] = Add-Number
    $Global:SpiralChartArray[1] = Add-Number
    $Global:SpiralChartArray[0] = Add-Number
    $Global:SpiralChartArray[7] = Add-Number
    $Global:SpiralChartArray[14] = Add-Number
    $Global:SpiralChartArray[21] = Add-Number
    $Global:SpiralChartArray[28] = Add-Number
    $Global:SpiralChartArray[35] = Add-Number
    $Global:SpiralChartArray[42] = Add-Number
    $Global:SpiralChartArray[43] = Add-Number
    $Global:SpiralChartArray[44] = Add-Number
    $Global:SpiralChartArray[45] = Add-Number
    $Global:SpiralChartArray[46] = Add-Number
    $Global:SpiralChartArray[47] = Add-Number
    $Global:SpiralChartArray[48] = Add-Number
}

function Reference-RegressionDrawNumberAndYear {
param(
    [int]$DrawNumber = 0,
    [int]$Year = 0,
    [int]$RegressionMultiplier = 0
)

[int]$RegressionDrawNumber = 0
[int]$RegressionYear = 0
[array]$TempArray = $NumberOfDraw.GetEnumerator() | sort Name
$ResultArray = @()

    if ($DrawNumber - $RegressionMultiplier -eq 0) {
        foreach ($i in $TempArray) {
            if ($i.Name -eq $Year - 1) {
                $RegressionYear = $Year - 1
                $RegressionDrawNumber = $i.Value
            }
        }
    } elseif ($DrawNumber - $RegressionMultiplier -eq -1) {
        foreach ($i in $TempArray) {
            if ($i.Name -eq $Year - 1) {
                $RegressionYear = $Year - 1
                $RegressionDrawNumber = $i.Value - 1
            }
        }
    } elseif ($DrawNumber - $RegressionMultiplier -lt -1) {
        foreach ($i in $TempArray) {
            if ($i.Name -eq $Year - 1) {
                $RegressionYear = $Year - 1
                $RegressionDrawNumber = $i.Value - ($RegressionMultiplier - $DrawNumber)
            }
        }
    } else {
        $RegressionYear = $Year
        $RegressionDrawNumber = $DrawNumber - $RegressionMultiplier
    }

    [string]$RegressionShortDrawNumber = $RegressionDrawNumber
    $RegressionShortDrawNumber = $RegressionShortDrawNumber.Trim()
    if ($RegressionShortDrawNumber.Length -eq 2) {
        $RegressionShortDrawNumber = "0" + $RegressionShortDrawNumber
    } elseif ($RegressionShortDrawNumber.Length -eq 1) {
        $RegressionShortDrawNumber = "00" + $RegressionShortDrawNumber
    }

    $obj = New-Object -TypeName PSobject
    Add-Member -InputObject $obj -MemberType NoteProperty -Name "Year" -Value $RegressionYear
    Add-Member -InputObject $obj -MemberType NoteProperty -Name "ShortDrawNumber" -Value $RegressionShortDrawNumber
    $ResultArray += $obj 

    return $ResultArray
}

function Reference-SumOfDrawNumber {
param(
    [array]$TargetDraw = @()
)
    #Local Variable
    $FullDrawResult = $Global:DrawResultMultipleYears | sort DrawNumber
    [int]$SumOfDrawNumber = 0

    foreach ($i in $FullDrawResult) {
        if ($i.DrawYear -eq $TargetDraw.Year -and $i.ShortDrawNumber -eq $TargetDraw.ShortDrawNumber) {
            $SumOfDrawNumber += [int]$i.Number1
            $SumOfDrawNumber += [int]$i.Number2
            $SumOfDrawNumber += [int]$i.Number3
            $SumOfDrawNumber += [int]$i.Number4
            $SumOfDrawNumber += [int]$i.Number5
            $SumOfDrawNumber += [int]$i.Number6
        }
    }
    return $SumOfDrawNumber
}

function Generate-RegressionNumber {
param(
    $Number = 0
)
    $Number = $Number - (49 - $Global:MiddleNumber) - 1
    
    while ($Number -gt 49) {
        $Number -= 49
    }

    return $Number
}

function Display-SpiralChartArray-Normal {
    [string]$OutputString = "`n"
    for ($i = 0; $i -lt $Global:SpiralChartArray.Count; $i++) {
        if ($i -ne 0 -and $i % 7 -eq 0) {
            $OutputString += "`n"
        }

        [string]$num = ($Global:SpiralChartArray[$i]).ToString()
        
        if ($num.Length -eq 1) {
            $num = "0" + $num
        }
        $OutputString += $num
        $OutputString += " "
    }
    Write-Host $OutputString
    Write-Host "`n"
}

function Display-SpiralChart-Color {
    for ($i = 0; $i -lt $Global:SpiralChartArray.Count; $i++) {
        if ($i -ne 0 -and $i % 7 -eq 0) {
            Write-Host " "
        }

        [int]$num = $Global:SpiralChartArray[$i]

        if ($Red_List -contains $num) {
            $Color = "Red"
        } elseif ($Blue_List -contains $num) {
            $Color = "Blue"
        } else {
            $Color = "DarkGreen"
        }

        [string]$num = ($Global:SpiralChartArray[$i]).ToString()
        
        if ($num.Length -eq 1) {
            $num = "0" + $num + " "
        } else {
            $num += " "
        }

        Write-Host $num -BackgroundColor $Color -NoNewline
    }
    Write-Host "`n"
}

function Display-SpiralChart-Color-Regression {
    Write-Host "`n"
    for ($i = 0; $i -lt $Global:SpiralChartArray.Count; $i++) {
        if ($i -ne 0 -and $i % 7 -eq 0) {
            Write-Host " "
        }

        [int]$num = $Global:SpiralChartArray[$i]

        if ($Global:RegressionNumberArray -contains $num) {
            $BackgroundColor = "Yellow"
            $ForegroundColor = "Black"
        } elseif ($i - 8 -ge 0 -and $i -gt 6 -and $i % 7 -ne 0 -and $Global:RegressionNumberArray -contains $Global:SpiralChartArray[$i - 8]) {
            $BackgroundColor = "White"
            $ForegroundColor = "Black"
        } elseif ($i - 7 -ge 0 -and $i -gt 6 -and $Global:RegressionNumberArray -contains $Global:SpiralChartArray[$i - 7]) {
            $BackgroundColor = "White"
            $ForegroundColor = "Black"
        } elseif ($i - 6 -gt 0 -and $i -gt 6 -and $i % 7 -ne 6 -and $Global:RegressionNumberArray -contains $Global:SpiralChartArray[$i - 6]) {
            $BackgroundColor = "White"
            $ForegroundColor = "Black"
        } elseif ($i -ne 0 -and $i % 7 -ne 0 -and $Global:RegressionNumberArray -contains $Global:SpiralChartArray[$i - 1]) {
            $BackgroundColor = "White"
            $ForegroundColor = "Black"
        } elseif ($i -ne 6 -and $i % 7 -ne 6 -and ($Global:RegressionNumberArray -contains $Global:SpiralChartArray[$i + 1])) {
            $BackgroundColor = "White"
            $ForegroundColor = "Black"
        } elseif ($i + 6 -lt 50 -and $i -lt 42 -and $i % 7 -ne 0 -and $Global:RegressionNumberArray -contains $Global:SpiralChartArray[$i + 6]) {
            $BackgroundColor = "White"
            $ForegroundColor = "Black"
        } elseif ($i + 7 -lt 50 -and $i -lt 42 -and $Global:RegressionNumberArray -contains $Global:SpiralChartArray[$i + 7]) {
            $BackgroundColor = "White"
            $ForegroundColor = "Black"
        } elseif ($i + 8 -lt 50 -and $i -lt 42 -and $i % 7 -ne 6 -and $Global:RegressionNumberArray -contains $Global:SpiralChartArray[$i + 8]) {
            $BackgroundColor = "White"
            $ForegroundColor = "Black"
        } elseif ($Red_List -contains $num) {
            $BackgroundColor = "Red"
            $ForegroundColor = "White"
        } elseif ($Blue_List -contains $num) {
            $BackgroundColor = "Blue"
            $ForegroundColor = "White"
        } else {
            $BackgroundColor = "DarkGreen"
            $ForegroundColor = "White"
        }

        [string]$num = ($Global:SpiralChartArray[$i]).ToString()
        
        if ($num.Length -eq 1) {
            $num = "0" + $num + " "
        } else {
            $num += " "
        }

        Write-Host $num -BackgroundColor $BackgroundColor -ForegroundColor $ForegroundColor -NoNewline
    }
    Write-Host "`n"
}

function Generate-RegressionNumberOfDraw-Method-1-7-14-21-49-100 {
param(
    [string]$TargetDrawYear,
    [string]$TargetShortDrawNumber,
    [string]$PastDrawYear,
    [string]$PastShortDrawNumber
) 
    $Global:RegressionNumberArray = @()

    [array]$ResultArray = Reference-RegressionDrawNumberAndYear -DrawNumber $TargetShortDrawNumber -Year $TargetDrawYear -RegressionMultiplier 1
    [int]$SumOfDrawNumber = Reference-SumOfDrawNumber -TargetDraw $ResultArray
    $Global:RegressionNumberArray += Generate-RegressionNumber -Number $SumOfDrawNumber 
    
    [array]$ResultArray = Reference-RegressionDrawNumberAndYear -DrawNumber $TargetShortDrawNumber -Year $TargetDrawYear -RegressionMultiplier 7
    [int]$SumOfDrawNumber = Reference-SumOfDrawNumber -TargetDraw $ResultArray
    $Global:RegressionNumberArray += Generate-RegressionNumber -Number $SumOfDrawNumber 
    
    [array]$ResultArray = Reference-RegressionDrawNumberAndYear -DrawNumber $TargetShortDrawNumber -Year $TargetDrawYear -RegressionMultiplier 14
    [int]$SumOfDrawNumber = Reference-SumOfDrawNumber -TargetDraw $ResultArray
    $Global:RegressionNumberArray += Generate-RegressionNumber -Number $SumOfDrawNumber 
    
    [array]$ResultArray = Reference-RegressionDrawNumberAndYear -DrawNumber $TargetShortDrawNumber -Year $TargetDrawYear -RegressionMultiplier 21
    [int]$SumOfDrawNumber = Reference-SumOfDrawNumber -TargetDraw $ResultArray
    $Global:RegressionNumberArray += Generate-RegressionNumber -Number $SumOfDrawNumber 
    
    [array]$ResultArray = Reference-RegressionDrawNumberAndYear -DrawNumber $TargetShortDrawNumber -Year $TargetDrawYear -RegressionMultiplier 49
    [int]$SumOfDrawNumber = Reference-SumOfDrawNumber -TargetDraw $ResultArray
    $Global:RegressionNumberArray += Generate-RegressionNumber -Number $SumOfDrawNumber 
    
    [array]$ResultArray = Reference-RegressionDrawNumberAndYear -DrawNumber $TargetShortDrawNumber -Year $TargetDrawYear -RegressionMultiplier 100
    [int]$SumOfDrawNumber = Reference-SumOfDrawNumber -TargetDraw $ResultArray
    $Global:RegressionNumberArray += Generate-RegressionNumber -Number $SumOfDrawNumber 
    
    $Global:RegressionNumberArray = $Global:RegressionNumberArray | sort $_
    
    Write-Host "`nResult of Past Draw $PastShortDrawNumber in $PastDrawYear`n$Global:PastDrawResult"
    Write-Host "`nResult of Past Draw $PastShortDrawNumber in $PastDrawYear (Sequential Order)`n$Global:NumberInSequential"
    Write-Host ("`n6 Regression Number of Target Draw $TargetShortDrawNumber in $TargetDrawYear`n" + $Global:RegressionNumberArray)
    Write-Host "`nMiddle Number of Target Draw is $Global:MiddleNumber"
}

function Download-DrawResult {
param(
    [string]$DrawYear = ""
) 
    #Local Variable 
    $YearList = @($DrawYear)
    $TempDrawResult = @()

    foreach ($Year in $YearList) {
        #Local Variable
        $uri = "https://www.cpzhan.com/liu-he-cai/all-results/?year=" + $Year + "&sort=cmp"
        $method = "GET"
        $firstDataRow = 1 #Skip first 1 information row

        #Retrieve Data
        $returnData = Invoke-WebRequest -Uri $uri -Method $method
        Start-Sleep -Milliseconds 100
        $htmlTable = $returnData.ParsedHtml.IHTMLDocument3_getElementsByTagName("table") 
        $totalRows = $($htmlTable[0].rows).count

        #Add to Temp Array
        for ($idx = $firstDataRow; $idx -lt $totalRows; $idx++) {
            $currentRow = $htmlTable[0].rows[$idx]
            $currentCells = $currentRow.cells    

            if ($currentCells[0].tagName -eq 'td') {
                $text = $currentCells | % {$_.InnerText -replace ' ',''}
                    
                $option = [System.StringSplitOptions]::RemoveEmptyEntries
                [array]$currentCellsArray = $text.Split("`n", $option)

                #Draw Year
                $DrawYear = $currentCellsArray[0]
                $DrawYear = $DrawYear.Trim()

                #Short Draw Number
                [string]$ShortDrawNumber = $currentCellsArray[1]
                $ShortDrawNumber = $ShortDrawNumber.Trim()
                if ($ShortDrawNumber.Length -eq 2) {
                    $ShortDrawNumber = "0" + $ShortDrawNumber
                } elseif ($ShortDrawNumber.Length -eq 1) {
                    $ShortDrawNumber = "00" + $ShortDrawNumber
                }

                #Full Draw Number
                $DrawNumber = $DrawYear + $ShortDrawNumber

                #Process
                $obj = New-Object -TypeName PSobject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "DrawNumber" -Value $DrawNumber
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "DrawYear" -Value $DrawYear 
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "ShortDrawNumber" -Value $ShortDrawNumber 
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "DrawDate" -Value $currentCellsArray[2]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Number1" -Value $currentCellsArray[3]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Number2" -Value $currentCellsArray[4]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Number3" -Value $currentCellsArray[5]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Number4" -Value $currentCellsArray[6]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Number5" -Value $currentCellsArray[7]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Number6" -Value $currentCellsArray[8]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "SpecialNumber" -Value $currentCellsArray[9]
                $TempDrawResult += $obj 
            }
        }
    }
    $Global:DrawResultMultipleYears += $TempDrawResult
}

function Download-SequentialDrawResult {
param(
    [string]$DrawYear = ""
) 
    #Local Variable 
    $YearList = @($DrawYear)
    $TempDrawResult = @()

    foreach ($Year in $YearList) {
        #Local Variable
        $uri = "https://www.cpzhan.com/liu-he-cai/all-results/?year=" + $Year + "&sort=seq"
        $method = "GET"
        $firstDataRow = 1 #Skip first 1 information row

        #Retrieve Data
        $returnData = Invoke-WebRequest -Uri $uri -Method $method
        Start-Sleep -Milliseconds 100
        $htmlTable = $returnData.ParsedHtml.IHTMLDocument3_getElementsByTagName("table") 
        $totalRows = $($htmlTable[0].rows).count

        #Add to Temp Array
        for ($idx = $firstDataRow; $idx -lt $totalRows; $idx++) {
            $currentRow = $htmlTable[0].rows[$idx]
            $currentCells = $currentRow.cells    

            if ($currentCells[0].tagName -eq 'td') {
                $text = $currentCells | % {$_.InnerText -replace ' ',''}
                    
                $option = [System.StringSplitOptions]::RemoveEmptyEntries
                [array]$currentCellsArray = $text.Split("`n", $option)

                #Draw Year
                $DrawYear = $currentCellsArray[0]
                $DrawYear = $DrawYear.Trim()

                #Short Draw Number
                [string]$ShortDrawNumber = $currentCellsArray[1]
                $ShortDrawNumber = $ShortDrawNumber.Trim()
                if ($ShortDrawNumber.Length -eq 2) {
                    $ShortDrawNumber = "0" + $ShortDrawNumber
                } elseif ($ShortDrawNumber.Length -eq 1) {
                    $ShortDrawNumber = "00" + $ShortDrawNumber
                }

                #Full Draw Number
                $DrawNumber = $DrawYear + $ShortDrawNumber

                #Process
                $obj = New-Object -TypeName PSobject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "DrawNumber" -Value $DrawNumber
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "DrawYear" -Value $DrawYear 
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "ShortDrawNumber" -Value $ShortDrawNumber 
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "DrawDate" -Value $currentCellsArray[2]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "S_Number1" -Value $currentCellsArray[3]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "S_Number2" -Value $currentCellsArray[4]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "S_Number3" -Value $currentCellsArray[5]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "S_Number4" -Value $currentCellsArray[6]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "S_Number5" -Value $currentCellsArray[7]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "S_Number6" -Value $currentCellsArray[8]
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "SpecialNumber" -Value $currentCellsArray[9]
                $TempDrawResult += $obj 
            }
        }
    }
    $Global:SequentialDrawResult += $TempDrawResult
}

function Reference-SequentialDrawResult {
param(
    [string]$DrawYear,
    [string]$ShortDrawNumber,
    [array]$Position
)  

    $TempArray = @()

    for ($i = 0; $i -lt $Global:SequentialDrawResult.Count; $i++) {
        if ($Global:SequentialDrawResult[$i].DrawYear -eq $DrawYear -and $Global:SequentialDrawResult[$i].ShortDrawNumber -eq $ShortDrawNumber) {
            $SelectedIndex = $i
        }
    }

    foreach ($i in $Position) {
        [int]$num = $i

        switch ($num) {
            1 { $TempArray += $Global:SequentialDrawResult[$SelectedIndex].S_Number1 }
            2 { $TempArray += $Global:SequentialDrawResult[$SelectedIndex].S_Number2 }
            3 { $TempArray += $Global:SequentialDrawResult[$SelectedIndex].S_Number3 }
            4 { $TempArray += $Global:SequentialDrawResult[$SelectedIndex].S_Number4 }
            5 { $TempArray += $Global:SequentialDrawResult[$SelectedIndex].S_Number5 }
            6 { $TempArray += $Global:SequentialDrawResult[$SelectedIndex].S_Number6 }
            7 { $TempArray += $Global:SequentialDrawResult[$SelectedIndex].SpecialNumber }
        }
    }
    return $TempArray
}

function Display-CompletedDraw-Result {
param(
    [string]$DrawYear,
    [string]$ShortDrawNumber
) 
    #Find the Draw Result List if the Draw is completed
    $FullDrawResult = $Global:DrawResultMultipleYears | sort DrawNumber -Descending 
    $DrawResultIndex = -1

    foreach ($i in $FullDrawResult) {
        if ($i.DrawYear -eq $DrawYear -and $i.ShortDrawNumber -eq $ShortDrawNumber) {
            $DrawResultIndex = $i
            $Global:TargetDrawResult = @([int]$i.Number1,[int]$i.Number2,[int]$i.Number3,[int]$i.Number4,[int]$i.Number5,[int]$i.Number6,[int]$i.SpecialNumber)
        }
    }

    if ($DrawResultIndex -ne -1) {
        Write-Host "`nResult of Target Draw $ShortDrawNumber in $DrawYear`n$Global:TargetDrawResult`n"
    }
}

#Analysis and Validation
Write-Host "`nA: Generate Spiral Chart of specific Draw"
Write-Host "`nB: Get the list of specific Middle Number"
Write-Host "`nC: Generate a series of Spiral Chart with specific Middle Number"
$OperationOption = Read-Host -Prompt "`nSelect the operation: "
Write-Host "---------------------------------------------------------------------------------------------------------`n"

#Import
$Global:DrawResultMultipleYears = Import-Csv .\DrawResult-2012-To-2019.csv -Delimiter ","
#$Global:DrawResultMultipleYears += Import-Csv .\DrawResult-2020.csv -Delimiter ","
Download-DrawResult -DrawYear "2020"

$Global:SequentialDrawResult = Import-Csv .\SequentialDrawResult-2012-To-2019.csv -Delimiter ","
#$Global:SequentialDrawResult += Import-Csv .\SequentialDrawResult-2020.csv -Delimiter ","
Download-SequentialDrawResult -DrawYear "2020"

if ($OperationOption -eq "A") {

    #Customization
    [string]$Global:PastDrawYear = "2020"
    [string]$Global:PastShortDrawNumber = "030"
    [string]$Global:TargetDrawYear = "2020"
    [string]$Global:TargetShortDrawNumber = "031"

    #Process
    Reference-MiddleNumber -DrawYear $Global:PastDrawYear -ShortDrawNumber $Global:PastShortDrawNumber
    $Global:NumberInSequential = Reference-SequentialDrawResult -DrawYear $Global:PastDrawYear -ShortDrawNumber $Global:PastShortDrawNumber -Position @(1,2,3,4,5,6,7)
    Generate-SpiralChartArray
    #Display-SpiralChartArray-Normal
    #Display-SpiralChart-Color
    Generate-RegressionNumberOfDraw-Method-1-7-14-21-49-100 -TargetDrawYear $Global:TargetDrawYear -TargetShortDrawNumber $Global:TargetShortDrawNumber -PastDrawYear $Global:PastDrawYear -PastShortDrawNumber $Global:PastShortDrawNumber
    Display-SpiralChart-Color-Regression
    Display-CompletedDraw-Result -DrawYear $Global:TargetDrawYear -ShortDrawNumber $Global:TargetShortDrawNumber
} elseif ($OperationOption -eq "B") {

    #Local Variable
    [int]$EnteredMiddleNumber = Read-Host -Prompt "`nEnter the Middle Number: "

    foreach ($i in $Global:DrawResultMultipleYears) {
        [int]$FirstNumber = $i.Number1
        [int]$LastNumber = $i.Number6
        [double]$SumOfFirstAndLastNumber = $FirstNumber + $LastNumber
    
        if ([math]::ceiling($SumOfFirstAndLastNumber / 2) -eq $EnteredMiddleNumber) {
            Write-Host ("Past Draw: " + $i.DrawNumber)
        }
    }
} elseif ($OperationOption -eq "C") {
    #Local Variable
    [string]$EnteredYear = Read-Host -Prompt "`nEnter the Start Past-Year "
    [string]$EnteredShortDrawNumber = Read-Host -Prompt "`nEnter the Start Past-Short-Draw-Number "
    [int]$NumberOfPastDraw = Read-Host -Prompt "`nEnter the number of previous draw "
    [int]$FirstSequentialNumber = Read-Host -Prompt "`nEnter the number of First Sequential Number (Enter -1 to skip)"
   
    #Find the Draw Result List if the Draw is completed
    $FullDrawResult = $Global:DrawResultMultipleYears | sort DrawNumber -Descending

    #Find the Start Index from Full List
    $StartIndex = -1
    $EnteredMiddleNumber = -1
    for ($i = 0; $i -lt $FullDrawResult.Count; $i++) {

        if ($FullDrawResult[$i].DrawYear -eq $EnteredYear -and $FullDrawResult[$i].ShortDrawNumber -eq $EnteredShortDrawNumber) {
            $StartIndex = $i

            #Get the Middle Number based on Past Draw
            [int]$FirstNumber = $FullDrawResult[$i].Number1
            [int]$LastNumber = $FullDrawResult[$i].Number6
            [double]$SumOfFirstAndLastNumber = $FirstNumber + $LastNumber
            $EnteredMiddleNumber = [math]::ceiling($SumOfFirstAndLastNumber / 2)
            break;
        }
    }

    #Display the Chart of selected number of Draw
    for ($i = $StartIndex ; $i -lt ($FullDrawResult.Count - $StartIndex); $i++) {

        #Get MiddleNumber of each Draw
        [int]$FirstNumber = $FullDrawResult[$i].Number1
        [int]$LastNumber = $FullDrawResult[$i].Number6
        [double]$SumOfFirstAndLastNumber = $FirstNumber + $LastNumber

        #Verify
        if ([math]::ceiling($SumOfFirstAndLastNumber / 2) -eq $EnteredMiddleNumber) {
            $NumberOfPastDraw -= 1

            #Reset Global Variable
            $Global:RegressionDrawNumberAndYearArray = @()
            $Global:SpiralChartArray = New-Object Object[] 49
            $Global:RegressionNumberArray = @()
            [int]$Global:AddNumberCount = 0

            #Prepare
            Reference-MiddleNumber -DrawYear $FullDrawResult[$i].DrawYear -ShortDrawNumber $FullDrawResult[$i].ShortDrawNumber
            $Global:NumberInSequential = Reference-SequentialDrawResult -DrawYear $FullDrawResult[$i].DrawYear -ShortDrawNumber $FullDrawResult[$i].ShortDrawNumber -Position @(1,2,3,4,5,6,7)

            if ($Global:NumberInSequential[0] -eq $FirstSequentialNumber -or $FirstSequentialNumber -eq -1) {
                #Display
                Write-Host ("`n------------------------------") -ForegroundColor Yellow
                Write-Host ("Past Draw: " + $FullDrawResult[$i].DrawNumber) -ForegroundColor Yellow
                Write-Host ("Target Draw: " + $FullDrawResult[$i-1].DrawNumber) -ForegroundColor Yellow

                #Process
                Generate-SpiralChartArray
                #Display-SpiralChartArray-Normal
                #Display-SpiralChart-Color
                Generate-RegressionNumberOfDraw-Method-1-7-14-21-49-100 -TargetDrawYear $FullDrawResult[$i-1].DrawYear -TargetShortDrawNumber $FullDrawResult[$i-1].ShortDrawNumber -PastDrawYear $FullDrawResult[$i].DrawYear -PastShortDrawNumber $FullDrawResult[$i].ShortDrawNumber
                Display-SpiralChart-Color-Regression
                Display-CompletedDraw-Result -DrawYear $FullDrawResult[$i-1].DrawYear -ShortDrawNumber $FullDrawResult[$i-1].ShortDrawNumber

                Write-Host ("`n")
            }
        }

        if ($NumberOfPastDraw -eq 0) {
            break;
        }
    }
}