<#
.SYNOPSIS
OU-Calculator.ps1 - UNOFFICIAL Script for calculating the final degree score at the Open University.

.DESCRIPTION 
This PowerShell script will calculate the final degree score based on the modules studied so far

Use $(Get-Command .\OU-Calculator.ps1 -Syntax) to get the different options

.PARAMETER Modules
[Mandatory] Specifies the source array of modules studied so far. This paramenter is an Array.

.EXAMPLE
##### GRADE LEGEND #####
# 1. Distinction  = 1  #
# 2. Grade 2 Pass = 2  #
# 3. Grade 3 Pass = 3  #
# 4. Grade 4 Pass = 4  #
########################

Example
ModuleName Level Grade Credits
---------- ----- ----- -------
TM470          3     2      30
TM354          3     2      30
TM355          3     2      30
TM352          3     2      30
TT284          2     2      30
TM255          2     1      30
TM254          2     3      30
M250           2     2      30

$modules = @( 
    [PSCustomObject]@{ModuleName = "TM470"; Level = 3; Grade = 2; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM354"; Level = 3; Grade = 2; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM355"; Level = 3; Grade = 2; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM352"; Level = 3; Grade = 2; Credits = 30; },

    [PSCustomObject]@{ModuleName = "TT284"; Level = 2; Grade = 2; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM255"; Level = 2; Grade = 1; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM254"; Level = 2; Grade = 3; Credits = 30; },
    [PSCustomObject]@{ModuleName = "M250"; Level = 2; Grade = 2; Credits = 30; }
)

OU-Calculator -Modules $modules

PS C:\Users\GD-HCK>.\OU-Calculator

.NOTES
Written by: GD-HCK

Find me on:

* Github:	https://github.com/GD-HCK

Official Documentations
* Website:	
    1. https://help.open.ac.uk/documents/policies/working-out-your-class-of-honours/files/50/honours-class-working-out.pdf

Change Log
V1.00, 02/07/2022 - Initial version,
V1.01, 21/08/2022 - improved calculation algorithm,
#>
function OU-Calculator {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [array]
        $Modules
    )

    
    $Thresholds = @(
        [PSCustomObject]@{
            Name    = "Full";
            Classes = @(
                [PSCustomObject]@{Value = "First class" ; Range = @(0..630) ; Colour = "Green" },
                [PSCustomObject]@{Value = "Upper second class" ; Range = @(631..900) ; Colour = "DarkYellow" },
                [PSCustomObject]@{Value = "Lower second class" ; Range = @(901..1170) ; Colour = "DarkCyan" },
                [PSCustomObject]@{Value = "Third Class" ; Range = @(1171..1440) ; Colour = "Red" }
            )
        },
        [PSCustomObject]@{
            Name    = "Quality";
            Classes = @(
                [PSCustomObject]@{Value = "First class" ; Range = @(0..60) ; Colour = "Green" },
                [PSCustomObject]@{Value = "Upper second class" ; Range = @(61..120) ; Colour = "DarkYellow" },
                [PSCustomObject]@{Value = "Lower second class" ; Range = @(121..180) ; Colour = "DarkCyan" },
                [PSCustomObject]@{Value = "Third class" ; Range = @(181..240) ; Colour = "Red" }
            )
        }
    )
    

    $errorFound = $false
    $levels = @(
        [PSCustomObject]@{Name = "Level3"; Modules = $($Modules | ? { $_.Level -eq 3 } | Sort-Object -Property Grade | Sort-Object -Property Credits -Descending) }
        [PSCustomObject]@{Name = "Level2"; Modules = $($Modules | ? { $_.Level -eq 2 } | Sort-Object -Property Grade | Sort-Object -Property Credits -Descending) }
    )

    #$BestLevel3Module = $Modules | ? { $_.Level -eq 3 } | Sort-Object -Property Grade | Sort-Object -Property Credits -Descending | Select-Object -First 1
    #$BestLevel2Module = $Modules | ? { $_.Level -eq 2 } | Sort-Object -Property Grade | Sort-Object -Property Credits -Descending | Select-Object -First 1


    $temp = @()
    

    foreach ($level in $levels) {
        $level.modules | % {
            if ($_.Credits -gt 60 -or $_.Credits -le 0) { $errorFound = $true; $temp += [PSCustomObject]@{Module = $_.ModuleName ; Value = "Credits: $($_.Credits)"; Error = "Modules cannot hold 0 or over 60 credits" } }
            if ($_.Grade -gt 4 -or $_.Grade -lt 1) { $errorFound = $true; $temp += [PSCustomObject]@{Module = $_.ModuleName ; Value = "Grade: $($_.Grade)"; Error = "Grades allowed are 1 to 4" } }
            if ($_.Level -gt 3 -or $_.Grade -lt 1) { $errorFound = $true; $temp += [PSCustomObject]@{Module = $_.ModuleName ; Value = "Level: $($_.Level)"; Error = "Levels allowed are 1 to 3" } }
        }
    }
    
    if ($errorFound) {
        Write-Host "Correct the below module(s)" -ForegroundColor Yellow
        $temp | % { Write-Host "Module: $($_.Module), $($_.Value), Error: $($_.Error)" -ForegroundColor Red }
        Read-Host
        break
    }

    $TotalScore = 0
    $QualityScore = 0
    foreach ($level in $levels) {

        $SortedModulesScores = @()
        [System.Collections.ArrayList]$mods = $Level.Modules

        while ($mods.Count -gt 0) {
            #Write-Host '$mods.count:' $mods.count -ForegroundColor Yellow
            for ($i = 0; $i -lt $mods.Count; $i++) {
                $count = 0
                #Write-Host '$mods[$i]:' $mods[$i].ModuleName  -ForegroundColor Red
                while ($count -lt $mods.Count) {    
                    $score = 0
                    $QualityScoreTemp = 0
                    $credits = 0
                    $Modules = @()
                    #Write-Host '$mods[$count]:' $mods[$count].ModuleName  -ForegroundColor Magenta
                    $credits = ($($mods[$i].Credits) + $($mods[$count].Credits))
                    if ($level.Name -ieq "level3") {
                        $score = ($($mods[$i].Credits * $mods[$i].Grade * 2) + $($mods[$count].Credits * $mods[$count].Grade * 2))
                        $QualityScoreTemp = ($($mods[$i].Credits * $mods[$i].Grade) + $($mods[$count].Credits * $mods[$count].Grade))
                    }
                    else {
                        $score = ($($mods[$i].Credits * $mods[$i].Grade) + $($mods[$count].Credits * $mods[$count].Grade))
                    }
                    $Modules = @($mods[$i],$mods[$count])

                    $obj = [PSCustomObject]@{
                        Credits      = $credits;
                        QualityScore = $QualityScoreTemp;
                        Score        = $score;
                        Modules      = $Modules;
                    }

                    if ($SortedModulesScores -notcontains $obj) {
                        #Write-Host 'Adding $obj:' $obj.Modules  -ForegroundColor Magenta
                        $SortedModulesScores += $obj
                    }

                    $obj = $null
                    $score = 0
                    $QualityScoreTemp = 0
                    $credits = 0
                    $Modules = @()

                    for ($y = 0; $y -le $count; $y++) {
                        #Write-Host '$mods[$y]:' $mods[$y].ModuleName  -ForegroundColor Green
                        $credits += $($mods[$y].Credits)
                        if ($level.Name -ieq "level3") {
                            $score += $($mods[$y].Credits * $mods[$y].Grade * 2)
                            $QualityScoreTemp += $($mods[$y].Credits * $mods[$y].Grade)
                        }
                        else {
                            $score += $($mods[$y].Credits * $mods[$y].Grade)
                        }
                        $Modules += $mods[$y]
                    }

                    $obj = [PSCustomObject]@{
                        Credits      = $credits;
                        QualityScore = $QualityScoreTemp;
                        Score        = $score;
                        Modules      = $Modules;
                    }
                    

                    if ($SortedModulesScores -notcontains $obj) {
                        #Write-Host 'Adding $obj:' $obj.Modules  -ForegroundColor Green
                        $SortedModulesScores += $obj
                    }
                    $count++
                }
            }
            $mods.Remove($mods[0])
        }
        $LevelFiltered = $SortedModulesScores | ? { $_.Credits -eq 120 } | Sort-Object -Property Score | Select-Object -First 1
        $TotalScore += $LevelFiltered.Score

        if ($level.Name -ieq "level3") {
            $Level3Filtered = $LevelFiltered
            $LevelQualityFiltered = $SortedModulesScores | ? { $_.Credits -eq 60 } | Sort-Object -Property QualityScore | Select-Object -First 1
            $QualityScore = $LevelQualityFiltered.QualityScore
            $BestLevel3Module = $LevelFiltered.Modules | ? { $_.Level -eq 3 } | Sort-Object -Property Grade | Sort-Object -Property Credits -Descending | Select-Object -First 1
        }
        else {
            $BestLevel2Module = $LevelFiltered.Modules | ? { $_.Level -eq 2 } | Sort-Object -Property Grade | Sort-Object -Property Credits -Descending | Select-Object -First 1
        }
    }


    foreach ($threshold in $Thresholds) {
        if ($threshold.Name -eq "Full") {
            $threshold.Classes | % {
                if ($_.Range -contains $TotalScore) {
                    $TotalValue = $_.Value
                    $resultColour = $_.Colour
                }
            }
        }
        else {
            $threshold.Classes | % {
                if ($_.Range -contains $QualityScore) {
                    $QualityValue = $_.Value
                    $resultColour = $_.Colour
                }
            }
        }
    }

    Clear-Host
    if ($TotalValue -eq $QualityValue) {
        Write-Host "+===================================================================+" -ForegroundColor Magenta
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host 'Assessment Score                   =  ' -ForegroundColor Yellow -NoNewline; Write-Host "$TotalScore points" -ForegroundColor Cyan
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host 'Best Lv.3 Module                   =  ' -ForegroundColor Yellow -NoNewline; Write-Host "$($BestLevel3Module.ModuleName), Score: $($BestLevel3Module.Grade)" -ForegroundColor Cyan
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host 'Best Lv.2 Module                   =  ' -ForegroundColor Yellow -NoNewline; Write-Host "$($BestLevel2Module.ModuleName), Score: $($BestLevel2Module.Grade)" -ForegroundColor Cyan
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host 'Lv.3 Assessment Modules            =  ' -ForegroundColor Yellow -NoNewline; Write-Host $($Level3Filtered.Modules.ModuleName -join ", ") -ForegroundColor Cyan
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host 'Lv.2 Assessment Modules            =  ' -ForegroundColor Yellow -NoNewline; Write-Host $($LevelFiltered.Modules.ModuleName -join ", ") -ForegroundColor Cyan
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host '                                      ' -ForegroundColor Yellow
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host 'Quality Assessment Score           =  ' -ForegroundColor Yellow -NoNewline; Write-Host "$QualityScore points" -ForegroundColor Cyan
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host 'Lv.3 Quality Assessment Modules    =  ' -ForegroundColor Yellow -NoNewline; Write-Host $($LevelQualityFiltered.Modules.ModuleName -join ", ") -ForegroundColor Cyan
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host '                                      ' -ForegroundColor Yellow
        Write-Host '  ' -ForegroundColor Magenta -NoNewline; Write-Host 'Your honour class is               ' -ForegroundColor DarkMagenta -NoNewline; Write-Host '=  ' -ForegroundColor Yellow -NoNewline; Write-Host $QualityValue -ForegroundColor $resultColour
        Write-Host "+===================================================================+" -ForegroundColor Magenta

    }
    else {
        Write-Host "Something went wrong. Scores test should match" -ForegroundColor Yellow
        Write-Host "Total test: $TotalValue" -ForegroundColor Yellow
        Write-Host "Total Score: $TotalScore" -ForegroundColor Yellow
        Write-Host "Quality test: $QualityValue" -ForegroundColor Yellow
        Write-Host "Quality Score: $QualityScore" -ForegroundColor Yellow
    }
}

##### GRADE LEGEND #####
# 1. Distinction  = 1  #
# 2. Grade 2 Pass = 2  #
# 3. Grade 3 Pass = 3  #
# 4. Grade 4 Pass = 4  #
########################


$modules = @( 
    [PSCustomObject]@{ModuleName = "TM470"; Level = 3; Grade = 2; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM354"; Level = 3; Grade = 2; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM355"; Level = 3; Grade = 2; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM352"; Level = 3; Grade = 2; Credits = 30; },

    [PSCustomObject]@{ModuleName = "TT284"; Level = 2; Grade = 2; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM255"; Level = 2; Grade = 1; Credits = 30; },
    [PSCustomObject]@{ModuleName = "TM254"; Level = 2; Grade = 3; Credits = 30; },
    [PSCustomObject]@{ModuleName = "M250"; Level = 2; Grade = 2; Credits = 30; }
)

OU-Calculator -Modules $modules

