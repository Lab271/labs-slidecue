# SPDX-License-Identifier: Apache-2.0
# SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
# PowerPoint COM Bridge - Maintains PowerPoint COM objects and responds to commands via JSON
# This script runs continuously and accepts commands via stdin, responds via stdout

$ErrorActionPreference = "Stop"

# Global state
$script:pptApp = $null
$script:presentation = $null
$script:slideShow = $null
$script:currentSlide = 1

function Write-Response {
    param([string]$Status, [string]$Data = "", [string]$Error = "")
    
    $response = @{
        status = $Status
        data = $Data
        error = $Error
    } | ConvertTo-Json -Compress
    
    Write-Output $response
    [Console]::Out.Flush()
}

function Invoke-CheckInstalled {
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
        $version = $ppt.Version
        $ppt.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        Write-Response -Status "success" -Data $version
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-OpenPresentation {
    param([string]$FilePath)
    
    try {
        $script:pptApp = New-Object -ComObject PowerPoint.Application
        # Open PowerPoint visible but minimized to avoid stealing focus
        $script:pptApp.Visible = 1  # msoTrue = -1, but 1 also works for visible
        $script:pptApp.WindowState = 2  # ppWindowMinimized = 2
        $script:presentation = $script:pptApp.Presentations.Open($FilePath)
        $slideCount = $script:presentation.Slides.Count
        Write-Response -Status "success" -Data $slideCount.ToString()
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-GetSlideMetadata {
    try {
        if (-not $script:presentation) {
            Write-Response -Status "error" -Error "No presentation open"
            return
        }
        
        $slides = @()
        $totalSlides = $script:presentation.Slides.Count
        
        for ($i = 1; $i -le $totalSlides; $i++) {
            $slide = $script:presentation.Slides.Item($i)
            
            # Check if slide is hidden
            # SlideShowTransition.Hidden property: msoTrue = -1, msoFalse = 0
            $isHidden = $slide.SlideShowTransition.Hidden -eq -1
            
            # Get animation count (approximate - counts effect timings)
            $animationCount = 0
            try {
                if ($slide.TimeLine.MainSequence) {
                    $animationCount = $slide.TimeLine.MainSequence.Count
                }
            } catch {
                # Ignore animation count errors
            }
            
            # Get speaker notes
            $notes = ""
            try {
                if ($slide.HasNotesPage) {
                    $notesPage = $slide.NotesPage
                    $shapes = $notesPage.Shapes
                    foreach ($shape in $shapes) {
                        if ($shape.HasTextFrame) {
                            if ($shape.TextFrame.HasText) {
                                $notes += $shape.TextFrame.TextRange.Text
                            }
                        }
                    }
                    $notes = $notes.Trim()
                }
            } catch {
                # Ignore notes errors
            }
            
            $slideInfo = @{
                slideNumber = $i
                hidden = $isHidden
                animationClicks = $animationCount
                notes = $notes
            }
            $slides += $slideInfo
        }
        
        $result = @{
            totalSlides = $totalSlides
            slides = $slides
        } | ConvertTo-Json -Depth 3 -Compress
        
        Write-Response -Status "success" -Data $result
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-ExportThumbnails {
    param([string]$OutputDir)
    
    try {
        if (-not $script:presentation) {
            Write-Response -Status "error" -Error "No presentation open"
            return
        }
        
        # Get slide dimensions from PageSetup (in points, 72 points = 1 inch)
        $slideWidth = $script:presentation.PageSetup.SlideWidth
        $slideHeight = $script:presentation.PageSetup.SlideHeight
        
        # Calculate export dimensions maintaining aspect ratio
        # Target max width of 1920 pixels
        $maxWidth = 1920
        $aspectRatio = $slideWidth / $slideHeight
        $exportWidth = $maxWidth
        $exportHeight = [int]($maxWidth / $aspectRatio)
        
        $count = $script:presentation.Slides.Count
        for ($i = 1; $i -le $count; $i++) {
            $slide = $script:presentation.Slides.Item($i)
            $fileName = "slide_" + $i.ToString("000") + ".png"
            $filePath = Join-Path $OutputDir $fileName
            $slide.Export($filePath, "PNG", $exportWidth, $exportHeight)
        }
        Write-Response -Status "success" -Data $count.ToString()
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-StartSlideshow {
    try {
        if (-not $script:presentation) {
            Write-Response -Status "error" -Error "No presentation open"
            return
        }
        
        $settings = $script:presentation.SlideShowSettings
        $settings.StartingSlide = 1
        $settings.EndingSlide = $script:presentation.Slides.Count
        
        # Run the slideshow
        $script:slideShow = $settings.Run()
        
        # Wait for slideshow to initialize
        Start-Sleep -Milliseconds 500
        
        # Try to set to windowed mode after starting
        try {
            if ($script:slideShow.View.State -eq 1) {  # ppSlideShowRunning
                # Slideshow is running, now we can work with it
                $script:currentSlide = $script:slideShow.View.CurrentShowPosition
                Write-Response -Status "success" -Data $script:currentSlide.ToString()
            } else {
                Write-Response -Status "error" -Error "Slideshow not in running state"
            }
        } catch {
            # Even if we can't check state, if we have a View, we're probably OK
            if ($script:slideShow.View) {
                $script:currentSlide = $script:slideShow.View.CurrentShowPosition
                Write-Response -Status "success" -Data $script:currentSlide.ToString()
            } else {
                Write-Response -Status "error" -Error "Slideshow.View not available: $($_.Exception.Message)"
            }
        }
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-NextSlide {
    try {
        # Debug info
        $hasSlideShow = $script:slideShow -ne $null
        $hasView = $false
        if ($hasSlideShow) {
            try {
                $hasView = $script:slideShow.View -ne $null
            } catch {
                $hasView = $false
            }
        }
        
        if (-not $hasSlideShow) {
            Write-Response -Status "error" -Error "slideShow object is null"
            return
        }
        
        if (-not $hasView) {
            Write-Response -Status "error" -Error "slideShow.View is null or not accessible"
            return
        }
        
        $script:slideShow.View.Next()
        Start-Sleep -Milliseconds 100
        $script:currentSlide = $script:slideShow.View.CurrentShowPosition
        Write-Response -Status "success" -Data $script:currentSlide.ToString()
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-PreviousSlide {
    try {
        if ($script:slideShow -and $script:slideShow.View) {
            $script:slideShow.View.Previous()
            Start-Sleep -Milliseconds 100
            $script:currentSlide = $script:slideShow.View.CurrentShowPosition
            Write-Response -Status "success" -Data $script:currentSlide.ToString()
        } else {
            Write-Response -Status "error" -Error "No slideshow running or view not available"
        }
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-GotoSlide {
    param([int]$SlideNumber)
    
    try {
        if ($script:slideShow -and $script:slideShow.View) {
            $script:slideShow.View.GotoSlide($SlideNumber)
            $script:currentSlide = $SlideNumber
            Write-Response -Status "success" -Data $SlideNumber.ToString()
        } else {
            Write-Response -Status "error" -Error "No slideshow running"
        }
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-GetCurrentSlide {
    try {
        # Check if we have a slideshow
        if (-not $script:slideShow) {
            Write-Response -Status "success" -Data $script:currentSlide.ToString()
            return
        }
        
        # Try to get current position from View
        if ($script:slideShow.View) {
            $current = $script:slideShow.View.CurrentShowPosition
            $script:currentSlide = $current
            Write-Response -Status "success" -Data $current.ToString()
        } else {
            Write-Response -Status "success" -Data $script:currentSlide.ToString()
        }
    } catch {
        Write-Response -Status "success" -Data $script:currentSlide.ToString()
    }
}

function Invoke-StopSlideshow {
    try {
        if ($script:slideShow) {
            $script:slideShow.View.Exit()
            $script:slideShow = $null
        }
        Write-Response -Status "success"
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-ClosePresentation {
    try {
        if ($script:slideShow) {
            Invoke-StopSlideshow
        }
        if ($script:presentation) {
            $script:presentation.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:presentation) | Out-Null
            $script:presentation = $null
        }
        if ($script:pptApp) {
            $script:pptApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:pptApp) | Out-Null
            $script:pptApp = $null
        }
        Write-Response -Status "success"
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

# Main loop - read commands from stdin
Write-Response -Status "ready"

while ($true) {
    try {
        $line = [Console]::ReadLine()
        if (-not $line) { break }
        
        $command = $line | ConvertFrom-Json
        
        switch ($command.action) {
            "check" { Invoke-CheckInstalled }
            "open" { Invoke-OpenPresentation -FilePath $command.filePath }
            "getMetadata" { Invoke-GetSlideMetadata }
            "export" { Invoke-ExportThumbnails -OutputDir $command.outputDir }
            "start" { Invoke-StartSlideshow }
            "next" { Invoke-NextSlide }
            "previous" { Invoke-PreviousSlide }
            "goto" { Invoke-GotoSlide -SlideNumber $command.slideNumber }
            "getCurrentSlide" { Invoke-GetCurrentSlide }
            "stop" { Invoke-StopSlideshow }
            "close" { Invoke-ClosePresentation }
            "quit" { 
                Invoke-ClosePresentation
                break 
            }
            default { Write-Response -Status "error" -Error "Unknown command: $($command.action)" }
        }
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}
