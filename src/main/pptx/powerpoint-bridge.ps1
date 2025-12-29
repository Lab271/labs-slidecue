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
        $script:pptApp.Visible = 1  # msoTrue = -1, but 1 also works for visible
        $script:presentation = $script:pptApp.Presentations.Open($FilePath)
        $slideCount = $script:presentation.Slides.Count
        Write-Response -Status "success" -Data $slideCount.ToString()
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
        
        $count = $script:presentation.Slides.Count
        for ($i = 1; $i -le $count; $i++) {
            $slide = $script:presentation.Slides.Item($i)
            $fileName = "slide_" + $i.ToString("000") + ".png"
            $filePath = Join-Path $OutputDir $fileName
            $slide.Export($filePath, "PNG", 1920, 1080)
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
        $script:slideShow = $settings.Run()
        $script:currentSlide = 1
        Write-Response -Status "success" -Data "1"
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
    }
}

function Invoke-NextSlide {
    try {
        if ($script:slideShow -and $script:slideShow.View) {
            $script:slideShow.View.Next()
            Start-Sleep -Milliseconds 100
            $script:currentSlide = $script:slideShow.View.CurrentShowPosition
            Write-Response -Status "success" -Data $script:currentSlide.ToString()
        } else {
            Write-Response -Status "error" -Error "No slideshow running"
        }
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
            Write-Response -Status "error" -Error "No slideshow running"
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
        if ($script:slideShow -and $script:slideShow.View) {
            $current = $script:slideShow.View.CurrentShowPosition
            Write-Response -Status "success" -Data $current.ToString()
        } else {
            Write-Response -Status "success" -Data $script:currentSlide.ToString()
        }
    } catch {
        Write-Response -Status "error" -Error $_.Exception.Message
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
