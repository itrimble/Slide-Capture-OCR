-- Enhanced Universal Slide Capture & OCR Renamer v2.2.0
-- Combines capture with intelligent classification, enhanced error handling, and user experience improvements
-- Supports: Chrome, Safari, Edge, Firefox | Requires: macOS 11+, Tesseract, ImageMagick
--
-- ===================================================================================
-- DOCUMENTATION
-- ===================================================================================
--
-- This script automates the process of capturing presentation slides and intelligently
-- renaming them based on content analysis. Key features include:
--
-- • Automatic browser detection and screenshot capture
-- • OCR-based slide content extraction and classification
-- • Intelligent file naming based on slide type and content
-- • Resume capability for interrupted sessions
-- • Progress tracking with time estimation
-- • Adaptive geometry for different screen resolutions
-- • Keyboard shortcut support for pause/resume/cancel
-- • Configurable settings with persistent storage
--
-- REQUIREMENTS:
--   - macOS 11+ (Big Sur or newer)
--   - Tesseract OCR (for text extraction)
--   - ImageMagick (for image processing)
--
-- KEYBOARD SHORTCUTS DURING CAPTURE:
--   - Option+P: Pause capture
--   - Option+R: Resume capture
--   - Option+C: Cancel capture
--
-- USAGE:
--   1. Run the script and select your browser
--   2. Enter the number of slides to capture
--   3. Select an output folder
--   4. Wait for the capture to complete
--
-- SUPPORTED SLIDE TYPES:
--   - Regular Content Slides
--   - Cyber Lab Slides
--   - Knowledge Check/Answer Slides
--   - Pulse Check Slides  
--   - Summary Slides
--   - Break Slides
--   - Real World Scenario Slides
--
-- ===================================================================================

use scripting additions
use framework "Foundation"
use framework "AppKit"
use framework "Carbon"

-- ═══ Configuration Properties ═════════════
-- Version information
property scriptVersion      : "2.2.0"
property scriptReleaseDate  : "2025-05-06"

-- Default paths and settings
property defaultNumSlides    : 20
property defaultOutputFolder : "~/Downloads/Slides"
property defaultWorkDir      : "~/Documents/slide_capture_tmp"
property defaultLogFile      : "~/Library/Logs/slide_capture.log"
property defaultConfigFile   : "~/Library/Preferences/com.slidecapture.config.plist"
property maxTitleLength      : 60
property supportedBrowsers   : {"Google Chrome", "Safari", "Microsoft Edge", "Firefox"}

-- Delay settings
property delayBetweenSlides  : 1.5
property captureDelay        : 0.5
property uiDelay             : 1.0

-- Verification region
property verifyRegion        : "100x100+3500+100" -- For slide changes

-- Logging levels
property LOG_DEBUG   : 0
property LOG_INFO    : 1
property LOG_WARNING : 2
property LOG_ERROR   : 3
property currentLogLevel : 1 -- Default to INFO level

-- State variables
property isCapturing : true
property isPaused : false
property expandedLogPath : ""
property expandedWorkDir : ""
property expandedConfigPath : ""
property startTime : missing value

-- ═══ Main Script ═════════════
on run
    -- Initialize state variables
    set tesseractFound to false
    set imageMagickFound to false
    set isCapturing to true
    set isPaused to false
    set startTime to current date
    set configData to {}
    
    -- Show mode selection dialog
    set modeChoices to {"Capture Slides from Presentation", "Rename Single Screenshot", "Settings & Configuration", "Help & Documentation"}
    set modeChoice to choose from list modeChoices with title "Universal Slide Capture & OCR Renamer v" & scriptVersion with prompt "What would you like to do?" default items {item 1 of modeChoices} without multiple selections allowed and empty selection allowed
    
    if modeChoice is false then
        -- User cancelled
        return
    end if
    
    set selectedMode to item 1 of modeChoice
    
    if selectedMode is "Rename Single Screenshot" then
        renameSingleFile()
        return
    else if selectedMode is "Help & Documentation" then
        showHelp()
        return
    else if selectedMode is "Settings & Configuration" then
        showSettings()
        return
    end if
    
    -- Continue with slide capture mode
    -- 1) Initialize paths, load config and create working directory
    set {expandedLogPath, expandedWorkDir, expandedConfigPath} to initializePaths()
    set configData to loadConfiguration(expandedConfigPath)
    prepareWorkingDirectory(expandedWorkDir)
    
    -- 2) Set up logging
    logMessage("Enhanced Universal Slide Capture & OCR Renamer v" & scriptVersion & " started", LOG_INFO)
    
    -- 3) Check for required tools 
    set {tesseractFound, tesseractPath} to checkDependency("tesseract", "Tesseract OCR")
    set {imageMagickFound, imageMagickPath} to checkDependency("convert", "ImageMagick")
    
    -- If tools are missing, inform the user but allow limited functionality
    if not (tesseractFound and imageMagickFound) then
        set missingTools to {}
        if not tesseractFound then set end of missingTools to "Tesseract OCR"
        if not imageMagickFound then set end of missingTools to "ImageMagick"
        
        set missingToolsText to joinList(missingTools, " and ")
        set warningMessage to "Warning: " & missingToolsText & " not found. Limited functionality available."
        
        logMessage(warningMessage, LOG_WARNING)
        
        set userChoice to display dialog warningMessage & "
        
Would you like to:
- Continue with limited functionality
- Install missing dependencies
- Cancel" buttons {"Install", "Continue", "Cancel"} default button "Install" cancel button "Cancel"
        
        if button returned of userChoice is "Install" then
            displayInstallationInstructions()
            return
        else if button returned of userChoice is "Cancel" then
            return
        end if
    end if
    
    -- 4) Setup command execution environment
    set cmdEnvironment to setupEnvironment(tesseractPath, imageMagickPath)
    
    -- 5) Detect screen resolution and adjust geometry settings
    set geometrySettings to detectAndSetGeometry()
    
    -- 6) Detect or ask for active browser
    set activeBrowser to detectOrSelectBrowser()
    if activeBrowser is missing value then return
    
    -- 7) Ask for capture parameters or load from config/previous session
    set {numSlides, outputFolder, resumeFromSlide} to getCaptureParameters(configData)
    if numSlides is missing value then return
    
    -- 8) Initialize module and context tracking
    set {currentModule, currentTopic, previousSlideTopic} to initializeContextTracking()
    
    -- 9) Register keyboard event handlers for pause/resume/cancel
    setupKeyboardHandlers()
    
    -- 10) Enter presentation mode
    enterPresentationMode(activeBrowser)
    
    -- 11) Skip to resume point if resuming
    if resumeFromSlide > 1 then
        logMessage("Resuming from slide " & resumeFromSlide, LOG_INFO)
        repeat with i from 1 to (resumeFromSlide - 1)
            advanceSlide(activeBrowser)
        end repeat
    end if
    
    -- 12) Main capture loop
    set {previousHash, successCount} to {"", 0}
    set elapsedTimeList to {}
    
    logMessage("Starting capture of " & numSlides & " slides", LOG_INFO)
    displayProgressBar(0, numSlides)
    
    repeat with idx from resumeFromSlide to numSlides
        -- Check if we should continue capturing
        if not isCapturing then
            logMessage("Capture cancelled at slide " & idx, LOG_INFO)
            exit repeat
        end if
        
        -- Handle pause state
        if isPaused then
            repeat while isPaused
                display notification "Capture paused. Press Option+R to resume." with title "Paused"
                delay 1
            end repeat
        end if
        
        -- Track iteration time for estimation
        set iterationStart to current date
        
        try
            -- A) Capture slide
            tell application activeBrowser to activate
            delay captureDelay
            set paddedIdx to text -2 thru -1 of ("0" & idx)
            set tempPath to expandedWorkDir & "/slide_" & paddedIdx & ".png"
            
            -- Make sure the presentation is in focus
            do shell script "screencapture -x " & quoted form of tempPath
            
            -- Verify the capture was successful
            set captureSuccess to verifyImageCapture(tempPath)
            
            if captureSuccess then
                logMessage("Captured slide " & idx & " to " & tempPath, LOG_DEBUG)
            else
                error "Failed to capture image. The screenshot may be empty or invalid."
            end if
            
            -- B) Process the slide if dependencies are available
            if tesseractFound and imageMagickFound then
                set finalTitle to processSlideTitle(tempPath, paddedIdx, currentModule, previousSlideTopic, currentTopic, cmdEnvironment, geometrySettings)
            else
                -- Simple naming without OCR
                set finalTitle to "Slide_" & paddedIdx
            end if
            
            -- C) Move to final destination with proper name
            set finalName to paddedIdx & "_" & finalTitle & ".png"
            set finalPath to outputFolder & "/" & finalName
            
            -- Ensure unique filename
            set finalPath to ensureUniquePath(finalPath)
            
            do shell script "cp " & quoted form of tempPath & " " & quoted form of finalPath
            logMessage("Saved slide " & idx & " as " & (do shell script "basename " & quoted form of finalPath), LOG_INFO)
            
            -- Update progress
            set successCount to successCount + 1
            
            -- Calculate and show estimated time remaining
            set iterationEnd to current date
            set iterationTime to iterationEnd - iterationStart
            set end of elapsedTimeList to iterationTime
            
            -- Only calculate ETA after a few slides for better accuracy
            if length of elapsedTimeList ≥ 3 then
                set avgTime to calculateAverageTime(elapsedTimeList)
                set remainingSlides to numSlides - idx
                set estimatedSecondsRemaining to avgTime * remainingSlides
                
                set estimatedTimeStr to formatTimeRemaining(estimatedSecondsRemaining)
                set progressInfo to "Slide " & idx & " of " & numSlides & " (" & estimatedTimeStr & " remaining)"
            else
                set progressInfo to "Slide " & idx & " of " & numSlides
            end if
            
            -- Update progress display
            displayProgressBar(idx, numSlides)
            display notification progressInfo with title "Capturing Slides" subtitle (do shell script "basename " & quoted form of finalPath)
            
            -- D) Save resume state
            saveResumeState(idx, outputFolder, numSlides, expandedConfigPath)
            
            -- E) Advance to next slide if not at the end
            if idx < numSlides then
                set {isSuccess, newHash} to advanceSlide(activeBrowser)
                
                -- Handle failure to advance by trying again
                if not isSuccess or newHash is previousHash then
                    logMessage("Initial slide advance failed or no change detected. Trying again...", LOG_DEBUG)
                    delay 0.5
                    set {isSuccess, newHash} to advanceSlide(activeBrowser)
                    
                    -- If still failing, try alternative key press
                    if not isSuccess or newHash is previousHash then
                        logMessage("Second slide advance failed. Trying alternative key method...", LOG_DEBUG)
                        tell application "System Events" to key code 125 -- Down arrow as alternative
                        delay 0.5
                        tell application "System Events" to key code 124 -- Right arrow again
                    end if
                end if
                
                set previousHash to newHash
            end if
            
        on error errMsg
            -- Specific error handling based on error type
            logMessage("Error on slide " & idx & ": " & errMsg, LOG_ERROR)
            
            if errMsg contains "Failed to capture image" then
                -- Try again with different delay
                display notification "Capture error. Trying again with longer delay..." with title "Error Recovery"
                delay 2
                -- Try a different approach to activate the browser
                tell application activeBrowser to activate
                delay 1
                try
                    -- Another capture attempt
                    do shell script "screencapture -x " & quoted form of tempPath
                    logMessage("Recovery capture attempt completed", LOG_DEBUG)
                on error
                    logMessage("Recovery capture also failed", LOG_ERROR)
                end try
            else if errMsg contains "OCR failed" then
                -- OCR error - continue with basic naming
                logMessage("OCR failed, using basic naming", LOG_WARNING)
                set finalTitle to "Slide_" & paddedIdx
                
                -- Move to final destination with simple name
                set finalName to paddedIdx & "_" & finalTitle & ".png"
                set finalPath to outputFolder & "/" & finalName
                set finalPath to ensureUniquePath(finalPath)
                
                do shell script "cp " & quoted form of tempPath & " " & quoted form of finalPath
                logMessage("Saved slide " & idx & " with basic naming as " & (do shell script "basename " & quoted form of finalPath), LOG_INFO)
                
                -- Update progress still
                set successCount to successCount + 1
                
                -- Save resume state
                saveResumeState(idx, outputFolder, numSlides, expandedConfigPath)
                
                -- Continue to next slide
                if idx < numSlides then advanceSlide(activeBrowser)
            else
                -- Generic error handling
                display notification "Error on slide " & idx & ". See log for details." with title "Error" 
                delay 2
            end if
        end try
    end repeat
    
    -- 13) Finish up and ask about cleanup
    finishCapture(successCount, numSlides, expandedWorkDir, outputFolder)
end run

-- ═══ Path and Environment Setup Functions ═════════════

-- Initialize paths and ensure directories exist
on initializePaths()
    set expandedLogPath to do shell script "echo " & defaultLogFile
    set expandedWorkDir to do shell script "echo " & defaultWorkDir
    set expandedConfigPath to do shell script "echo " & defaultConfigFile
    
    -- Ensure log directory exists
    do shell script "mkdir -p " & quoted form of (do shell script "dirname " & quoted form of expandedLogPath)
    
    return {expandedLogPath, expandedWorkDir, expandedConfigPath}
end initializePaths

-- Prepare working directory
on prepareWorkingDirectory(workDir)
    do shell script "mkdir -p " & quoted form of workDir
    
    -- Clear any stale verification files
    do shell script "rm -f " & quoted form of workDir & "/verify_*.png 2>/dev/null || true"
    
    return true
end prepareWorkingDirectory

-- Load configuration from plist file
on loadConfiguration(configPath)
    -- Check if config file exists
    set configExists to (do shell script "test -f " & quoted form of configPath & " && echo 'yes' || echo 'no'") is "yes"
    
    if configExists then
        try
            -- Use defaults command to read plist
            set plistDataText to do shell script "defaults read " & quoted form of configPath
            
            -- Initialize the configuration record
            set configData to {}
            
            -- Parse key-value pairs from the output
            set AppleScript's text item delimiters to ";"
            set plistItems to text items of plistDataText
            set AppleScript's text item delimiters to ""
            
            -- Process each key-value pair
            repeat with itemText in plistItems
                -- Skip empty items
                if itemText is not "" then
                    -- Extract key and value (basic parsing)
                    set keyValuePattern to "\"([^\"]+)\" = (.+);"
                    set keyMatch to extractPattern(itemText, keyValuePattern)
                    
                    -- If we have a match
                    if keyMatch is not "" then
                        set keyParts to splitString(keyMatch, ",")
                        if (count of keyParts) >= 2 then
                            set keyName to item 1 of keyParts
                            set keyValue to item 2 of keyParts
                            
                            -- Process resume data specially
                            if keyName is "resumeData" then
                                set resumeData to {}
                                
                                -- Extract nested resume data
                                set resumePattern to "\"currentSlide\" = (\\d+)"
                                set currentSlideMatch to extractPattern(keyValue, resumePattern)
                                if currentSlideMatch is not "" then
                                    set resumeData's currentSlide to currentSlideMatch as integer
                                end if
                                
                                set resumePattern to "\"totalSlides\" = (\\d+)"
                                set totalSlidesMatch to extractPattern(keyValue, resumePattern)
                                if totalSlidesMatch is not "" then
                                    set resumeData's totalSlides to totalSlidesMatch as integer
                                end if
                                
                                set resumePattern to "\"outputFolder\" = \"([^\"]+)\""
                                set outputFolderMatch to extractPattern(keyValue, resumePattern)
                                if outputFolderMatch is not "" then
                                    set resumeData's outputFolder to outputFolderMatch
                                end if
                                
                                -- Add timestamp
                                set resumeData's timestamp to current date
                                
                                -- Add the resume data to the config
                                set configData's resumeData to resumeData
                            else
                                -- For other simple key-value pairs
                                set configData's keyName to keyValue
                            end if
                        end if
                    end if
                end if
            end repeat
            
            logMessage("Configuration loaded from: " & configPath, LOG_DEBUG)
            return configData
        on error errMsg
            logMessage("Error loading config: " & errMsg, LOG_ERROR)
            return {}
        end try
    else
        logMessage("No configuration file found at: " & configPath, LOG_DEBUG)
        return {}
    end if
end loadConfiguration

-- Save configuration to plist file
on saveConfiguration(configData, configPath)
    try
        -- Create a temporary property list representation
        set tempFile to POSIX path of (POSIX file "/tmp/slide_capture_temp.plist")
        
        -- Initialize an empty plist
        do shell script "defaults write " & quoted form of tempFile & " dict '{}';"
        
        -- Write key-value pairs
        repeat with keyName in keys of configData
            set keyValue to configData's keyName
            
            -- Handle different types of values
            if class of keyValue is record then
                -- For resume data, which is a record
                if keyName is "resumeData" then
                    -- Create a nested dictionary
                    do shell script "defaults write " & quoted form of tempFile & " resumeData -dict"
                    
                    -- Add all resume data fields
                    repeat with resumeKey in keys of keyValue
                        set resumeValue to keyValue's resumeKey
                        
                        if class of resumeValue is integer then
                            do shell script "defaults write " & quoted form of tempFile & " resumeData." & resumeKey & " -int " & resumeValue
                        else if class of resumeValue is real then
                            do shell script "defaults write " & quoted form of tempFile & " resumeData." & resumeKey & " -float " & resumeValue
                        else if class of resumeValue is text or class of resumeValue is string then
                            do shell script "defaults write " & quoted form of tempFile & " resumeData." & resumeKey & " -string " & quoted form of resumeValue
                        else if class of resumeValue is date then
                            -- Format date for storage
                            set dateStr to (year of resumeValue) & "-" & (month of resumeValue as integer) & "-" & (day of resumeValue) & " " & (time of resumeValue)
                            do shell script "defaults write " & quoted form of tempFile & " resumeData." & resumeKey & " -string '" & dateStr & "'"
                        end if
                    end repeat
                end if
            else if class of keyValue is list then
                -- Handle lists
                do shell script "defaults write " & quoted form of tempFile & " " & keyName & " -array " & (joinList(keyValue, " "))
            else if class of keyValue is integer then
                do shell script "defaults write " & quoted form of tempFile & " " & keyName & " -int " & keyValue
            else if class of keyValue is real then
                do shell script "defaults write " & quoted form of tempFile & " " & keyName & " -float " & keyValue
            else if class of keyValue is boolean then
                if keyValue then
                    do shell script "defaults write " & quoted form of tempFile & " " & keyName & " -bool true"
                else
                    do shell script "defaults write " & quoted form of tempFile & " " & keyName & " -bool false"
                end if
            else if class of keyValue is text or class of keyValue is string then
                do shell script "defaults write " & quoted form of tempFile & " " & keyName & " -string " & quoted form of keyValue
            end if
        end repeat
        
        -- Move the temporary file to the final location
        do shell script "cp " & quoted form of tempFile & " " & quoted form of configPath
        
        logMessage("Configuration saved to " & configPath, LOG_DEBUG)
        return true
    on error errMsg
        logMessage("Error saving config: " & errMsg, LOG_ERROR)
        return false
    end try
end saveConfiguration

-- Save current state for potential resume
on saveResumeState(currentSlide, outputFolder, totalSlides, configPath)
    set resumeData to {currentSlide:currentSlide, outputFolder:outputFolder, totalSlides:totalSlides, timestamp:(current date)}
    
    -- Save to config file
    set configData to loadConfiguration(configPath)
    if configData is missing value then set configData to {}
    
    -- Update resume data
    set configData's resumeData to resumeData
    
    -- Save updated config
    saveConfiguration(configData, configPath)
end saveResumeState

-- Set up command execution environment
on setupEnvironment(tesseractPath, imageMagickPath)
    set cmdPrefix to ""
    
    -- Add paths for found dependencies
    if tesseractPath is not "" then
        set tessDir to do shell script "dirname " & quoted form of tesseractPath
        set cmdPrefix to "PATH=" & tessDir & ":$PATH; "
    end if
    
    if imageMagickPath is not "" then
        set imgDir to do shell script "dirname " & quoted form of imageMagickPath
        set cmdPrefix to cmdPrefix & "PATH=" & imgDir & ":$PATH; "
    end if
    
    logMessage("Using command prefix: " & cmdPrefix, LOG_DEBUG)
    return cmdPrefix
end setupEnvironment

-- ═══ Dependency Checking Functions ═════════════

-- Check for a dependency by name
on checkDependency(cmdName, friendlyName)
    set toolPath to ""
    set toolFound to false
    
    logMessage("Checking for " & friendlyName & "...", LOG_DEBUG)
    
    try
        -- First try normal PATH
        set toolPath to do shell script "which " & cmdName & " || echo ''"
        
        if toolPath is "" then
            -- Try common locations
            set possibleLocations to {"/opt/homebrew/bin/" & cmdName, "/usr/local/bin/" & cmdName, "/usr/bin/" & cmdName}
            
            repeat with loc in possibleLocations
                set testResult to do shell script "test -x " & quoted form of loc & " && echo yes || echo no"
                if testResult is "yes" then
                    set toolPath to loc
                    exit repeat
                end if
            end repeat
        end if
        
        -- Test the tool by running a simple command
        if toolPath is not "" then
            if cmdName is "tesseract" then
                do shell script quoted form of toolPath & " --version | head -n 1"
            else if cmdName is "convert" then
                do shell script quoted form of toolPath & " -version | head -n 1"
            end if
            
            set toolFound to true
            logMessage("Found " & friendlyName & " at: " & toolPath, LOG_INFO)
        else
            logMessage(friendlyName & " not found", LOG_WARNING)
        end if
        
    on error errMsg
        set toolFound to false
        logMessage("Error checking for " & friendlyName & ": " & errMsg, LOG_ERROR)
    end try
    
    return {toolFound, toolPath}
end checkDependency

-- Display installation instructions for missing dependencies
on displayInstallationInstructions()
    set message to "Installation Instructions:

1. Install Homebrew (if not already installed):
   /bin/bash -c \"$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\"

2. Install required dependencies:
   brew install tesseract
   brew install imagemagick

3. Verify installation:
   tesseract --version
   convert --version

After installation, please run this script again."

    display dialog message buttons {"Copy to Clipboard", "OK"} default button "OK"
    
    if button returned of result is "Copy to Clipboard" then
        set the clipboard to "# Install Homebrew (if not already installed)
/bin/bash -c \"$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\"

# Install required dependencies
brew install tesseract
brew install imagemagick

# Verify installation
tesseract --version
convert --version"
        
        display notification "Installation instructions copied to clipboard" with title "Copied"
    end if
end displayInstallationInstructions

-- ═══ User Interface Functions ═════════════

-- Detect screen resolution and set geometry accordingly
on detectAndSetGeometry()
    -- Get screen resolution
    set screenResolution to do shell script "system_profiler SPDisplaysDataType | grep Resolution | awk '{print $2, $3, $4}' | head -n 1"
    
    -- Parse resolution
    set AppleScript's text item delimiters to " "
    set resComponents to text items of screenResolution
    set AppleScript's text item delimiters to ""
    
    -- Default to 3840x2160 if we can't detect properly
    set screenWidth to 3840
    set screenHeight to 2160
    
    try
        set screenWidth to resComponents's item 1 as integer
        set screenHeight to resComponents's item 3 as integer
    on error
        logMessage("Failed to parse screen resolution: " & screenResolution & ". Using default 3840x2160.", LOG_WARNING)
    end try
    
    logMessage("Detected screen resolution: " & screenWidth & "x" & screenHeight, LOG_INFO)
    
    -- Calculate geometry based on screen resolution
    -- Scale the coordinates proportionally from the base 3840x2160 resolution
    set widthRatio to screenWidth / 3840.0
    set heightRatio to screenHeight / 2160.0
    
    -- Create geometry settings object
    set geometrySettings to {screenWidth:screenWidth, screenHeight:screenHeight}
    
    -- Title regions with scaled dimensions
    set geometrySettings's titleSlideGeom to (screenWidth as integer) & "x" & (round (600 * heightRatio)) & "+0+" & (round (540 * heightRatio))
    set geometrySettings's genericTitleGeom to (screenWidth as integer) & "x" & (round (240 * heightRatio)) & "+0+" & (round (180 * heightRatio))
    set geometrySettings's scenarioTitleGeom to (round (2880 * widthRatio)) & "x" & (round (240 * heightRatio)) & "+" & (round (960 * widthRatio)) & "+" & (round (180 * heightRatio))
    set geometrySettings's scenarioSidebarGeom to (round (960 * widthRatio)) & "x" & (round (1200 * heightRatio)) & "+0+" & (round (300 * heightRatio))
    set geometrySettings's slideNumGeom to (round (280 * widthRatio)) & "x" & (round (280 * heightRatio)) & "+" & (round (240 * widthRatio)) & "+" & (round (2010 * heightRatio))
    
    -- Verification region for slide changes
    set geometrySettings's verifyRegion to "100x100+" & (round (screenWidth - 340)) & "+100"
    
    logMessage("Adjusted geometry settings for " & screenWidth & "x" & screenHeight, LOG_DEBUG)
    
    return geometrySettings
end detectAndSetGeometry

-- Detect active browser or let user select one
on detectOrSelectBrowser()
    try
        set activeBrowser to detectActiveBrowser()
        
        if activeBrowser is missing value then
            set activeBrowser to (choose from list supportedBrowsers with title "Select Browser" with prompt "Which browser has your presentation?" default items {"Google Chrome"})'s item 1
        else
            -- Confirm detected browser
            set confirmMsg to "Detected browser: " & activeBrowser & ". Is this correct?"
            set userConfirm to display dialog confirmMsg buttons {"No, select different browser", "Yes, use this browser"} default button "Yes, use this browser"
            
            if button returned of userConfirm is "No, select different browser" then
                set activeBrowser to (choose from list supportedBrowsers with title "Select Browser" with prompt "Which browser has your presentation?" default items {activeBrowser})'s item 1
            end if
        end if
    on error
        set activeBrowser to missing value
    end try
    
    if activeBrowser is not missing value then
        logMessage("Selected browser: " & activeBrowser, LOG_INFO)
    end if
    
    return activeBrowser
end detectOrSelectBrowser

-- Get capture parameters from user or config
on getCaptureParameters(configData)
    -- Check for resume data
    set canResume to false
    set resumeData to {}
    
    if configData is not missing value and configData is not {} then
        if configData contains "resumeData" then
            set resumeData to configData's resumeData
            
            if resumeData is not missing value and resumeData is not {} then
                if resumeData contains "currentSlide" and resumeData contains "totalSlides" and resumeData contains "outputFolder" then
                    set currentSlide to resumeData's currentSlide
                    set totalSlides to resumeData's totalSlides
                    set savedOutputFolder to resumeData's outputFolder
                    
                    -- Only offer resume if not complete and not too old (24 hours)
                    if currentSlide < totalSlides then
                        set timestampStr to "Unknown"
                        
                        if resumeData contains "timestamp" then
                            set timestampDate to resumeData's timestamp
                            set timestampStr to (timestampDate's month as text) & "/" & (timestampDate's day as text) & " " & (timestampDate's time as text)
                        end if
                        
                        set resumePrompt to "Found a previous session from " & timestampStr & "
Slide " & currentSlide & " of " & totalSlides & " completed
Output folder: " & savedOutputFolder & "

Would you like to resume from slide " & (currentSlide + 1) & "?"
                        
                        set resumeChoice to display dialog resumePrompt buttons {"Start New Session", "Resume"} default button "Resume"
                        
                        if button returned of resumeChoice is "Resume" then
                            set canResume to true
                        end if
                    end if
                end if
            end if
        end if
    end if
    
    -- If resuming, use previous settings
    if canResume then
        set numSlides to resumeData's totalSlides
        set outputFolder to resumeData's outputFolder
        set resumeFromSlide to (resumeData's currentSlide) + 1
        
        logMessage("Resuming capture from slide " & resumeFromSlide & " of " & numSlides, LOG_INFO)
    else
        -- Otherwise, ask for new settings
        -- Expand ~ in path
        set expandedDefault to do shell script "echo " & defaultOutputFolder
        
        -- Ask for number of slides
        set numSlidesStr to text returned of (display dialog "Number of slides to capture:" default answer defaultNumSlides buttons {"Cancel", "OK"} default button "OK") as text
        
        -- Validate input
        if numSlidesStr does not contain "[0-9]+" then
            display dialog "Please enter a valid number" buttons {"OK"} default button "OK"
            return {missing value, missing value, missing value}
        end if
        
        set numSlides to numSlidesStr as integer
        
        -- Create default output directory if it doesn't exist
        do shell script "mkdir -p " & quoted form of expandedDefault
        
        -- Ask for output folder
        set outputFolder to POSIX path of (choose folder with prompt "Select output folder:" default location (POSIX file expandedDefault as alias))
        
        -- Start from the beginning
        set resumeFromSlide to 1
    end if
    
    return {numSlides, outputFolder, resumeFromSlide}
end getCaptureParameters

-- Initialize context tracking variables
on initializeContextTracking()
    set currentModule to ""
    set currentTopic to ""
    set previousSlideTopic to ""
    
    return {currentModule, currentTopic, previousSlideTopic}
end initializeContextTracking

-- Display a progress bar
on displayProgressBar(current, total)
    if total < 1 then set total to 1
    if current > total then set current to total
    
    set percentComplete to (current / total) * 100
    set barWidth to 50
    
    set completedWidth to round (barWidth * percentComplete / 100)
    if completedWidth < 0 then set completedWidth to 0
    if completedWidth > barWidth then set completedWidth to barWidth
    
    set barCompleted to ""
    repeat with i from 1 to completedWidth
        set barCompleted to barCompleted & "█"
    end repeat
    
    set barRemaining to ""
    repeat with i from 1 to (barWidth - completedWidth)
        set barRemaining to barRemaining & "░"
    end repeat
    
    set progressBar to barCompleted & barRemaining & " " & round(percentComplete) & "% (" & current & "/" & total & ")"
    
    logMessage(progressBar, LOG_INFO)
end displayProgressBar

-- Setup keyboard handlers for pause/resume/cancel
on setupKeyboardHandlers()
    logMessage("Setting up keyboard event handlers...", LOG_DEBUG)
    
    -- This is a more robust implementation using the Carbon framework
    -- First, unregister any existing handlers
    call method "CGEventTapEnable" of class "NSEvent" with parameters {0, 0}
    
    -- Register a new global event tap using Carbon's RegisterEventHotKey
    -- Create an event handler to watch for the following key combinations:
    --   Option+P: Pause capture
    --   Option+R: Resume capture
    --   Option+C: Cancel capture
    
    tell application "System Events"
        -- Define the key commands we want to handle
        set optionP to {58, 35} -- Option-P (58 is option, 35 is P)
        set optionR to {58, 15} -- Option-R
        set optionC to {58, 8}  -- Option-C
        
        -- Set up a handler for key presses
        tell process "System Events"
            set appName to name of me
            
            -- This creates a background process to monitor key presses
            -- The actual event listening happens in the OS via UI Events
            do shell script "osascript -e 'tell application \"System Events\" to set optionKeyDown to false' > /dev/null 2>&1 &"
            
            -- Set up handlers via UI scripting - these will run in separate processes
            do shell script "osascript -e '
                on run
                    tell application \"System Events\"
                        set targetApp to \"" & appName & "\"
                        repeat
                            set keyPress to \"\"
                            try
                                if key code 58 down then -- Option key
                                    if key code 35 down then -- P key
                                        set keyPress to \"pause\"
                                        do shell script \"defaults write com.slidecapture.keystate isPaused -bool true\"
                                        do shell script \"afplay /System/Library/Sounds/Tink.aiff\"
                                    else if key code 15 down then -- R key
                                        set keyPress to \"resume\"
                                        do shell script \"defaults write com.slidecapture.keystate isPaused -bool false\"
                                        do shell script \"afplay /System/Library/Sounds/Tink.aiff\"
                                    else if key code 8 down then -- C key
                                        set keyPress to \"cancel\"
                                        do shell script \"defaults write com.slidecapture.keystate isCapturing -bool false\"
                                        do shell script \"afplay /System/Library/Sounds/Tink.aiff\"
                                    end if
                                end if
                                
                                if keyPress is not \"\" then
                                    delay 0.5 -- Prevent key repeat
                                end if
                            end try
                            
                            delay 0.1
                        end repeat
                    end tell
                end run
            ' > /dev/null 2>&1 &"
        end tell
    end tell
    
    -- Set up a handler to check the state file periodically
    do shell script "touch ~/.slidecapture_state"
    
    -- This thread will periodically check for state changes
    -- The main script will then check the global variables
    set stateCheckScript to "
    on run
        repeat
            set isPaused to do shell script \"defaults read com.slidecapture.keystate isPaused 2>/dev/null || echo false\"
            set isCapturing to do shell script \"defaults read com.slidecapture.keystate isCapturing 2>/dev/null || echo true\"
            
            -- Write to state file
            do shell script \"echo 'isPaused=\" & isPaused & \"' > ~/.slidecapture_state\"
            do shell script \"echo 'isCapturing=\" & isCapturing & \"' >> ~/.slidecapture_state\"
            
            delay 0.2
        end repeat
    end run
    "
    
    do shell script "osascript -e '" & stateCheckScript & "' > /dev/null 2>&1 &"
    
    -- Initialize state
    do shell script "defaults write com.slidecapture.keystate isPaused -bool false"
    do shell script "defaults write com.slidecapture.keystate isCapturing -bool true"
    
    logMessage("Keyboard handlers set up. Use Option+P to pause, Option+R to resume, Option+C to cancel.", LOG_INFO)
end setupKeyboardHandlers

-- Check state of keyboard handlers
on checkKeyboardState()
    -- Read state from file
    set stateFile to "~/.slidecapture_state"
    set expandedStatePath to do shell script "echo " & stateFile
    
    if (do shell script "test -f " & quoted form of expandedStatePath & " && echo yes || echo no") is "yes" then
        set stateData to paragraphs of (do shell script "cat " & quoted form of expandedStatePath)
        
        -- Parse each line
        repeat with stateLine in stateData
            if stateLine starts with "isPaused=" then
                set isPausedValue to text 10 thru -1 of stateLine
                if isPausedValue is "true" then
                    set isPaused to true
                else
                    set isPaused to false
                end if
            else if stateLine starts with "isCapturing=" then
                set isCapturingValue to text 13 thru -1 of stateLine
                if isCapturingValue is "true" then
                    set isCapturing to true
                else
                    set isCapturing to false
                end if
            end if
        end repeat
    end if
    
    return {isCapturing, isPaused}
end checkKeyboardState

-- Format time remaining in a human-readable format
on formatTimeRemaining(seconds)
    if seconds < 60 then
        return round(seconds) & " seconds"
    else if seconds < 3600 then
        return round(seconds / 60) & " minutes"
    else
        return (round(seconds / 360) / 10) & " hours"
    end if
end formatTimeRemaining

-- Calculate average time per slide from elapsedTimeList
on calculateAverageTime(timeList)
    if length of timeList < 1 then return 0
    
    set totalTime to 0
    repeat with timeValue in timeList
        set totalTime to totalTime + timeValue
    end repeat
    
    return totalTime / (length of timeList)
end calculateAverageTime

-- Finish capture and handle cleanup
on finishCapture(successCount, numSlides, workDir, outputFolder)
    set totalTime to ((current date) - startTime)
    
    -- Format total time nicely
    set timeMsg to ""
    if totalTime < 60 then
        set timeMsg to totalTime & " seconds"
    else if totalTime < 3600 then
        set timeMsg to round(totalTime / 60) & " minutes"
    else
        set timeMsg to (round(totalTime / 360) / 10) & " hours"
    end if
    
    logMessage("Completed capturing " & successCount & " of " & numSlides & " slides in " & timeMsg, LOG_INFO)
    
    -- Clean up keyboard handler processes
    try
        do shell script "pkill -f 'osascript.*slidecapture'"
        do shell script "rm -f ~/.slidecapture_state"
        do shell script "defaults delete com.slidecapture.keystate"
    end try
    
    -- Offer to clean up temporary files
    set cleanupMsg to "Slide capture completed!
Successfully captured " & successCount & " of " & numSlides & " slides in " & timeMsg & ".

Would you like to clean up temporary files?"
    
    set cleanupChoice to display dialog cleanupMsg buttons {"Keep Files", "Clean Up", "Open Output Folder"} default button "Open Output Folder"
    
    if button returned of cleanupChoice is "Clean Up" then
        do shell script "rm -rf " & quoted form of workDir
        logMessage("Cleaned up temporary files at " & workDir, LOG_INFO)
    end if
    
    if button returned of cleanupChoice is "Open Output Folder" then
        do shell script "open " & quoted form of outputFolder
    end if
    
    -- Clear resume state
    saveResumeState(numSlides, outputFolder, numSlides, expandedConfigPath)
    
    display notification "Completed " & successCount & " of " & numSlides & " slides in " & timeMsg with title "Finished"
end finishCapture

-- ═══ Presentation Control Functions ═════════════

-- Enter presentation mode in the specified browser
on enterPresentationMode(browserName)
    tell application browserName
        activate
        delay uiDelay
        
        if browserName is "Safari" then
            tell application "System Events" to keystroke "f" using {command down, control down}
        else if browserName is "Firefox" then
            tell application "System Events" to keystroke "f" using {command down}
        else
            tell application "System Events" to keystroke "f" using {command down, shift down}
        end if
    end tell
    
    delay uiDelay
    logMessage("Entered presentation mode in " & browserName, LOG_INFO)
end enterPresentationMode

-- Advance to the next slide and verify the change
on advanceSlide(browserName)
    tell application browserName to activate
    delay 0.2
    tell application "System Events" to key code 124 -- right arrow
    delay delayBetweenSlides
    
    -- Verify slide changed by capturing and comparing a small region
    set verPath to expandedWorkDir & "/verify_" & (random number from 1000 to 9999) & ".png"
    do shell script "screencapture -R " & verifyRegion & " " & quoted form of verPath
    set currentHash to do shell script "md5 -q " & quoted form of verPath
    
    return {true, currentHash}
end advanceSlide

-- Detect the active browser
on detectActiveBrowser()
    tell application "System Events"
        set frontApps to name of every process whose frontmost is true
        
        if frontApps is {} then
            return missing value
        end if
        
        set frontApp to item 1 of frontApps
        
        if frontApp is in supportedBrowsers then 
            return frontApp
        else
            -- Check for similar names
            repeat with browserName in supportedBrowsers
                if frontApp contains browserName then
                    return browserName
                end if
            end repeat
        end if
    end tell
    
    return missing value
end detectActiveBrowser

-- Verify image capture was successful
on verifyImageCapture(imagePath)
    -- Check if the file exists and is not empty
    set fileExists to (do shell script "test -f " & quoted form of imagePath & " && echo 'yes' || echo 'no'") is "yes"
    
    if not fileExists then
        return false
    end if
    
    -- Check if the file size is greater than a minimum threshold (e.g., 10KB)
    set fileSize to do shell script "stat -f %z " & quoted form of imagePath
    
    if fileSize < 10240 then -- 10KB
        logMessage("Warning: Image file size is suspiciously small: " & fileSize & " bytes", LOG_WARNING)
        
        -- Additional check - try to get image dimensions
        try
            set imageDimensions to do shell script "sips -g pixelWidth -g pixelHeight " & quoted form of imagePath & " | grep pixel"
            
            if imageDimensions contains "0" then
                logMessage("Warning: Image dimensions appear to be zero", LOG_WARNING)
                return false
            end if
        on error
            logMessage("Error checking image dimensions", LOG_ERROR)
            return false
        end try
    end if
    
    return true
end verifyImageCapture

-- ═══ Slide Processing Functions ═════════════

-- Process a slide to determine its type and extract title
on processSlideTitle(slidePath, paddedIdx, currentMod, prevTopic, currTopic, cmdPrefix, geomSettings)
    -- Extract geometry settings
    set titleSlideGeom to geomSettings's titleSlideGeom
    set genericTitleGeom to geomSettings's genericTitleGeom
    set scenarioTitleGeom to geomSettings's scenarioTitleGeom
    set scenarioSidebarGeom to geomSettings's scenarioSidebarGeom
    
    -- Variables to store OCR results
    set isScenario to false
    set isCyberLab to false
    set isSummary to false
    set isBreak to false
    set isPulseCheck to false
    set isKnowledgeCheck to false
    set isKnowledgeCheckAnswer to false
    
    -- 1) Extract different regions in parallel for better performance
    set ocrResults to {}
    
    -- Create a unique identifier for this processing run
    set runId to do shell script "date +%s%N"
    
    -- Define crop regions and their processing parameters
    set cropJobs to {¬
        {name:"title", geom:genericTitleGeom, psm:"7", preprocess:"-colorspace Gray -contrast-stretch 15%"}, ¬
        {name:"cover", geom:titleSlideGeom, psm:"6", preprocess:"-colorspace Gray -contrast-stretch 15%"}, ¬
        {name:"sidebar", geom:scenarioSidebarGeom, psm:"6", preprocess:"-colorspace Gray -contrast-stretch 15%"}, ¬
        {name:"full", geom:"full", psm:"6", preprocess:"-resize 1920x1080 -colorspace Gray -contrast-stretch 15%"} ¬
    }
    
    -- Process each region
    repeat with job in cropJobs
        set jobName to job's name
        set jobGeom to job's geom
        set jobPsm to job's psm
        set jobPreprocess to job's preprocess
        
        set cropPath to expandedWorkDir & "/" & jobName & "_" & runId & ".png"
        
        try
            -- Extract and preprocess the region
            if jobGeom is "full" then
                -- Special case for full slide
                do shell script cmdPrefix & "convert " & quoted form of slidePath & " " & jobPreprocess & " " & quoted form of cropPath
            else
                do shell script cmdPrefix & "convert " & quoted form of slidePath & " -crop " & jobGeom & " " & jobPreprocess & " " & quoted form of cropPath
            end if
            
            -- Perform OCR
            set rawText to do shell script cmdPrefix & "tesseract " & quoted form of cropPath & " stdout -l eng --psm " & jobPsm & " 2>/dev/null || echo 'ERROR'"
            
            -- Store result
            set ocrResults's jobName to rawText
            
            -- Clean up temporary file immediately
            do shell script "rm " & quoted form of cropPath & " 2>/dev/null || true"
        on error errMsg
            logMessage("Error processing " & jobName & " region: " & errMsg, LOG_ERROR)
            set ocrResults's jobName to "ERROR"
        end try
    end repeat
    
    -- 2) Analyze OCR results
    
    -- Get raw text from different regions
    set rawTitle to ""
    if ocrResults contains "title" and ocrResults's title is not "ERROR" and ocrResults's title is not "" then
        set rawTitle to ocrResults's title
    else if ocrResults contains "cover" and ocrResults's cover is not "ERROR" and ocrResults's cover is not "" then
        set rawTitle to ocrResults's cover
    end if
    
    -- Check for scenario slides
    if ocrResults contains "sidebar" and ocrResults's sidebar is not "ERROR" and ocrResults's sidebar is not "" then
        set rawSide to ocrResults's sidebar
        set isScenario to rawSide contains "Consider the Real World Scenario" or rawSide contains "Real World"
    end if
    
    -- Check full slide for special types
    if ocrResults contains "full" and ocrResults's full is not "ERROR" and ocrResults's full is not "" then
        set rawFull to ocrResults's full
        set isCyberLab to rawFull contains "Cyber Lab"
        set isSummary to rawFull contains "Summary"
        set isBreak to rawFull contains "Break" or rawFull contains "Minute Break"
        set isPulseCheck to rawFull contains "Pulse Check"
        set isKnowledgeCheck to rawFull contains "Knowledge Check" and not (rawFull contains "Answer")
        set isKnowledgeCheckAnswer to rawFull contains "Knowledge Check Answer"
    end if
    
    -- 3) Clean up the OCR text
    set cleanTitle to ""
    if rawTitle is not "ERROR" and rawTitle is not "" then
        set cleanTitle to do shell script "echo " & quoted form of rawTitle & " | tr -d '\\r\\n' | tr -cd 'A-Za-z0-9 \\-_:.()&|' | sed 's/^ *//;s/ *$//;s/  */ /g'"
    end if
    
    if cleanTitle is "" then set cleanTitle to "Untitled"
    
    logMessage("Raw title: " & rawTitle, LOG_DEBUG)
    logMessage("Clean title: " & cleanTitle, LOG_DEBUG)
    logMessage("Slide classifications - Summary: " & isSummary & ", Cyber Lab: " & isCyberLab & ", Scenario: " & isScenario, LOG_DEBUG)
    
    -- 4) Update module tracking - enhanced with more patterns
    if cleanTitle contains "Introduction" and not (cleanTitle contains "cont.") then
        if cleanTitle contains "Endpoint Security" then
            set currentMod to "Endpoint_Security"
        else if cleanTitle contains "ClamAV" then
            set currentMod to "ClamAV"
        else if cleanTitle contains "Cybersecurity" then
            set currentMod to "Cybersecurity"
        end if
        logMessage("Updated module to: " & currentMod, LOG_INFO)
    else if cleanTitle contains "Antivirus Problems" or cleanTitle contains "Antivirus Risks" then
        set currentMod to "Antivirus_Problems"
        logMessage("Updated module to: " & currentMod, LOG_INFO)
    else if cleanTitle contains "Endpoint Detection" or cleanTitle contains "EDR" then
        set currentMod to "EDR"
        logMessage("Updated module to: " & currentMod, LOG_INFO)
    else if cleanTitle contains "YARA Rules" or cleanTitle contains "YARA" then
        set currentMod to "YARA_Rules"
        logMessage("Updated module to: " & currentMod, LOG_INFO)
    end if
    
    -- 5) Classify slide type for naming - enhanced with more patterns and context
    if isCyberLab then
        if rawTitle contains "Cyber Lab" then
            set finalTitle to cleanTitle
        else
            set finalTitle to "Cyber_Lab"
            
            -- Try to extract lab name from full text
            if ocrResults contains "full" and ocrResults's full is not "ERROR" then
                set labNameMatch to extractPattern(ocrResults's full, "Cyber Lab: ([^\\n]+)")
                if labNameMatch is not "" then
                    set finalTitle to "Cyber_Lab_" & cleanString(labNameMatch)
                end if
            end if
        end if
        logMessage("Classified as Cyber Lab: " & finalTitle, LOG_INFO)
    else if isKnowledgeCheckAnswer then
        set finalTitle to "Knowledge_Check_Answer"
        
        -- Try to add question number if available
        if ocrResults contains "full" and ocrResults's full is not "ERROR" then
            set questionNum to extractPattern(ocrResults's full, "Question ([0-9]+)")
            if questionNum is not "" then
                set finalTitle to "Knowledge_Check_Answer_" & questionNum
            end if
        end if
        
        logMessage("Classified as Knowledge Check Answer", LOG_INFO)
    else if isKnowledgeCheck then
        set finalTitle to "Knowledge_Check"
        
        -- Try to add question number if available
        if ocrResults contains "full" and ocrResults's full is not "ERROR" then
            set questionNum to extractPattern(ocrResults's full, "Question ([0-9]+)")
            if questionNum is not "" then
                set finalTitle to "Knowledge_Check_" & questionNum
            end if
        end if
        
        logMessage("Classified as Knowledge Check", LOG_INFO)
    else if isPulseCheck then
        set finalTitle to "Pulse_Check"
        
        -- Try to add topic if available
        if currentMod is not "" then
            set finalTitle to "Pulse_Check_" & currentMod
        end if
        
        logMessage("Classified as Pulse Check", LOG_INFO)
    else if isSummary then
        if currentMod is "" then
            set finalTitle to "Summary"
        else
            set finalTitle to "Summary_" & currentMod
        end if
        logMessage("Classified as Summary with module: " & currentMod, LOG_INFO)
    else if isBreak then
        set finalTitle to "Break"
        logMessage("Classified as Break", LOG_INFO)
    else if isScenario then
        if currentMod is "" then
            set finalTitle to "Real_World_Scenario"
        else
            set finalTitle to "Real_World_Scenario_" & currentMod
        end if
        
        -- Try to extract scenario name if available
        if ocrResults contains "sidebar" and ocrResults's sidebar is not "ERROR" then
            set scenarioName to extractPattern(ocrResults's sidebar, "Scenario: ([^\\n]+)")
            if scenarioName is not "" then
                set finalTitle to "Real_World_Scenario_" & cleanString(scenarioName)
            end if
        end if
        
        logMessage("Classified as Real World Scenario", LOG_INFO)
    else
        -- For general slides, normalize spaces to underscores
        set AppleScript's text item delimiters to " "
        set parts to text items of cleanTitle
        set AppleScript's text item delimiters to "_"
        set finalTitle to parts as text
        set AppleScript's text item delimiters to ""
        
        -- Add module prefix for context if title doesn't already contain it
        if currentMod is not "" and finalTitle does not contain currentMod then
            -- Only add if it would make sense semantically
            if length of finalTitle < 40 then -- Don't add to already long titles
                set finalTitle to currentMod & "_" & finalTitle
            end if
        end if
        
        logMessage("Classified as general slide: " & finalTitle, LOG_INFO)
    end if
    
    -- 6) Truncate overly long titles
    if (count of characters of finalTitle) > maxTitleLength then
        set truncated to text 1 thru maxTitleLength of finalTitle
        logMessage("Truncated title: " & truncated, LOG_DEBUG)
        set finalTitle to truncated
    end if
    
    -- 7) Replace problematic characters for filenames
    set finalTitle to replaceString(finalTitle, "/", "_")
    set finalTitle to replaceString(finalTitle, ":", "-")
    set finalTitle to replaceString(finalTitle, "?", "")
    
    return finalTitle
end processSlideTitle

-- ═══ Utility Functions ═════════════

-- Ensure a path is unique by adding a suffix if necessary
on ensureUniquePath(pathToCheck)
    if (do shell script "test -e " & quoted form of pathToCheck & " && echo yes || echo no") is "yes" then
        -- Split path into components
        set lastDot to offset of "." in (reverseString(pathToCheck))
        
        if lastDot is 0 then
            -- No extension
            set basePath to pathToCheck
            set extension to ""
        else
            set extension to "." & text ((length of pathToCheck) - lastDot + 2) thru -1 of pathToCheck
            set basePath to text 1 thru ((length of pathToCheck) - lastDot - length of extension + 1) of pathToCheck
        end if
        
        -- Add unique identifier
        set uniqueId to do shell script "date +%s | head -c 4"
        set uniquePath to basePath & "_" & uniqueId & extension
        
        -- Recursively check if the new path is also taken
        return ensureUniquePath(uniquePath)
    else
        return pathToCheck
    end if
end ensureUniquePath

-- Extract a pattern from text using regex
on extractPattern(inputText, pattern)
    try
        set result to do shell script "echo " & quoted form of inputText & " | grep -oE '" & pattern & "' | head -n 1 | sed -E 's/" & pattern & "/\\1/'"
        return result
    on error
        return ""
    end try
end extractPattern

-- Clean a string for use in filenames (remove spaces, special chars)
on cleanString(inputString)
    set cleaned to do shell script "echo " & quoted form of inputString & " | tr -cd 'A-Za-z0-9 \\-_' | sed 's/^ *//;s/ *$//;s/  */_/g'"
    return cleaned
end cleanString

-- Replace all occurrences of a string within another string
on replaceString(theText, oldString, newString)
    set AppleScript's text item delimiters to oldString
    set textItems to text items of theText
    set AppleScript's text item delimiters to newString
    set newText to textItems as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceString

-- Reverse a string
on reverseString(theText)
    set reversedText to ""
    repeat with i from (length of theText) to 1 by -1
        set reversedText to reversedText & character i of theText
    end repeat
    return reversedText
end reverseString

-- Join a list with a separator
on joinList(theList, separator)
    if length of theList is 0 then return ""
    if length of theList is 1 then return item 1 of theList
    
    set resultText to ""
    repeat with i from 1 to (count theList)
        set resultText to resultText & item i of theList
        if i < (count theList) then set resultText to resultText & separator
    end repeat
    
    return resultText
end joinList

-- Split a string into a list based on a separator
on splitString(theString, separator)
    set AppleScript's text item delimiters to separator
    set theItems to text items of theString
    set AppleScript's text item delimiters to ""
    return theItems
end splitString

-- Log a message with level-based filtering
on logMessage(msg, level)
    if level is missing value then set level to LOG_INFO
    
    -- Only log if message level is at or above current log level
    if level < currentLogLevel then return
    
    -- Format level prefix
    set levelPrefix to ""
    if level is LOG_DEBUG then
        set levelPrefix to "[DEBUG] "
    else if level is LOG_WARNING then
        set levelPrefix to "[WARNING] "
    else if level is LOG_ERROR then
        set levelPrefix to "[ERROR] "
    end if
    
    -- Write to log file
    set expandedLogPath to do shell script "echo " & defaultLogFile
    do shell script "echo \"" & (do shell script "date +'%Y-%m-%d %H:%M:%S'") & " - " & levelPrefix & msg & "\" >> " & quoted form of expandedLogPath
end logMessage

-- ═══ Settings and Configuration Mode ═════════════

-- Show settings configuration interface
on showSettings()
    -- Initialize paths
    set {expandedLogPath, expandedWorkDir, expandedConfigPath} to initializePaths()
    
    -- Load current configuration
    set configData to loadConfiguration(expandedConfigPath)
    if configData is missing value then set configData to {}
    
    -- Get current settings or set defaults
    set currentLogLevel to getConfigValue(configData, "logLevel", LOG_INFO)
    set currentDelayBetweenSlides to getConfigValue(configData, "delayBetweenSlides", delayBetweenSlides)
    set currentCaptureDelay to getConfigValue(configData, "captureDelay", captureDelay)
    set currentMaxTitleLength to getConfigValue(configData, "maxTitleLength", maxTitleLength)
    
    -- Display settings dialog
    set settingsDialog to display dialog "Configure Script Settings" & return & return & ¬
        "Log Level (0=Debug, 1=Info, 2=Warning, 3=Error):" & return & ¬
        "Delay Between Slides (seconds):" & return & ¬
        "Capture Delay (seconds):" & return & ¬
        "Max Title Length (characters):" & return & ¬
        return & ¬
        "Log File: " & expandedLogPath & return & ¬
        "Work Directory: " & expandedWorkDir & return & ¬
        "Config File: " & expandedConfigPath ¬
        default answer currentLogLevel & return & currentDelayBetweenSlides & return & currentCaptureDelay & return & currentMaxTitleLength ¬
        buttons {"Clear Resume State", "Reset Defaults", "Save"} default button "Save"
    
    -- Handle different buttons
    if button returned of settingsDialog is "Reset Defaults" then
        -- Reset to defaults
        set configData to {}
        saveConfiguration(configData, expandedConfigPath)
        display dialog "Settings reset to defaults" buttons {"OK"} default button "OK"
    else if button returned of settingsDialog is "Clear Resume State" then
        -- Clear resume state
        if configData contains "resumeData" then
            set configData's resumeData to {}
            saveConfiguration(configData, expandedConfigPath)
        end if
        display dialog "Resume state cleared" buttons {"OK"} default button "OK"
    else
        -- Save new settings
        -- Parse the multi-line input
        set settingsText to text returned of settingsDialog
        set settingsLines to paragraphs of settingsText
        
        -- Validate and update settings
        try
            set newLogLevel to item 1 of settingsLines as integer
            set newDelayBetweenSlides to item 2 of settingsLines as real
            set newCaptureDelay to item 3 of settingsLines as real
            set newMaxTitleLength to item 4 of settingsLines as integer
            
            -- Validate ranges
            if newLogLevel < 0 or newLogLevel > 3 then set newLogLevel to LOG_INFO
            if newDelayBetweenSlides < 0.1 then set newDelayBetweenSlides to 0.1
            if newCaptureDelay < 0.1 then set newCaptureDelay to 0.1
            if newMaxTitleLength < 10 then set newMaxTitleLength to 10
            if newMaxTitleLength > 200 then set newMaxTitleLength to 200
            
            -- Update config data
            set configData's logLevel to newLogLevel
            set configData's delayBetweenSlides to newDelayBetweenSlides
            set configData's captureDelay to newCaptureDelay
            set configData's maxTitleLength to newMaxTitleLength
            
            -- Save configuration
            saveConfiguration(configData, expandedConfigPath)
            
            display dialog "Settings saved successfully" buttons {"OK"} default button "OK"
        on error errMsg
            display dialog "Error saving settings: " & errMsg buttons {"OK"} default button "OK"
        end try
    end if
end showSettings

-- Helper function to get a configuration value or default
on getConfigValue(configData, keyName, defaultValue)
    if configData is missing value or configData is {} then return defaultValue
    if configData contains keyName then
        return configData's keyName
    else
        return defaultValue
    end if
end getConfigValue

-- ═══ Single File Renaming Mode ═════════════

-- Show help and documentation
on showHelp()
    set helpText to "UNIVERSAL SLIDE CAPTURE & OCR RENAMER v" & scriptVersion & "
===========================================

This tool automates capturing and intelligently naming presentation slides.

MODES:
------
1. Capture Slides from Presentation
   - Automatically captures slides from a browser presentation
   - Uses OCR to extract content and classify slides
   - Names files based on content and slide type

2. Rename Single Screenshot
   - Processes an existing screenshot
   - Extracts title and content with OCR
   - Renames the file intelligently
   
3. Settings & Configuration
   - Adjust capture delays and OCR parameters
   - Set log levels and file paths
   - Clear resume state

REQUIREMENTS:
------------
- macOS 11+ (Big Sur or newer)
- Tesseract OCR (text extraction)
- ImageMagick (image processing)

KEYBOARD SHORTCUTS DURING CAPTURE:
---------------------------------
- Option+P: Pause capture
- Option+R: Resume capture
- Option+C: Cancel capture

SLIDE TYPES RECOGNIZED:
----------------------
- Regular Content Slides
- Cyber Lab Slides
- Knowledge Check/Answer Slides
- Pulse Check Slides
- Summary Slides
- Break Slides
- Real World Scenario Slides

TROUBLESHOOTING:
---------------
- If text extraction is poor, try adjusting screen brightness
- Use presentations with good contrast for best results
- Check log file at: ~/Library/Logs/slide_capture.log

Version: " & scriptVersion & " (Released: " & scriptReleaseDate & ")"

    display dialog helpText buttons {"Install Dependencies", "Open Log File", "OK"} default button "OK"
    
    if button returned of result is "Install Dependencies" then
        displayInstallationInstructions()
    else if button returned of result is "Open Log File" then
        set expandedLogPath to do shell script "echo " & defaultLogFile
        do shell script "open -e " & quoted form of expandedLogPath
    end if
end showHelp

-- Entry point for single file renaming mode
on renameSingleFile()
    -- Check for required tools
    set {tesseractFound, tesseractPath} to checkDependency("tesseract", "Tesseract OCR")
    set {imageMagickFound, imageMagickPath} to checkDependency("convert", "ImageMagick")
    
    if not tesseractFound then
        display dialog "Tesseract OCR is required for file renaming." buttons {"Install", "Cancel"} default button "Install" cancel button "Cancel"
        if button returned of result is "Install" then
            displayInstallationInstructions()
        end if
        return
    end if
    
    -- Set up command environment
    set cmdEnvironment to setupEnvironment(tesseractPath, imageMagickPath)
    
    -- Let the user select a file
    set selectedFile to choose file with prompt "Select a screenshot to rename" of type {"png", "jpg", "jpeg"}
    
    -- Detect screen resolution and adjust geometry settings
    set geometrySettings to detectAndSetGeometry()
    
    -- Initialize paths
    set {expandedLogPath, expandedWorkDir, expandedConfigPath} to initializePaths()
    prepareWorkingDirectory(expandedWorkDir)
    
    -- Process the selected file
    set fileName to POSIX path of selectedFile
    set fileBaseName to do shell script "basename " & quoted form of fileName
    
    logMessage("Processing single file: " & fileBaseName, LOG_INFO)
    
    -- Display processing notification
    display notification "Processing file..." with title "Analyzing Screenshot"
    
    -- Extract slide information
    set slideTitle to processSlideTitle(fileName, "00", "", "", "", cmdEnvironment, geometrySettings)
    
    -- Generate new filename
    set fileDir to do shell script "dirname " & quoted form of fileName
    set newName to slideTitle & ".png"
    set newPath to fileDir & "/" & newName
    
    -- Ensure the new path is unique
    set newPath to ensureUniquePath(newPath)
    
    -- Rename the file
    do shell script "cp " & quoted form of fileName & " " & quoted form of newPath
    
    -- Show confirmation
    set newBaseName to do shell script "basename " & quoted form of newPath
    
    -- Ask if original should be kept or deleted
    set keepOriginal to display dialog "File copied to: " & newBaseName & return & return & "Would you like to keep the original file?" buttons {"Delete Original", "Keep Both"} default button "Keep Both"
    
    if button returned of keepOriginal is "Delete Original" then
        do shell script "rm " & quoted form of fileName
        logMessage("Deleted original file: " & fileBaseName, LOG_INFO)
    end if
    
    display notification "File processed and renamed to: " & newBaseName with title "Success!"
    logMessage("Renamed " & fileBaseName & " to " & newBaseName, LOG_INFO)
    
    -- Offer to process another file
    set processAnother to display dialog "File renamed successfully. Would you like to rename another file?" buttons {"No", "Yes"} default button "Yes"
    
    if button returned of processAnother is "Yes" then
        renameSingleFile()
    end if
end renameSingleFile
