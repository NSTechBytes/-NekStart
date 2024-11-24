param (
    [string[]]$skinNames 
)

# Check if at least one skin name is provided
if (-not $skinNames -or $skinNames.Count -eq 0) {
    Write-Host "Error: At least one skin name must be provided. Use -skinNames parameter to specify them."
    exit
}

# Function to get Rainmeter skin folder path
function Get-RainmeterSkinsFolder {
    $defaultProgramFilesPath = "$env:ProgramFiles\Rainmeter\Skins"
    $defaultProgramFilesX86Path = "$env:ProgramFiles(x86)\Rainmeter\Skins"
    $defaultDocumentsPath = "$env:UserProfile\Documents\Rainmeter\Skins"

    if (Test-Path $defaultProgramFilesPath) { return $defaultProgramFilesPath }
    if (Test-Path $defaultProgramFilesX86Path) { return $defaultProgramFilesX86Path }
    if (Test-Path $defaultDocumentsPath) { return $defaultDocumentsPath }

    $rainmeterConfigPath = "$env:AppData\Rainmeter\Rainmeter.ini"
    if (Test-Path $rainmeterConfigPath) {
        $config = Get-Content $rainmeterConfigPath
        foreach ($line in $config) {
            if ($line -like "SkinPath=*") {
                $customSkinPath = $line -replace "SkinPath=", ""
                if (Test-Path $customSkinPath) {
                    return $customSkinPath
                }
            }
        }
    }
    return "Rainmeter Skins folder not found!"
}

# Function to stop Rainmeter
function Stop-Rainmeter {
    $rainmeterProcess = Get-Process -Name "Rainmeter" -ErrorAction SilentlyContinue
    if ($rainmeterProcess) {
        Write-Host "Rainmeter is running. Stopping it..."
        Stop-Process -Name "Rainmeter" -Force
        Write-Host "Rainmeter has been stopped."
    } else {
        Write-Host "Rainmeter is not running."
    }
}

# Stop Rainmeter
Stop-Rainmeter

# Common variables
$destinationFolder = "C:\nstechbytes"
$skinsDirectory = Get-RainmeterSkinsFolder
$is64Bit = [Environment]::Is64BitOperatingSystem
$rainmeterPluginsPath = Join-Path -Path $env:UserProfile -ChildPath "AppData\Roaming\Rainmeter\Plugins"

# Process each skin
foreach ($skinName in $skinNames) {
    Write-Host "Processing skin: $skinName"
    

    # Download Version.nek
    $versionUrl = "https://raw.githubusercontent.com/NSTechBytes/$skinName/main/%40Resources/Version.nek"
    $tempVersionFile = "$env:TEMP\$skinName-Version.nek"
    Invoke-WebRequest -Uri $versionUrl -OutFile $tempVersionFile

    # Extract version
    $versionData = Get-Content $tempVersionFile
    $versionLine = $versionData | Where-Object { $_ -match "Version=" }
    $version = $versionLine -replace "Version=", ""

    # Download skin
    $url = "https://github.com/NSTechBytes/$skinName/releases/download/v$version/${skinName}_$version.rmskin"
    $destination = "$destinationFolder\${skinName}_$version.rmskin"

    if (-Not (Test-Path -Path $destinationFolder)) {
        New-Item -ItemType Directory -Path $destinationFolder
    }

    Invoke-WebRequest -Uri $url -OutFile $destination
    Write-Host "Downloaded $skinName version $version to $destination"

    # Extract .rmskin file
    $extractionFolder = "$destinationFolder\$skinName"
    if (-Not (Test-Path -Path $extractionFolder)) {
        New-Item -ItemType Directory -Path $extractionFolder
    }

    Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
    [System.IO.Compression.ZipFile]::ExtractToDirectory($destination, $extractionFolder)
    Write-Host "Extracted $skinName to $extractionFolder"

    # Clean up temporary version file
    Remove-Item $tempVersionFile

    # Remove existing skin folder
    if (Test-Path -Path $skinsDirectory) {
        $skinToRemove = Join-Path -Path $skinsDirectory -ChildPath $skinName
        if (Test-Path -Path $skinToRemove) {
            Remove-Item -Path $skinToRemove -Recurse -Force
            Write-Host "Removed the existing skin folder: $skinToRemove"
        } else {
            Write-Host "No existing skin folder found for: $skinName"
        }
    }

    # Copy new skin folder
    $extractedSkinFolder = Join-Path -Path $extractionFolder -ChildPath "Skins\$skinName"
    if (Test-Path -Path $extractedSkinFolder) {
        $destinationSkinPath = Join-Path -Path $skinsDirectory -ChildPath $skinName
        Copy-Item -Path $extractedSkinFolder -Destination $destinationSkinPath -Recurse -Force
        Write-Host "Copied the extracted folder to: $destinationSkinPath"
    } else {
        Write-Host "Extracted skin folder does not exist: $extractedSkinFolder"
    }

    # Copy plugins
    $pluginsFolder = Join-Path -Path $extractionFolder -ChildPath "Plugins\64bit"
    if (-Not $is64Bit) {
        $pluginsFolder = Join-Path -Path $extractionFolder -ChildPath "Plugins\32bit"
    }

    if (Test-Path -Path $pluginsFolder) {
        if (Test-Path -Path $rainmeterPluginsPath) {
            Get-ChildItem -Path $pluginsFolder -Filter "*.dll" | ForEach-Object {
                Copy-Item -Path $_.FullName -Destination $rainmeterPluginsPath -Force
                Write-Host "Copied DLL: $($_.Name) to $rainmeterPluginsPath"
            }
        } else {
            Write-Host "Rainmeter Plugins folder not found!"
        }
    } else {
        Write-Host "No DLL files found in: $pluginsFolder"
    }
}

# Start Rainmeter
$rainmeterPath = Join-Path -Path $env:ProgramFiles -ChildPath "Rainmeter\Rainmeter.exe"
if (-Not (Test-Path -Path $rainmeterPath)) {
    $rainmeterPath = Join-Path -Path $env:ProgramFiles(x86) -ChildPath "Rainmeter\Rainmeter.exe"
}

if (Test-Path -Path $rainmeterPath) {
    Start-Process -FilePath $rainmeterPath
    Write-Host "Rainmeter started."
} else {
    Write-Host "Rainmeter executable not found!"
}

# Remove the nstechbytes folder
if (Test-Path -Path $destinationFolder) {
    Remove-Item -Path $destinationFolder -Recurse -Force
    Write-Host "Removed the nstechbytes folder: $destinationFolder"
} else {
    Write-Host "nstechbytes folder not found!"
}

Write-Host "Update process completed!"
