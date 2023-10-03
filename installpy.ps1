# Install python batch file
# Set the version and download URL for Python
$version = "3.9.5"
$url = "https://www.python.org/ftp/python/$version/python-$version-amd64.exe"

# Download and install Python
$installPath = "$($env:ProgramFiles)\Python$version"
Invoke-WebRequest $url -OutFile python-$version.exe
Start-Process python-$version.exe -ArgumentList "/quiet", "TargetDir=$installPath" -Wait

# Add Python to the system PATH
$envVariable = [Environment]::GetEnvironmentVariable("Path", "Machine")
if ($envVariable -notlike "*$installPath*") {
    [Environment]::SetEnvironmentVariable("Path", "$envVariable;$installPath", "Machine")
    Write-Host "Added Python to PATH."
}

# Install packages in requirements.txt
pip install -r requirements.txt


# Clean up
Remove-Item python-$version.exe