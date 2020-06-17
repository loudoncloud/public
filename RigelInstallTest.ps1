############ utility functions ############################################################

function Resolve-Url {
<#
    .SYNOPSIS
        Recursively follow URL redirections until a non-redirecting URL is reached.

    .DESCRIPTION
        Chase URL redirections (e.g., FWLinks, safe links, URL-shortener links)
        until a non-redirection URL is found, or the redirection chain is deemed
        to be too long.

    .OUTPUT System.Uri
#>
param(
    [Parameter(Mandatory=$true)]
    [String]$url <# The URL to (recursively) resolve to a concrete target. #>
)
    $orig = $url
    $result = $null
    $depth = 0
    $maxdepth = 10

    do {
        if ($depth -ge $maxdepth) {
            Write-Error "Unable to resolve $orig after $maxdepth redirects."
        }
        $depth++
        $resolve = [Net.WebRequest]::Create($url)
        $resolve.Method = "HEAD"
        $resolve.AllowAutoRedirect = $false
        $result = $resolve.GetResponse()
        $url = $result.GetResponseHeader("Location")
    } while ($result.StatusCode -eq "Redirect")

    if ($result.StatusCode -ne "OK") {
        Write-Error ("Unable to resolve {0} due to status code {1}" -f $orig, $result.StatusCode)
    }

    return $result.ResponseUri
}

function Save-Url {
<#
    .SYNOPSIS
        Given a URL, download the target file to the same path as the currently-
        running script.

    .DESCRIPTION
        Download a file referenced by a URL, with some added niceties:

          - Tell the user the file is being downloaded
          - Skip the download if the file already exists
          - Keep track of partial downloads, and don't count them as "already
            downloaded" if they're interrupted

        Optionally, an output file name can be specified, and it will be used. If
        none is specified, then the file name is determined from the (fully
        resolved) URL that was provided.

    .OUTPUT string
#>
param(
    [Parameter(Mandatory=$true)]
    [String]$url, <# URL to download #>
    [Parameter(Mandatory=$true)]
    [String]$name, <# A friendly name describing what (functionally) is being downloaded; for the user. #>
    [Parameter(Mandatory=$false)]
    [String]$output = $null <# An optional file name to download the file as. Just a file name -- not a path! #>
)

    $res = (Resolve-Url $url)

    # If the filename is not specified, use the filename in the URL.
    if ([string]::IsNullOrEmpty($output)) {
        $output = (Split-Path $res.LocalPath -Leaf)
    }

    $File = Join-Path $PSScriptRoot $output
    if (!(Test-Path $File)) {
        Write-Host "Downloading $name... " -NoNewline
        $TmpFile = "${File}.downloading"

        # Clean up any existing (unfinished, previous) download.
        Remove-Item $TmpFile -Force -ErrorAction SilentlyContinue

        # Download to the temp file, then rename when the download is complete
        (New-Object System.Net.WebClient).DownloadFile($res, $TmpFile)
        Rename-Item $TmpFile $File -Force

        Write-Host "done"
    } else {
        Write-Host "Found $name already downloaded."
    }

    return $File
}

function Test-Signature {
<#
    .SYNOPSIS
        Verify the AuthentiCode signature of a file, deleting the file and writing
        an error if it fails verification.

    .DESCRIPTION
        Given a path, check that the target file has a valid AuthentiCode signature.
        If it does not, delete the file, and write an error to the error stream.
#>
param(
    [Parameter(Mandatory=$true)]
    [String]$Path <# The path of the file to verify the Authenticode signature of. #>
)
    if (!(Test-Path $Path)) {
        Write-Error ("File does not exist: {0}" -f $Path)
    }

    $name = (Get-Item $Path).Name
    Write-Host ("Validating signature for {0}... " -f $name) -NoNewline

    switch ((Get-AuthenticodeSignature $Path).Status) {
        ("Valid") {
            Write-Host "success."
        }

        default {
            Write-Host "failed."

            # Invalid files should not remain where they could do harm.
            Remove-Item $Path | Write-Debug
            Write-Error ("File {0} failed signature validation." -f $name)
        }
    }
}

function Remove-Directory {
  <#
    .SYNOPSIS
        
        Recursively remove a directory and all its children.

    .DESCRIPTION

        Powershell can't handle 260+ character paths, but robocopy can. This
        function allows us to safely remove a directory, even if the files
        inside exceed Powershell's usual 260 character limit.
  #>
param(
    [parameter(Mandatory=$true)]
    [string]$path <# The path to recursively remove #>
)

    # Make an empty reference directory
    $cleanup = Join-Path $PSScriptRoot "empty-temp"
    if (Test-Path $cleanup) {
        Remove-Item -Path $cleanup -Recurse -Force
    }
    New-Item -ItemType Directory $cleanup | Write-Debug

    # Use robocopy to clear out the guts of the victim path
    (Invoke-Native "& robocopy '$cleanup' '$path' /mir" $robocopy_success) | Write-Debug

    # Remove the folders, now that they're empty.
    Remove-Item $path -Force
    Remove-Item $cleanup -Force
}

function Invoke-Native {
<#
    .SYNOPSIS
        Run a native command and process its exit code.

    .DESCRIPTION
        Invoke a command line specified in $command, and check the resulting $LASTEXITCODE against
        $success to determine if the command succeeded or failed. If the command failed, error out.
#>
[CmdletBinding()]
param(
    [parameter(Mandatory=$true)]
    [string]$command, <# The native command to execute. #>
    [parameter(Mandatory=$false)]
    [ScriptBlock]$success = {$_ -eq 0} <# Test of $_ (last exit code) that returns $true if $command was successful, $false otherwise. #>
)

    Invoke-Expression $command
    $result = $LASTEXITCODE
    if (!($result |% $success)) {
        Write-Error "Command '$command' failed test '$success' with code '$result'."
        exit 1
    }
}

function Expand-Archive {
<#
    .SYNOPSIS
        Extract files from supported archives.

    .NOTES
        Supported file types are .msi and .cab.
#>
[CmdletBinding()]
param(
    [parameter(Mandatory=$true)]
    [string]$source, <# The archive file to expand. #>
    [parameter(Mandatory=$true)]
    [string]$destination <# The directory to place the extracted archive files in. #>
)

    if (!(Test-Path $destination)) {
        mkdir $destination | Write-Debug
    }

    switch ([IO.Path]::GetExtension($source)) {
        ".msi" {
            Start-Process "msiexec" -ArgumentList ('/a "{0}" /qn TARGETDIR="{1}"' -f $source, $destination) -NoNewWindow -Wait
        }
        ".cab" {
            (& expand.exe "$source" -F:* "$destination") | Write-Debug
        }
        default {
            Write-Error "Unsupported archive type."
            exit 1
        }
    }
}

############### script runtime ############################################################


### Acquire the SRS deployment kit

    $SRSDK = Save-Url "https://go.microsoft.com/fwlink/?linkid=851168" "deployment kit"
    Test-Signature $SRSDK


### Extract the deployment kit.

    $RigelMedia = 'C:\Users\Admin\Documents\Rigel'


    Write-Host "Extracting the deployment kit... " -NoNewline
    Expand-Archive $SRSDK $RigelMedia
    Write-Host "done." 

    $AppPath = $RigelMedia + '\Skype Room System Deployment Kit\$oem$\$1\Rigel\x64\Ship\AppPackages\*\*.appx'
    $DependencyPath = $RigelMedia + '\Skype Room System Deployment Kit\$oem$\$1\Rigel\x64\Ship\AppPackages\*\Dependencies\x64\*.appx'

### update the App package for Rigerl
    Write-Host "Updating the Rigel app... " -NoNewline 
    Add-AppxPackage -ForceApplicationShutdown -Path $AppPath -DependencyPath (Get-ChildItem $DependencyPath | Foreach-Object {$_.FullName})
    Write-Host "done."
### log the result

    $date = get-date
    $path = $Rigelmedia + "testlog.txt"
    "the script was run on " + (get-date) | out-file $path