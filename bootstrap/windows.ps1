#
# Windows Let's Encrypt Bootstrapping
# Felix Rieseberg, felix.rieseberg@microsoft.com
#

" 

  _          _   _       _____                            _   
 | |        | | ( )     |  ___|                          | |  
 | |     ___| |_|/ ___  | |__ _ __   ___ _ __ _   _ _ __ | |_ 
 | |    / _ \ __| / __| |  __| '_ \ / __| '__| | | | '_ \| __|
 | |___|  __/ |_  \__ \ | |__| | | | (__| |  | |_| | |_) | |_ 
 \_____/\___|\__| |___/ \____/_| |_|\___|_|   \__,_| .__/ \__|
                                                   | |        
                                                   |_|        

Executing Bootstrapper for Windows...
"

#
# Variables
#

$is64bit = [Environment]::Is64BitOperatingSystem
$OSversion = [Environment]::OSVersion.Version

#
# Functions
#

function IsPythonInstalled
{
    if ((Get-Command "python.exe" -ErrorAction SilentlyContinue) -and (Test-Path "C:\Python27"))
    { 
        return "true"
    } else {
        return "false"
    }
}

function DownloadPython
{
    if ($is64bit) 
    {
        "$(Get-Date -format t): Downloading Python 2.7.9 (x64)"
        $source = "https://www.python.org/ftp/python/2.7.9/python-2.7.9rc1.msi"
    } else {
        "$(Get-Date -format t): Downloading Python 2.7.9 (x86)"
        $source = "https://www.python.org/ftp/python/2.7.9/python-2.7.9rc1.amd64.msi"
    }
    
    $destination = "./python-2.7.le-bootstrapper.msi"
    Invoke-WebRequest $source -OutFile $destination
}

function RunPythonInstaller
{
    "$(Get-Date -format t): Starting Python Installer..."
    return (Start-Process -FilePath "msiexec.exe" -ArgumentList "/i python-2.7.le-bootstrapper.msi /qn" -Wait -Passthru).ExitCode
}

function AddPythonToPath
{
    [Environment]::SetEnvironmentVariable("Path", "$env:Path;C:\Python27\;C:\Python27\Scripts\", "User")
}

function InstallVirtualEnv
{
    if (Get-Command "C:\Python27\Scripts\pip") 
    {
        "$(Get-Date -format t): Installing virtualenv..."
        C:\Python27\Scripts\pip.exe install virtualenv *> $null
        "$(Get-Date -format t): Installing virtualenvwrapper-powershell..."
        C:\Python27\Scripts\pip.exe install virtualenvwrapper-powershell *> $null
        "$(Get-Date -format t): Creating virtualenv directory..."
        mkdir '~\.virtualenvs' *> $null
        "$(Get-Date -format t): Activating virtualenvwrapper-powershell module..."
        Import-Module virtualenvwrapper
    } else 
    {
        throw "$(Get-Date -format t): Could not run pip. Please install Python 2.7 manually by visiting www.python.org. Aborting."
    }
}

function InstallVCPython {
    "$(Get-Date -format t): Installing VC Compiler for Python 2.7..."
    "Insallation of the VC Compiler has to be manual. Please visit http://aka.ms/vcpython27"
    "to download and install the compiler. Once you're done, please continue this bootstrapper"
    "by pressing any key."
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} 

function IsAdministrator
{
    $Identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object System.Security.Principal.WindowsPrincipal($Identity)
    $Principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}


function IsUacEnabled
{
    (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System).EnableLua -ne 0
}

function SetupLetsEncryptVenv
{
    "$(Get-Date -format t): Creating virtualenv directory..."
    cd ..
    virtualenv -p C:\python27\python.exe venv
}

function InstallSwig
{
    "$(Get-Date -format t): Installing Swig..."
    "Insallation of the Swig has to be manual. Please visit http://aka.ms/swig3"
    "to download Swig. Unzip SWIG in a folder of your choice and enter it here:"
    "ZIP."
    $swigZiplocation = Read-Host "Please enter the path to the folder containing swig.exe:"

    "$(Get-Date -format t): Adding Swig to path..."
    [Environment]::SetEnvironmentVariable("Path", $env:Path + ";" + $swigZiplocation, [System.EnvironmentVariableTarget]::Machine)

    "$(Get-Date -format t): Reloading path for current session"
    # Reload the PATH environment variable for the current session
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")

}

function SetupSwigEnvVars
{
    $env:PYTHON_INCLUDE = "C:\Python27\include"
    $env:PYTHON_LIB = "C:\Python27\libs\python27.lib"
}

# ----------------------------------------------------------------------------------------------------------
# Bootstrapper ---------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------

# Ensure that we're on Windows 8.1 or greater
if (($OSversion.Major -eq 6 -AND $OSversion.Minor -gt 1) -or ($OSversion.Major -gt 6)) 
{
    "$(Get-Date -format t): Current Windows version is supported."
} else {
    "$(Get-Date -format t): This bootstrapper is for Windows 8.1 or greater only."
    exit
}

# Self-Elevate
if (!(IsAdministrator))
{
    if (IsUacEnabled)
    {
        [string[]]$argList = @('-NoProfile', '-NoExit', '-File', $MyInvocation.MyCommand.Path)
        $argList += $MyInvocation.BoundParameters.GetEnumerator() | Foreach {"-$($_.Key)", "$($_.Value)"}
        $argList += $MyInvocation.UnboundArguments
        Start-Process PowerShell.exe -Verb Runas -WorkingDirectory $pwd -ArgumentList $argList
        return
    }
    else
    {
        throw "$(Get-Date -format t): You must be administrator to run this script. Aborting."
    }
}

if (-Not (IsPythonInstalled))
{
    $downloadPython = Read-Host "$(Get-Date -format t): We checked your system and could not find Python 2.7. Download and Install now?"
    if ($downloadPython) 
    {
        DownloadPython
        if (RunPythonInstaller -eq 0) 
        {
            "$(Get-Date -format t): Python successfully installed."
            AddPythonToPath
        } else {
            throw "$(Get-Date -format t): Python 2.7 Installation failed. Please install it manually by visiting www.python.org. Aborting."
        }
    } else {
        throw "$(Get-Date -format t): We need Python to run this script. Aborting."
    }
} else 
{
    "$(Get-Date -format t): Python 2.7 is installed. We only took a look, please install it if you're surprised."
}

InstallVirtualEnv
SetupLetsEncryptVenv
InstallVCPython
InstallSwig
SetupSwigEnvVars