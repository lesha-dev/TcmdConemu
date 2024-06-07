param(
    [Parameter()]                               [string]$WindowTitle = "Tcmd ConEmu",
    [Parameter(Mandatory)]                      [string]$ConEmuFolder,
    [Parameter(Mandatory)]                      [string]$SourcePath,
    [Parameter(Mandatory)] [AllowEmptyString()] [string]$SourceFocus,
    [Parameter(Mandatory)]                      [string]$TargetPath,
    [Parameter(Mandatory)] [AllowEmptyString()] [string]$TargetFocus,
    [Parameter(Mandatory)]                      [string]$LeftPath,
    [Parameter(Mandatory)] [AllowEmptyString()] [string]$LeftFocus,
    [Parameter(Mandatory)]                      [string]$RightPath,
    [Parameter(Mandatory)] [AllowEmptyString()] [string]$RightFocus,
    [Parameter()]                               [switch]$Execute
)

Write-Host "Arguments:"
Write-Host "Source path:" `"$SourcePath`"
Write-Host "Selected:   " `"$SourceFocus`"
Write-Host "Target path:" `"$TargetPath`"
Write-Host "Selected:   " `"$TargetFocus`"
Write-Host "Left path:  " `"$LeftPath`"
Write-Host "Selected:   " `"$LeftFocus`"
Write-Host "Right path: " `"$RightPath`"
Write-Host "Selected:   " `"$RightFocus`"
Write-Host "Window Title:" `"$WindowTitle`"
Write-Host "ConEmu folder:" `"$ConEmuFolder`"
Write-Host "Execute flag:" $Execute
Write-Host "---------------------------------------"

[string]$CONEMU = $ConemuFolder + "\ConEmu" + (&{If([Environment]::Is64BitOperatingSystem) {"64"} Else {""}}) + ".exe"
[string]$CONEMUC = $ConemuFolder + "\ConEmu\ConEmuC"  + (&{If([Environment]::Is64BitOperatingSystem) {"64"} Else {""}}) + ".exe"
[int]$RETRY_COUNT = 10

if (![System.IO.File]::Exists($CONEMU)) {
  Write-Host "Error: ConEmu executable doesn't exist `"$CONEMU`""
  Return
}

if (![System.IO.File]::Exists($CONEMUC)) {
  Write-Host "Error: ConEmu executable doesn't exist `"$CONEMUC`""
  Return
}


Write-Host "Find ConEmu window by title `"$WindowTitle`""

$extlib = @"
  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern IntPtr FindWindow(IntPtr sClassName, String sAppName);
"@
$win32  = Add-Type -Namespace Win32 -Name Funcs -MemberDefinition $extlib -PassThru
[int]$handle = [int]$win32::FindWindow([IntPtr]::Zero, $WindowTitle)
[string]$handleHex = ""

if ($handle -gt 0) {
  Write-Host "Window found with handle $handle, bring to front"
  Add-Type @"
  using System;
  using System.Runtime.InteropServices;

  public class SFW {
   [DllImport("user32.dll")]
   [return: MarshalAs(UnmanagedType.Bool)]
   public static extern bool SetForegroundWindow(IntPtr hWnd);
  }
"@
  [SFW]::SetForegroundWindow($handle)
  $handleHex = "x$($handle.ToString('X'))"
}
else {
  Write-Host "Window not found, run new console instance"
  & $CONEMU -max -title $WindowTitle -runlist cmd -new_console:d:`"$LeftPath`" '|||' cmd -new_console:s1THn -new_console:d:`"$RightPath`"

  Write-Host "Wait for window to appear..."
  [int]$count = 1
  while (($handle -eq 0) -and !($count -gt $RETRY_COUNT)) {
    Start-Sleep -Milliseconds (100*$count)
    Write-Host "Find ConEmu window by title `"$WindowTitle`", try $count"
    $handle = [int]$win32::FindWindow([IntPtr]::Zero, $WindowTitle)
    $handleHex = "x$($handle.ToString('X'))"
    $count = $count + 1
  }
  if ($count -gt $RETRY_COUNT) {
    Write-Host "Error: exceeded retry count $RETRY_COUNT to find ConEmu window"
    Return
  }
  Write-Host "Window found with handle $handle"

  Write-Host "Wait while all tabs are created..."
  [string]$status = ""
  $count = 1
  while (!($status -eq "Yes") -and !($status -eq "No") -and !($count -gt $RETRY_COUNT)) {
    Start-Sleep -Milliseconds (100*($count-1))
    Write-Host "Check second tab status, try $count"
    $status = & $CONEMUC "-GuiMacro:$($handleHex):T2" IsConsoleActive
    Write-Host "IsConsoleActive:" $status
    $count = $count + 1
  }
  if ($count -gt $RETRY_COUNT) {
    Write-Host "Error: exceeded retry count $RETRY_COUNT for IsConsoleActive command"
    Return
  }
}


[int]$selectedTab = 2
if (($SourcePath -eq $LeftPath) -and ($SourceFocus -eq $LeftFocus)) {
  $selectedTab = 1
}
Write-Host "Select tab $selectedTab according to tcmd active panel"
& c:\bin\ConEmu\ConEmu\ConEmuC64.exe "-GuiMacro:$($handleHex)" Tab 7 $selectedTab



Write-Host "Check if there is no running subprocess in current tab"
[string]$activePID = & c:\bin\ConEmu\ConEmu\ConEmuC64.exe "-GuiMacro:$($handleHex)" GetInfo ActivePID
Write-Host "ActivePID:" $activePID
[string]$rootXML = & c:\bin\ConEmu\ConEmu\ConEmuC64.exe "-GuiMacro:$($handleHex)" GetInfo Root
Write-Host "RootXML:" $rootXML
[string]$rootPID = ""

if ($rootXML -match 'PID="(.*?)"') {
  $rootPID = $Matches.1
  Write-Host "RootPID:" $rootPID
}
else {
  Write-Host "Error: Can't retrieve PID from info"
}

if (($rootPID -eq $activePID) -or ($rootPID -eq "")) {
  Write-Host "Changing directory"
  [string]$CurDir = & c:\bin\ConEmu\ConEmu\ConEmuC64.exe "-GuiMacro:x$($handle.ToString('X'))" GetInfo CurDir
  if ($CurDir[$CurDir.Length-1] -eq "\") {
    $CurDir = $CurDir.Substring(0, $CurDir.Length - 1)
  }
  [string]$SourcePathTrimSlash = $SourcePath.Substring(0, $SourcePath.Length - 1) 
  Write-Host "ConEmu directory:" $CurDir
  Write-Host "Tcmd directory:  " $SourcePathTrimSlash

  if (!($CurDir -eq $SourcePathTrimSlash)) {
    [string]$cdCommand = "cd /D " + $SourcePathTrimSlash.Replace('\', '\\')
    Write-Host "Send command $cdCommand"
    & c:\bin\ConEmu\ConEmu\ConEmuC64.exe "-GuiMacro:$($handleHex)" print $cdCommand"`n"
  } else {
    Write-Host "Directories equals, skip"
  }

  if ($Execute.IsPresent -and ($SourceFocus.Length -gt 0)) {
    Write-Host "Execute script from tcmd"

    [string]$executeCommand = $SourceFocus
    if ($executeCommand -contains " ") {
      $executeCommand = "`"$($executeCommand)`""
    }
    Write-Host "Send command $executeCommand"
    & c:\bin\ConEmu\ConEmu\ConEmuC64.exe "-GuiMacro:$($handleHex)" print $executeCommand"`n"
  }
} else {
  Write-Host "Skip sending commands, some subprocess is running"
}
