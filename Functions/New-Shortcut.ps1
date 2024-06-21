function New-Shortcut {
   <#
      .SYNOPSIS
      Creates a new shortcut with specified properties.

      .DESCRIPTION
      This function creates a new shortcut at the specified path, setting its target path,
      icon, and additional arguments if provided. It uses the WScript.Shell COM object to 
      handle the creation and configuration of the shortcut.

      .PARAMETER DestinationShortcutPath
      The file path where the shortcut will be created. Must end with '.lnk'.

      .PARAMETER TargetShortcutPath
      The target path the shortcut points to.

      .PARAMETER Arguments
      (Optional) Additional arguments to pass to the target application.

      .PARAMETER SourceIconPath
      (Optional) The path to the icon file to be copied and used for the shortcut. Must end with '.ico'.
      Both SourceIconPath and DestinationIconPath must be provided together.

      .PARAMETER DestinationIconPath
      (Optional) The destination path where the icon file will be copied. Must end with '.ico'.
      Both SourceIconPath and DestinationIconPath must be provided together.

      .EXAMPLE
      New-Shortcut -DestinationShortcutPath "C:\Users\Public\Desktop\MyApp.lnk" -TargetShortcutPath "C:\Program Files\MyApp\MyApp.exe"

      .EXAMPLE
      New-Shortcut -DestinationShortcutPath "C:\Users\Public\Desktop\MyApp.lnk" -TargetShortcutPath "C:\Program Files\MyApp\MyApp.exe" -SourceIconPath "C:\Icons\MyIcon.ico" -DestinationIconPath "C:\Users\Public\Desktop\MyIcon.ico"

      .EXAMPLE
      New-Shortcut -DestinationShortcutPath "C:\Users\Public\Desktop\MyApp.lnk" -TargetShortcutPath "C:\Program Files\MyApp\MyApp.exe" -Arguments "/start"
   #>
   [CmdletBinding()]
   param (
      [parameter(Mandatory = $true)]
      [ValidateScript({
         if ($_ -notmatch "\.lnk$") {
               throw "DestinationShortcutPath must end with '.lnk'."
         }
         $true
      })]
      [string]$DestinationShortcutPath,

      [parameter(Mandatory = $true)]
      [string]$TargetShortcutPath,

      [parameter(Mandatory = $false)]
      [string]$Arguments,

      [parameter(Mandatory = $false)]
      [ValidateScript({
         if ($_ -notmatch "\.ico$") {
               throw "SourceIconPath must end with '.ico'."
         }
         $true
      })]
      [string]$SourceIconPath,

      [parameter(Mandatory = $false)]
      [ValidateScript({
         if ($_ -notmatch "\.ico$") {
               throw "DestinationIconPath must end with '.ico'."
         }
         $true
      })]
      [string]$DestinationIconPath
   )

   begin {
      # Create a WScript.Shell object
      $WScriptShell = New-Object -ComObject WScript.Shell

      # Validation for SourceIconPath and DestinationIconPath
      if (($SourceIconPath -and -not $DestinationIconPath) -or (-not $SourceIconPath -and $DestinationIconPath)) {
         throw "Both SourceIconPath and DestinationIconPath must be provided together."
      }
   }
   
   process {
      # Copy Source Icon to Destination Icon Path
      if ($DestinationIconPath -and $SourceIconPath) {
         Copy-Item -Path $SourceIconPath -Destination $DestinationIconPath
      }
      # Create the shortcut
      $Shortcut = $WScriptShell.CreateShortcut($DestinationShortcutPath)
      # Set the target path
      $Shortcut.TargetPath = $TargetShortcutPath
      # Set the additional parameters for the shortcut if they are provided
      if ($Arguments) {
         $Shortcut.Arguments = $Arguments
      }
      # set the icon
      if ($DestinationIconPath) {
         $Shortcut.IconLocation = $DestinationIconPath
      }
      # Save the shortcut
      $Shortcut.Save()
   }
}
