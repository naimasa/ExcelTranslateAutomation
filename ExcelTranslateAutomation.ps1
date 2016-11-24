################################################################################
# Excel Translate Automation
#
# Copyright (C) 2016 Masaki Naito
# Released under the MIT license
# http://opensource.org/licenses/mit-license.php
#
# Prerequests:
#   - Microsoft Windows 7
#   - Micorsoft Excel 2010
#   - Windows PowerShell
#   - UI Automation PowerShell Extensions (https://uiautomation.codeplex.com/)
#
# Usage (on PowerShell prompt):
#   ExcelTranslate.ps1 {cnt}
#   {cnt}: Count for repetition
#   * Before running script, activate 'çZâ{' tab and cell to be translated on the Excel window.
#
################################################################################

$ErrorActionPreference = "stop"
[UIAutomation.Preferences]::Highlight = $false

ipmo C:\System\UIAutomation\UIAutomation.dll

$cnt = $args[0]

$form = Get-UiaWindow -Class 'XLMAIN'
$wshell = New-Object -ComObject wscript.shell;

1..$cnt | foreach {

    # Translate active cell.
    $transBtn = $form | Get-UiaCustom -Class 'NetUIOrderedGroup' -Name 'çZâ{' | Get-UiaGroup -Class 'NetUIChunk' -Name 'åæåÍ' | Get-UiaButton -Class 'NetUIRibbonButton' -Name 'ñ|ñÛ'
    $transBtn | Invoke-UiaButtonClick
    
    # Re-activate Excel window to update translation pain.
    Get-UiaWindow -Class 'ConsoleWindowClass' -Name 'Administrator: Windows PowerShell (x86)' | Set-UiaFocus
    Sleep 1
    Get-UiaWindow -Class 'XLMAIN' | Set-UiaFocus
    
    # Wait for few seconds until getting result of online translation.
    Sleep 3
    
    # Replace cell value with translated one.
    $insBtn = $form | Get-UiaGroup -Class 'NetUIGroupBox' -Name 'ñ|ñÛ' | Get-UiaButton -Class 'NetUISplitDropdown' -Name 'ë}ì¸(I)' | Get-UiaButton -Class 'NetUISplitDropdownButton' -Name 'ë}ì¸(I)'
    $insBtn | Invoke-UiaButtonClick

    # Move to next cell.
    Get-UiaPane -Class 'EXCEL7' | Set-UiaFocus
    $wshell.SendKeys("{ENTER}")

    Sleep 1
}

