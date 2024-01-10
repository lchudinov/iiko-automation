Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class User32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
"@

$iikoWindowId = ((Get-Process).where{$_.MainWindowTitle -eq 'Кондитория (Лаборатория вкуса ЕКБ (УФА), Лаборатория Вкуса ЕКБ, Лаборатория Вкуса ЕКБ (ТМН))  - iikoChain 2022'})[0].Id
Write-Output $iikoWindowId

$root = [Windows.Automation.AutomationElement]::RootElement
$condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $iikoWindowId)
$iikoUI = $root.FindFirst([Windows.Automation.TreeScope]::Children, $condition)

Write-Output $iikoUI

# $hwnd = $iikoUI.Current.NativeWindowHandle
# [User32]::SetForegroundWindow($hwnd)

$window = $iikoUI.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
$window.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
Start-Sleep -s 2


# function FindPanel() {
# 	Start-Sleep -s 2
# 	$condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, "Новая расходная накладная")
# 	$panel = $iikoUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition)
# 	return $panel
# }


# $panel = FindPanel
# if ($null -eq $panel) {
# 	Write-Host "panel not found"
# 	exit
# }
# Write-Host "panel is" $panel.Current.Name $panel.Current.ClassName

function FindMainPropertiesPanel() {
	Start-Sleep -s 2
	# $condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ClassNameProperty, "WindowsForms10.Window.8.app.0.35e60a0_r6_ad1")
	# $condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "dateEditIncomingDateTime")
	$condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, "Основные свойства")
	# $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "gridInvoice")
	# $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "gridInvoice")
	# $condition = New-Object Windows.Automation.AndCondition($condition1, $condition2)
	$mainPropertiesPanel = $iikoUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition1)
	# $gridPattern = $grid.GetCurrentPattern([Windows.Automation.GridPattern]::Pattern)
	if ($null -ne $mainPropertiesPanel) {
		Write-Host "main properties panel found" $grid
	} else {
		Write-Host "main properties panel not found"
		exit
	}
	return $mainPropertiesPanel
	# $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
	# $grid.GetCurrentPattern([Windows.Automation.GridPattern]::Pattern).Invoke()
}

$mainPropertiesPanel = FindMainPropertiesPanel $panel
Write-Host "mainPropertiesPanel is $($mainPropertiesPanel.Current.Name) $($mainPropertiesPanel.Current.ClassName)"

if ($null -eq $mainPropertiesPanel) {
	exit
}
