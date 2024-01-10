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

# $calc = [Diagnostics.Process]::Start('calc')
#wait for the UI to appear
# $null = $calc.WaitForInputIdle(5000)
# Start-Sleep -s 2
$iikoWindowId = ((Get-Process).where{$_.MainWindowTitle -eq 'Кондитория (Лаборатория вкуса ЕКБ (УФА), Лаборатория Вкуса ЕКБ, Лаборатория Вкуса ЕКБ (ТМН))  - iikoChain 2022'})[0].Id
Write-Output $iikoWindowId

$root = [Windows.Automation.AutomationElement]::RootElement
$condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, $iikoWindowId)
$iikoUI = $root.FindFirst([Windows.Automation.TreeScope]::Children, $condition)

Write-Output $iikoUI

# $hwnd = $iikoUI.Current.NativeWindowHandle
# [User32]::SetForegroundWindow($hwnd)

# $window = $iikoUI.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
# $window.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
# Start-Sleep -s 2


function FindPanel() {
	Start-Sleep -s 2
	$condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, "Новая расходная накладная")
	$panel = $iikoUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition)
	return $panel
}


$panel = FindPanel
if ($null -eq $panel) {
	Write-Host "panel not found"
	exit
}
Write-Host "panel is" $panel.Current.Name $panel.Current.ClassName

function FindDateField($panel){
	Start-Sleep -s 2
	# $condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ClassNameProperty, "WindowsForms10.Window.8.app.0.35e60a0_r6_ad1")
	# $condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "labelDateIncoming")
	$condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, "Дата и время выдачи товара:")
	# $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "gridInvoice")
	# $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "gridInvoice")
	# $condition = New-Object Windows.Automation.AndCondition($condition1, $condition2)
	$dateField = $iikoUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition1)
	# $gridPattern = $grid.GetCurrentPattern([Windows.Automation.GridPattern]::Pattern)
	if ($null -ne $dateField) {
		Write-Host "date field found" $grid
	} else {
		Write-Host "date field not found"
		exit
	}
	return $dateField
	# $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
	# $grid.GetCurrentPattern([Windows.Automation.GridPattern]::Pattern).Invoke()
}

function FindCell(){
	Start-Sleep -s 2
	$condition1 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, "Панель данных")
	# $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, "таблицу")
	# $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "gridInvoice")
	# $condition2 = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "gridInvoice")
	# $condition = New-Object Windows.Automation.AndCondition($condition1, $condition2)
	$cell = $grid.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition1)
	if ($null -ne $cell) {
		Write-Host "cell found" $cell
	} else {
		Write-Host "cell not found"
		exit
	}
	return $cell
	# $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
	# $grid.GetCurrentPattern([Windows.Automation.GridPattern]::Pattern).Invoke()
}

$dateField = FindDateField $panel

Write-Host "dateField is $($dateField.Current.Name) $($dateField.Current.ClassName)"
exit
$cell = FindCell


#get and click the buttons for the calculation

FindAndClickButton Пять
FindAndClickButton Плюс
FindAndClickButton Девять
FindAndClickButton Равно

#get the result
$condition = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::AutomationIdProperty, "CalculatorResults")
$result = $iikoUI.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition)
$result.current.name