# Importa la biblioteca de UIAutomation
Add-Type -AssemblyName UIAutomationClient, PresentationCore, PresentationFramework

# Función para simular un clic en un botón
function Click-Button {
    param(
        [System.Windows.Automation.AutomationElement]$element
    )
    $invokePattern = [System.Windows.Automation.InvokePattern]::Identify($element)
    $invokePattern.Invoke()
}

# Abre "Propiedades de Internet"
Start-Process "inetcpl.cpl"

# Espera a que la ventana "Propiedades de Internet" aparezca
$internetPropertiesWindow = $null
while ($internetPropertiesWindow -eq $null) {
    Start-Sleep -Milliseconds 500
    $internetPropertiesWindow = Get-Process | Where-Object { $_.MainWindowTitle -eq "Propiedades de Internet" }
}

# Obtiene la ventana como un elemento de automatización
$automationElement = [System.Windows.Automation.AutomationElement]::FromHandle($internetPropertiesWindow.MainWindowHandle)

# Selecciona la pestaña "Seguridad"
$tabSecurity = $automationElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "Seguridad")))
$tabSecurity.Select()

# Haz clic en "Nivel personalizado"
$buttonCustomLevel = $automationElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "Nivel personalizado…")))
Click-Button $buttonCustomLevel

# Encuentra y marca "Ejecutar aplicaciones y archivos no seguros"
$checkBox = $automationElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "Ejecutar aplicaciones y archivos no seguros")))
$checkBox.Select()

# Haz clic en "Aceptar"
$buttonOK = $automationElement.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "Aceptar")))
Click-Button $buttonOK

# Cierra la ventana de "Propiedades de Internet"
$internetPropertiesWindow.CloseMainWindow()

Write-Host "Proceso completado. La advertencia de seguridad al abrir archivos ha sido desactivada."

# Termina el proceso de PowerShell
exit
