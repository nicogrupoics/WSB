Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "Realizado por Nico Elvira" -ForegroundColor Yellow
Write-Host "--------------------------" -ForegroundColor Yellow
Write-Host ""

# Ruta de la tarea
$TaskPath = "\Microsoft\Windows\"  # Ruta de la carpeta donde se encuentra la tarea
$TaskName = "WSB"                 # Nombre de la tarea
$NewTime = "08:00"                # Nueva hora de ejecución (8:00 AM)

try {
    # Obtener la tarea programada
    $Task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName

    if ($Task) {
        # Crear un nuevo disparador para la tarea
        $Trigger = New-ScheduledTaskTrigger -At $NewTime -Once

        # Actualizar la tarea programada con el nuevo disparador
        Register-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -Trigger $Trigger -Action $Task.Actions -Force
    }
} catch {
    # Manejo del error
    $ErrorMessage = "Error al modificar la tarea programada: $($_.Exception.Message)"
}



# Asegurar que la variable COMPANY_NAME esté configurada
if (-not $env:COMPANY_NAME) {
    [Environment]::SetEnvironmentVariable("COMPANY_NAME", $env:COMPUTERNAME, "Machine")
}

# Obtener nombre del equipo y de la empresa
$CompanyName = $env:COMPANY_NAME
$ComputerName = $env:COMPUTERNAME
$Date = (get-date).AddHours(-48)
$ToEmail = "revisiones@grupoicsolutions.com"  # Dirección de correo del destinatario principal
$BccEmail = "nico.elvira@grupoicsolutions.com"  # Dirección de correo para BCC (copia oculta)

$footerMessage = "`nNo responder al correo, sistema automático."  # Mensaje de pie de página

try {
    $Summary = Get-WBSummary
    $CurrentJob = Get-WBJob
} catch {
    $errorMsg = "No se pudo ejecutar el comando de PowerShell para obtener el estado de la copia de seguridad: $($_.Exception.Message)"
    $Body = $errorMsg + $footerMessage

    # Crear el mensaje de correo electrónico con prioridad alta
    $Message = New-Object Net.Mail.MailMessage
    $Message.From = "revisiones@grupoicsolutions.com"
    $Message.To.Add($ToEmail)
    $Message.Bcc.Add($BccEmail)
    $Message.Subject = "$$CompanyName Backup - Error"
    $Message.Body = $Body
    $Message.Priority = "High"

    # Enviar el correo electrónico
    $SMTPClient = New-Object Net.Mail.SmtpClient("smtp.grupoicsolutions.com", 25)
    $SMTPClient.Send($Message)
    exit 1
}

$FailedBackups = invoke-command {
    if ($CurrentJob.JobState -eq 'Running' -and $CurrentJob.StartTime -lt (get-date).AddHours(-23)) { "`nADVERTENCIA`nCopia o recuperación en curso. Comenzó a las $($CurrentJob.StartTime)" }
    if ($Summary.LastBackupResultHR -ne '0') { "La copia de seguridad se completó con el código de error $($Summary.LastBackupResultHR)." }
    if ($Summary.LastSuccessfulBackupTime -lt $date) { "No se ha realizado una copia de seguridad exitosa en las últimas 48 horas." }
}

# Filtrar y eliminar las propiedades específicas del resumen
$filteredSummary = $Summary | Select-Object -Property * -ExcludeProperty LastSuccessfulBackupTime, LastBackupTime, LastSuccessfulBackupTargetPath, LastSuccessfulBackupTargetLabel

if ($FailedBackups) {
    $errorMsg = "ERROR. Por favor, revisa la información de diagnóstico."
    $Body = $errorMsg + "`n" + ($FailedBackups, $filteredSummary | Out-String) + $footerMessage

    # Crear el mensaje de correo electrónico con prioridad alta
    $Message = New-Object Net.Mail.MailMessage
    $Message.From = "revisiones@grupoicsolutions.com"
    $Message.To.Add($ToEmail)
    $Message.Bcc.Add($BccEmail)
    $Message.Subject = "$CompanyName Backup - Fallida"
    $Message.Body = $Body
    $Message.Priority = "High"

    # Enviar el correo electrónico
    $SMTPClient = New-Object Net.Mail.SmtpClient("smtp.grupoicsolutions.com", 25)
    $SMTPClient.Send($Message)
    exit 1
} else {
    $successMsg = "CORRECTA. No se encontraron copias de seguridad fallidas"
    $Body = $successMsg + "`n" + ($FailedBackups, $filteredSummary | Out-String)

    # Integrar la recuperación de archivo
    $FileToRecover = "C:\PRUEBA RECUPERACION\PRUEBA.TXT.txt"  # Ruta del archivo a recuperar
    $RecoveryDestination = "C:\PRUEBA RECUPERACION"                        # Ruta donde se recuperará el archivo

    try {
        # Obtener el conjunto de copias de seguridad más reciente
        $BackupSet = Get-WBBackupSet | Sort-Object -Property BackupTime -Descending | Select-Object -First 1
        if (-not $BackupSet) { throw "No se encontró ningún conjunto de copias de seguridad disponible." }

        # Recuperar el archivo
        Start-WBFileRecovery -BackupSet $BackupSet -SourcePath $FileToRecover -TargetPath $RecoveryDestination -Option OverwriteIfExists -Force

        # Verificar si la recuperación fue exitosa
        if (!(Test-Path -Path "$RecoveryDestination\$(Split-Path -Leaf $FileToRecover)")) {
            throw "El archivo no se pudo recuperar."
        }

        # Confirmación de recuperación
        $Body += "`nRecuperación exitosa del archivo: $FileToRecover"
    } catch {
        $Body += "`nError al intentar recuperar el archivo: $($_.Exception.Message)"
    }

    # Añadir el pie de página
    $Body += "`n" + $footerMessage

    # Crear el mensaje de correo electrónico con prioridad normal
    $Message = New-Object Net.Mail.MailMessage
    $Message.From = "revisiones@grupoicsolutions.com"
    $Message.To.Add($ToEmail)
    $Message.Bcc.Add($BccEmail)
    $Message.Subject = "$CompanyName Backup - Correcta"
    $Message.Body = $Body
    $Message.Priority = "Normal"

    # Enviar el correo electrónico
    $SMTPClient = New-Object Net.Mail.SmtpClient("smtp.grupoicsolutions.com", 25)
    $SMTPClient.Send($Message)
}
