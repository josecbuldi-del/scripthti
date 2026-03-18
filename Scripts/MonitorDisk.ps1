# MonitorDisk.ps1

# Parâmetros iniciais do script #

$driveToMonitor = "C:"   # Substitua pela letra do disco desejado
$usageThreshold = 90     # Substitua pelo limite de uso desejado (%)

# Funcao para Enviar E-mail
function Send-Email {
    param (
        [string]$Subject,
        [string]$Body
    )
    
    # Configuracoes de e-mail
    $smtpServer = "smtp.zoho.com"
    $smtpPort = 587
    $smtpUser = "prb2020@zohomail.com"
    $smtpPassword = "MM4FT/rq4JaGEu&"
    $recipient = "suporte@horusti.com.br"
    $sender = "prb2020@zohomail.com"

    try {
        # Criar cliente SMTP
        $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
        $smtpClient.EnableSsl = $true
        $smtpClient.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPassword)

        # Criar mensagem de e-mail
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $sender
        $mailMessage.To.Add($recipient)
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $Body

        # Enviar e-mail
        $smtpClient.Send($mailMessage)
        Write-Host "E-mail enviado com sucesso!" -ForegroundColor Green
    } catch {
        Write-Error "Erro ao enviar o e-mail: $_"
    }
}

# Funcao para verificar e exibir o uso do armazenamento
function Monitor-DiskSpace {
    param (
        [string]$DriveLetter = "C:",       # Unidade padrao
        [int]$Threshold = 90              # Limite de uso em percentual
    )

    try {
        # Verificar se a unidade existe
        $drive = Get-PSDrive -Name $DriveLetter.TrimEnd(':') -ErrorAction Stop

        # Obter nome do computador, endereco IP e usuario
        $computerName = $env:COMPUTERNAME
        $ipAddress = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 | Select-Object -ExpandProperty IPV4Address)
        $currentUser = $env:USERNAME

        # Calcular o percentual de uso do disco
        $usedSpacePercent = [math]::Round((($drive.Used / ($drive.Free + $drive.Used)) * 100), 2)

        # Exibir informacoes na tela
        Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Informacoes sobre a unidade: $DriveLetter" -ForegroundColor Green
        Write-Host "Nome do Computador: $computerName" -ForegroundColor Green
        Write-Host "Usuario Atual: $currentUser" -ForegroundColor Green
        Write-Host "Endereco IP: $ipAddress" -ForegroundColor Green
        Write-Host "`nPercentual de Uso: ${usedSpacePercent}%" -ForegroundColor Green
        Write-Host "`nEspaco Total: $([math]::Round(($drive.Free + $drive.Used) / 1GB, 2)) GB" -ForegroundColor Green
        Write-Host "Espaco Livre: $([math]::Round($drive.Free / 1GB, 2)) GB" -ForegroundColor Green
        Write-Host "Espaco Usado: $([math]::Round($drive.Used / 1GB, 2)) GB" -ForegroundColor Green
        Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan

        # Verificar se o percentual de uso esta acima do limite
        if ($usedSpacePercent -ge $Threshold) {
            Write-Host "ALERTA: O disco $DriveLetter esta com ${usedSpacePercent}% de uso, acima do limite de ${Threshold}%!" -ForegroundColor Red

            # Enviar e-mail automaticamente
            $subject = "Alerta de Uso do Disco - $DriveLetter ($computerName)"
            $body = @"
Alerta de Uso do Disco - Unidade $DriveLetter

Informacoes do Computador:
- Nome do Computador: $computerName
- Usuario Atual: $currentUser
- Endereco IP: $ipAddress

Informacoes do Disco:
- Percentual de Uso: ${usedSpacePercent}%
- Espaco Total: $([math]::Round(($drive.Free + $drive.Used) / 1GB, 2)) GB
- Espaco Livre: $([math]::Round($drive.Free / 1GB, 2)) GB
- Espaco Usado: $([math]::Round($drive.Used / 1GB, 2)) GB

Recomendado liberar espaco imediatamente!!!
"@
            # Chamar funcao para enviar e-mail
            Send-Email -Subject $subject -Body $body
        }

    } catch {
        Write-Host "A unidade '$DriveLetter' nao foi encontrada." -ForegroundColor Red
        Write-Host "Unidade invalida. Tente novamente." -ForegroundColor Yellow
    }
}

# Executar monitoramento
Monitor-DiskSpace -DriveLetter $driveToMonitor -Threshold $usageThreshold
