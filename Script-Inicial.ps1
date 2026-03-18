# Atencao: antes de executar este script, abra o PowerShell como Administrador e execute:
# Set-ExecutionPolicy Bypass -Scope Process -Force

Clear-Host

# =========================
# CONFIGURACAO CENTRAL
# =========================
$GitBaseUrl = "https://raw.githubusercontent.com/josecbuldi-del/scripthti/main"

$BasePath     = "C:\_HTI"
$ScriptsPath  = Join-Path $BasePath "Scripts"
$UtilsPath    = Join-Path $BasePath "Utilitarios"
$ProgramsPath = Join-Path $BasePath "Programas"
$CleanupPath  = Join-Path $ScriptsPath "Limpeza"

function Initialize-HTIStructure {
    @($BasePath, $ScriptsPath, $UtilsPath, $ProgramsPath, $CleanupPath) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -Path $_ -ItemType Directory -Force | Out-Null
        }
    }
}

function Get-GitRawUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )
    return "$GitBaseUrl/$RelativePath"
}

function Download-FromGit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $url = Get-GitRawUrl -RelativePath $RelativePath
    Invoke-WebRequest -Uri $url -OutFile $DestinationPath -UseBasicParsing -ErrorAction Stop
}

# =========================
# FUNCOES BASICAS
# =========================

function Check-WingetInstalled {
    $wingetPackage = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*Microsoft.DesktopAppInstaller*" }

    if ($null -eq $wingetPackage -or -not $wingetPackage.InstallLocation) {
        Write-Host "Winget nao esta instalado ou nao foi encontrado." -ForegroundColor Yellow
        return $false
    }

    $wingetPath = $wingetPackage.InstallLocation
    $wingetExePath = Join-Path $wingetPath "winget.exe"

    if (Test-Path $wingetExePath) {
        $env:Path = $env:Path + ";" + $wingetPath
    } else {
        Write-Host "Winget.exe nao encontrado no caminho esperado: $wingetExePath" -ForegroundColor Red
        return $false
    }

    try {
        $null = Get-Command winget -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Show-LoadingAnimation {
    $animation = @("|", "/", "-", "\")
    $i = 0
    $count = 0
    while ($count -lt 5) {
        Write-Host -NoNewline "$($animation[$i % $animation.Length])" -ForegroundColor Green
        Start-Sleep -Milliseconds 200
        Write-Host -NoNewline "`r"
        $i++
        $count++
    }
    Write-Host ""
}

function Accept-WingetTerms {
    echo Y | winget list > $null
}

function Get-ScoreColor {
    param([double]$Score)

    if ($Score -ge 1 -and $Score -le 3) {
        return "Red"
    } elseif ($Score -ge 4 -and $Score -le 5) {
        return "DarkYellow"
    } elseif ($Score -ge 5.1 -and $Score -le 6.9) {
        return "Yellow"
    } elseif ($Score -ge 7 -and $Score -le 8.9) {
        return "Green"
    } elseif ($Score -ge 9) {
        return "Blue"
    } else {
        return "White"
    }
}

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# =========================
# E-MAIL
# =========================

function Send-Email {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$Body
    )

    $smtpServer = "smtp.zoho.com"
    $smtpPort   = 587
    $smtpUser   = "prb2020@zohomail.com"
    $recipient  = "suporte@horusti.com.br"
    $sender     = "prb2020@zohomail.com"

    Write-Host "`nInforme a senha do e-mail para envio SMTP:" -ForegroundColor Yellow
    $securePassword = Read-Host "Senha SMTP" -AsSecureString
    $credential = New-Object System.Management.Automation.PSCredential($smtpUser, $securePassword)

    try {
        Send-MailMessage `
            -SmtpServer $smtpServer `
            -Port $smtpPort `
            -UseSsl `
            -Credential $credential `
            -From $sender `
            -To $recipient `
            -Subject $Subject `
            -Body $Body

        Write-Host "E-mail enviado com sucesso!" -ForegroundColor Green
        Start-Sleep -Seconds 2
        Clear-Host
    }
    catch {
        Write-Host "Erro ao enviar o e-mail: $($_.Exception.Message)" -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

# =========================
# INFORMACOES DO WINDOWS
# =========================

function Info-Windows {
    Clear-Host
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host "Aguarde, coletando informacoes do computador..." -ForegroundColor Yellow
    Show-LoadingAnimation

    $computerInfo = Get-ComputerInfo
    Clear-Host

    Write-Output "--------------------------------------------------------------"
    Write-Host "             -------   INFORMACOES DO WINDOWS   -------         " -ForegroundColor Blue
    Write-Output "--------------------------------------------------------------`n"

    Write-Output "Script para o Setup de Windows 10 e 11`n"
    Write-Output "Data de Execucao: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')"
    Write-Output "Nome do Computador: $env:COMPUTERNAME"
    Write-Output "Versao do Windows: $($computerInfo.OsName)"

    $winVersion = $null
    try {
        $winVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion
        if (-not $winVersion) {
            $winVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId -ErrorAction SilentlyContinue).ReleaseId
        }
    } catch {
        $winVersion = "Nao disponivel"
    }

    Write-Output "Release: $winVersion"
    Write-Output "Data Instalacao do Windows: $($computerInfo.OsInstallDate)"

    $psVersion = $PSVersionTable.PSVersion
    Write-Output "Versao do PowerShell: $psVersion"

    if ($psVersion.Major -lt 5) {
        Write-Host "ATENCAO: O PowerShell esta desatualizado! Considere atualizar para a versao mais recente." -ForegroundColor Red
    }

    $wingetStatus = Check-WingetInstalled
    if ($wingetStatus) {
        Write-Host "Winget: Instalado" -ForegroundColor Green
    } else {
        Write-Host "Winget: Nao Instalado" -ForegroundColor Red
    }

    Write-Output "Nome do Dominio (ou Workgroup): $($computerInfo.CsDomain)"
    Write-Output "Marca do Computador: $($computerInfo.CsManufacturer)"
    Write-Output "Modelo do Computador: $($computerInfo.CsModel)"
    Write-Output "Serie do Computador: $($computerInfo.CsSystemFamily)"

    $processorName = (Get-WmiObject Win32_Processor).Name
    Write-Output "Processador: $processorName"

    $totalMemoryGB = [math]::Round($computerInfo.OsTotalVisibleMemorySize / 1MB, 2)
    Write-Output "Total de Memoria (GB): $totalMemoryGB"

    $uptime = $computerInfo.OsUptime
    $uptimeDays = $uptime.Days
    $uptimeFormatted = $uptime.ToString('d\.hh\:mm\:ss')

    if ($uptimeDays -ge 3) {
        Write-Host "Tempo de Uptime: $uptimeFormatted" -ForegroundColor Red
    } elseif ($uptimeDays -ge 1) {
        Write-Host "Tempo de Uptime: $uptimeFormatted" -ForegroundColor Yellow
    } else {
        Write-Output "Tempo de Uptime: $uptimeFormatted"
    }

    Write-Output "Nome do usuario: $env:USERNAME"

    try {
        $localIP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 -ErrorAction Stop).IPV4Address.IPAddressToString
        Write-Output "IP do computador local: $localIP"
    } catch {
        $localIP = "Nao foi possivel determinar"
        Write-Output "IP do computador local: $localIP"
    }

    function Get-PublicIP {
        param([string]$url)

        try {
            if (-not ("System.Net.Http.HttpClient" -as [Type])) {
                Add-Type -AssemblyName "System.Net.Http"
            }

            $client = New-Object System.Net.Http.HttpClient
            $client.Timeout = New-TimeSpan -Seconds 10
            $responseTask = $client.GetStringAsync($url)

            if ($responseTask.Wait(10000)) {
                return $responseTask.Result
            } else {
                throw "Tempo limite excedido"
            }
        } catch {
            return "Nao foi possivel obter o endereco IP"
        } finally {
            if ($client) { $client.Dispose() }
        }
    }

    $publicIPv4 = Get-PublicIP -url "https://api.ipify.org?format=text"
    $publicIPv6 = Get-PublicIP -url "https://api64.ipify.org?format=text"

    if ($publicIPv4 -like "*Nao foi possivel obter*" -and $publicIPv6 -like "*Nao foi possivel obter*") {
        Write-Output "Nao foi possivel obter nenhum endereco IP publico."
    } elseif ($publicIPv4 -eq $publicIPv6 -or $publicIPv6 -like "*Nao foi possivel obter*") {
        Write-Output "Endereco IP Publico (IPv4): $publicIPv4"
        Write-Output "Nao ha endereco IP Publico IPv6 disponivel ou identificado."
    } else {
        Write-Output "Endereco IP Publico (IPv4): $publicIPv4"
        Write-Output "Endereco IP Publico (IPv6): $publicIPv6"
    }

    Write-Output "--------------------------------------------------------------"
    Write-Host "## Informacoes de Desempenho (WinSAT)"

    try {
        $winSAT = Get-CimInstance Win32_WinSAT -ErrorAction Stop
        $scores = @{
            "Processador" = $winSAT.CPUScore
            "Graficos 3D" = $winSAT.D3DScore
            "Disco"       = $winSAT.DiskScore
            "Graficos"    = $winSAT.GraphicsScore
            "Memoria"     = $winSAT.MemoryScore
        }

        $minScore = 10.0
        $maxScore = 0.0
        $minComponentName = ""
        $maxComponentName = ""

        Write-Host "`nPontuacoes de Desempenho (WinSAT):" -ForegroundColor White
        foreach ($component in $scores.GetEnumerator()) {
            $score = $component.Value
            $name = $component.Key
            $color = Get-ScoreColor -Score $score
            Write-Host "  ${name}: $score" -ForegroundColor $color

            if ($score -lt $minScore) {
                $minScore = $score
                $minComponentName = $name
            }
            if ($score -gt $maxScore) {
                $maxScore = $score
                $maxComponentName = $name
            }
        }

        Write-Host ""
        Write-Host "  Sua menor nota e no componente: $minComponentName (Score: $minScore)" -ForegroundColor Yellow
        Write-Host "  Sua maior nota e no componente: $maxComponentName (Score: $maxScore)" -ForegroundColor Green
        Write-Host "  Nivel de Experiencia do Windows (WinSPRLevel): $($winSAT.WinSPRLevel)" -ForegroundColor (Get-ScoreColor -Score $winSAT.WinSPRLevel)
    } catch {
        Write-Host "Nao foi possivel obter as pontuacoes do WinSAT." -ForegroundColor Red
    }

    Write-Output "--------------------------------------------------------------`n"

    $tiService = Get-Service -Name "TiService" -ErrorAction SilentlyContinue
    if ($null -ne $tiService) {
        if ($tiService.Status -eq "Running") {
            Write-Host "Agente TiFlux: Instalado e em Execucao" -ForegroundColor Green
        } else {
            Write-Host "Agente TiFlux: Instalado e Parado" -ForegroundColor Red
        }
    } else {
        Write-Host "Agente TiFlux: Nao Instalado" -ForegroundColor Red
    }

    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Yellow
    Write-Output "ATENCAO: Este Script funciona apenas na linguagem Portugues/BR"
    Write-Host "--------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host ""

    $sendToSupport = Read-Host "Deseja adicionar o endereco do AnyDesk e enviar estas informacoes por e-mail? (S/[N])"
    if ($sendToSupport -notmatch "^[Ss](im)?$") {
        Write-Host "Voltando ao menu principal..." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        Clear-Host
        return
    }

    $anydeskAddress = Read-Host "Digite o endereco do AnyDesk"
    $clientName = Read-Host "Nome do Cliente? (ou pressione Enter para pular)"
    $additionalNotes = Read-Host "Deseja adicionar alguma observacao? [Enter para nao]"

    $emailContent = @(
        "Nome do Computador: $env:COMPUTERNAME",
        "Endereco IP Local: $localIP",
        "Endereco IP Publico (IPv4): $publicIPv4",
        "Endereco IP Publico (IPv6): $publicIPv6",
        "Endereco do AnyDesk: $anydeskAddress",
        "",
        "Data de Execucao: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')",
        "Agente TiFlux: $(if ($null -ne $tiService) { if ($tiService.Status -eq 'Running') { 'Instalado e em Execucao' } else { 'Instalado e Parado' } } else { 'Nao Instalado' })",
        "Versao do Windows: $($computerInfo.OsName)",
        "Release: $winVersion",
        "Data Instalacao do Windows: $($computerInfo.OsInstallDate)",
        "Nome do Dominio (ou Workgroup): $($computerInfo.CsDomain)",
        "Marca do Computador: $($computerInfo.CsManufacturer)",
        "Modelo do Computador: $($computerInfo.CsModel)",
        "Serie do Computador: $($computerInfo.CsSystemFamily)",
        "Processador: $processorName",
        "Total de Memoria (GB): $totalMemoryGB",
        "Tempo de Uptime: $uptimeFormatted",
        "Nome do usuario: $env:USERNAME"
    )

    if (-not [string]::IsNullOrWhiteSpace($clientName)) {
        $emailContent += "Nome do Cliente: $clientName"
    }
    if (-not [string]::IsNullOrWhiteSpace($additionalNotes)) {
        $emailContent += "Observacoes: $additionalNotes"
    }

    $subject = "Informacoes do Windows - $env:COMPUTERNAME"
    if (-not [string]::IsNullOrWhiteSpace($clientName)) {
        $subject += " - $clientName"
    }

    Send-Email -Subject $subject -Body ($emailContent -join "`n")

    Write-Host "As informacoes foram enviadas para o suporte." -ForegroundColor Green
    Start-Sleep -Seconds 2
    Clear-Host
}

# =========================
# COMPUTADOR / AGENTE
# =========================

function Rename-ComputerCustom {
    Write-Host "Nome atual do computador: $env:COMPUTERNAME"

    do {
        $novoNome = Read-Host "Digite o novo nome para o computador (ou pressione Enter para cancelar)"

        if ([string]::IsNullOrWhiteSpace($novoNome)) {
            Write-Host "Operacao cancelada."
            Start-Sleep -Seconds 1
            Clear-Host
            return
        }

        if ($novoNome -match "^[a-zA-Z0-9\.\-_]+$") {
            try {
                Rename-Computer -NewName $novoNome -Force
                Write-Host "Computador renomeado para '$novoNome'. Reinicie para aplicar as mudancas." -ForegroundColor Yellow
                return
            } catch {
                Write-Host "Erro ao renomear o computador: $($_.Exception.Message)" -ForegroundColor Red
                pause
            }
        } else {
            Write-Host "Nome invalido! Use apenas letras, numeros, '-', '.' e '_'." -ForegroundColor Red
        }
    } while ($true)
}

function Get-ClientLinksFromCSV {
    $csvFilePath = Join-Path $ScriptsPath "clientes.csv"

    if (-not (Test-Path $csvFilePath)) {
        Write-Host "Baixando o arquivo dos links dos clientes..." -ForegroundColor Yellow
        try {
            Download-FromGit -RelativePath "Scripts/clientes.csv" -DestinationPath $csvFilePath
            Write-Host "Arquivo baixado com sucesso!" -ForegroundColor Green
        }
        catch {
            Write-Host "Erro ao baixar o arquivo CSV: $($_.Exception.Message)" -ForegroundColor Red
            return $null
        }
    }

    try {
        return Import-Csv -Path $csvFilePath
    }
    catch {
        Write-Host "Erro ao ler o clientes.csv: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Download-And-InstallAgent {
    param (
        [string]$clientName,
        [string]$downloadLink
    )

    $fileName = "C:\_HTI\$clientName.msi"

    Write-Host "Baixando o agente para $clientName..."
    try {
        Invoke-WebRequest -Uri $downloadLink -OutFile $fileName
        Write-Host "Agente baixado com sucesso!" -ForegroundColor Green
    } catch {
        Write-Host "Erro ao baixar o agente. Verifique o link de download." -ForegroundColor Red
        return
    }

    Write-Host "Instalando o agente de $clientName..."
    Start-Process msiexec.exe -ArgumentList "/i", $fileName, "/quiet", "/norestart" -NoNewWindow -Wait
    Write-Host "Instalacao concluida para $clientName."
    pause
    Clear-Host
}

function Install-Agente {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "Instalar Agente" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    $service = Get-Service -Name "TiService" -ErrorAction SilentlyContinue
    if ($service -ne $null -and $service.Status -eq "Running") {
        Write-Host "`nO agente TiFlux ja esta instalado no computador." -ForegroundColor Yellow
        $removeConfirmation = Read-Host "Deseja remover o agente antes de prosseguir? (S/N)"
        if ($removeConfirmation -match "S") {
            $programsToRemove = @("Ti Service And Agent", "TiPeerToPeer", "timessenger")

            foreach ($program in $programsToRemove) {
                try {
                    Write-Host "`nRemovendo $program..." -ForegroundColor Cyan
                    winget uninstall --name $program -e
                } catch {
                    Write-Host "Erro ao tentar remover $program. Pode ser que ele nao esteja instalado." -ForegroundColor Yellow
                }
            }

            $installConfirmation = Read-Host "Deseja continuar com a instalacao do Agente? (S/N)"
            if ($installConfirmation -notmatch "S") {
                Write-Host "Instalacao cancelada." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                Clear-Host
                return
            }
        } else {
            Write-Host "Instalacao cancelada pelo usuario." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Clear-Host
            return
        }
    }

    $csvPath = Join-Path $ScriptsPath "clientes.csv"
    if (-not (Test-Path $csvPath)) {
        Write-Host "Baixando arquivo clientes.csv..." -ForegroundColor Yellow
        try {
            Download-FromGit -RelativePath "Scripts/clientes.csv" -DestinationPath $csvPath
            Write-Host "Arquivo clientes.csv baixado com sucesso!" -ForegroundColor Green
        } catch {
            Write-Host "Erro ao baixar o arquivo clientes.csv: $($_.Exception.Message)" -ForegroundColor Red
            Clear-Host
            return
        }
    }

    $clients = Get-ClientLinksFromCSV
    if ($clients -eq $null) {
        Write-Host "Erro ao carregar os dados dos clientes." -ForegroundColor Red
        Clear-Host
        return
    }

    while ($true) {
        $nomeCliente = Read-Host "`nDigite o nome (ou parte do nome) do cliente que deseja instalar o agente (ou '0' para sair)"

        if ($nomeCliente -eq "0") {
            Write-Host "Saindo para o menu principal..." -ForegroundColor Cyan
            Remove-Item -Path $csvPath -Force -ErrorAction SilentlyContinue
            Write-Host "Arquivo clientes.csv excluido." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Clear-Host
            return
        }

        $resultados = $clients | Where-Object { $_.Cliente -like "*$nomeCliente*" }

        if ($resultados.Count -eq 0) {
            Write-Host "Nenhum cliente encontrado com esse nome." -ForegroundColor Red
            $listarTodos = Read-Host "Deseja listar todos os clientes? (S/N)"
            if ($listarTodos -match "S") {
                Write-Host "`nLista de clientes:" -ForegroundColor Green
                $clients | ForEach-Object { Write-Host "$($_.Cliente) - Indice: $([array]::IndexOf($clients, $_) + 1)" }
            }
        } else {
            Write-Host "`nClientes encontrados:" -ForegroundColor Cyan
            $resultados | ForEach-Object {
                $index = [array]::IndexOf($clients, $_) + 1
                Write-Host "$($_.Cliente) - Indice: $index"
            }

            $input = Read-Host "`nDigite o numero do cliente (indice) para instalar ou '0' para cancelar"
            if ($input -eq "0") {
                Write-Host "Instalacao cancelada." -ForegroundColor Yellow
                pause
                Clear-Host
                return
            }

            $clientIndex = 0
            if ([int]::TryParse($input, [ref]$clientIndex) -and $clientIndex -ge 1 -and $clientIndex -le $clients.Count) {
                $selectedClient = $clients[$clientIndex - 1]
                Write-Host "`nVoce selecionou o cliente: $($selectedClient.Cliente)" -ForegroundColor Green
                $confirmation = Read-Host "Confirma? (S/N)"
                if ($confirmation -match "S") {
                    Download-And-InstallAgent -clientName $selectedClient.Cliente -downloadLink $selectedClient.'Link-Agente'
                    return
                } else {
                    Write-Host "Instalacao cancelada. Voltando a selecao..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                }
            } else {
                Write-Host "Indice invalido. Tente novamente." -ForegroundColor Red
            }
        }
    }
}

# =========================
# INSTALACAO DE PROGRAMAS
# =========================

function Install-BasicPackage {
    if (-not (Check-WingetInstalled)) {
        Write-Host "`nO Winget nao esta instalado neste computador. Por favor, instale-o antes de prosseguir." -ForegroundColor Red
        pause
        Clear-Host
        return
    }

    Accept-WingetTerms

    $packages = @(
        @{ Name = "7-Zip"; Command = "winget install --id=7zip.7zip -e" },
        @{ Name = "Google Chrome"; Command = "winget install --id=Google.Chrome -e" },
        @{ Name = "AnyDesk"; Command = "winget install --id=AnyDeskSoftwareGmbH.AnyDesk -e" },
        @{ Name = "Lightshot"; Command = "winget install --id=Skillbrains.Lightshot -e" },
        @{ Name = "Adobe Acrobat Reader (64-bit)"; Command = "winget install --id=Adobe.Acrobat.Reader.64-bit -e" }
    )

    Write-Host "`nOs seguintes programas serao instalados e/ou atualizados:" -ForegroundColor Cyan
    foreach ($package in $packages) {
        Write-Host "- $($package.Name)" -ForegroundColor Yellow
    }

    $confirm = Read-Host "`nTem certeza de que deseja iniciar a instalacao e atualizacao dos programas? (S/N)"
    if ($confirm -ne "S") {
        Write-Host "Operacao cancelada pelo usuario." -ForegroundColor Red
        Start-Sleep -Seconds 2
        Clear-Host
        return
    }

    foreach ($package in $packages) {
        $name = $package.Name
        $command = $package.Command

        $installed = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
                     Where-Object { $_.DisplayName -like "*$name*" }

        if ($installed) {
            Write-Host "`n$name ja esta instalado." -ForegroundColor Yellow
            $choice = Read-Host "Deseja atualizar $name? (Padrao: N, pressione Enter para Nao, ou digite 'S' para Sim)"
            if ($choice -eq "S") {
                Write-Host "Atualizando $name..." -ForegroundColor Green
                Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c $command --silent --force"
                Write-Host "$name foi atualizado com sucesso." -ForegroundColor Green
            } else {
                Write-Host "Atualizacao de $name foi cancelada." -ForegroundColor Red
            }
        } else {
            Write-Host "`n$name nao esta instalado. Instalando agora..." -ForegroundColor Yellow
            Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c $command --silent"
            Write-Host "$name foi instalado com sucesso." -ForegroundColor Green
        }
    }

    Write-Host "`nOperacao concluida. Pressione qualquer tecla para voltar ao menu principal..." -ForegroundColor Cyan
    [void][System.Console]::ReadKey($true)
    Clear-Host
}

function Download-Utilities {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "         Download de Aplicativos Utilitarios de Suporte" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------" -ForegroundColor DarkGray

    $confirmDownload = Read-Host "`nTem certeza que deseja baixar os aplicativos utilitarios? (S/N)"
    if ($confirmDownload -notmatch "^(Sim|sim|S|s)$") {
        Write-Host "`nOperacao cancelada. Retornando ao menu principal ..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Clear-Host
        Main
        return
    }

    $fileListPath = Join-Path $UtilsPath "files.txt"

    Write-Host "Baixando a lista de arquivos..."
    try {
        Download-FromGit -RelativePath "Scripts/files.txt" -DestinationPath $fileListPath
    } catch {
        Write-Host "Erro ao baixar files.txt: $($_.Exception.Message)" -ForegroundColor Red
        pause
        Clear-Host
        return
    }

    $files = Get-Content $fileListPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($fileName in $files) {
        $cleanName = $fileName.Trim()
        $outputPath = Join-Path -Path $UtilsPath -ChildPath $cleanName

        Write-Host "Baixando: $cleanName..."
        try {
            Download-FromGit -RelativePath "Utilitarios/$cleanName" -DestinationPath $outputPath
            Write-Host "$cleanName baixado com sucesso." -ForegroundColor Green
        } catch {
            Write-Host "Falha ao baixar $cleanName : $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "`nTodos os downloads foram processados.`n" -ForegroundColor Green
    Write-Host "Arquivos disponiveis em C:\_HTI\Utilitarios`n" -ForegroundColor Green
    pause
    Clear-Host
}

function Install-Office {
    try {
        Write-Host "`nDeseja instalar o Office neste computador? [S/n]" -ForegroundColor Yellow
        $confirm = Read-Host "(Pressione Enter para SIM)"
        if ($confirm -and $confirm -notmatch "^[Ss]$") {
            Write-Host "`nOperacao cancelada pelo usuario. Retornando ao menu principal..." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Clear-Host
            return
        }

        try {
            $officeProcesses = Get-Process | Where-Object { $_.Name -like "*office*" }
            if ($officeProcesses) {
                Write-Host "`nOs seguintes processos relacionados ao Office foram encontrados:" -ForegroundColor Yellow
                $officeProcesses | ForEach-Object { Write-Host "- $($_.Name)" -ForegroundColor Cyan }

                Write-Host "`nDeseja forcar o encerramento desses processos? [S/n]" -ForegroundColor Red
                $forceClose = Read-Host "(Pressione Enter para SIM)"
                if ($forceClose -and $forceClose -notmatch "^[Ss]$") {
                    Write-Host "`nOperacao cancelada pelo usuario. Retornando ao menu principal..." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    return
                }

                $officeProcesses | ForEach-Object {
                    try {
                        Stop-Process -Id $_.Id -Force
                        Write-Host "Processo '$($_.Name)' encerrado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Host "Erro ao encerrar o processo '$($_.Name)': $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        } catch {
            Write-Host "Erro ao verificar ou encerrar processos relacionados ao Office: $($_.Exception.Message)" -ForegroundColor Red
        }

        $officeInstallerPath = Join-Path $ProgramsPath "OfficeSetup.exe"

        Write-Host "`nBaixando o instalador do Office..." -ForegroundColor Cyan
        try {
            Download-FromGit -RelativePath "Programas/OfficeSetup.exe" -DestinationPath $officeInstallerPath
            Write-Host "Instalador do Office baixado com sucesso para '$officeInstallerPath'." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao baixar o instalador do Office: $($_.Exception.Message)" -ForegroundColor Red
            return
        }

        Write-Host "`nIniciando a instalacao do Office..." -ForegroundColor Yellow
        try {
            Start-Process -FilePath $officeInstallerPath
            Write-Host "Continue com a instalacao manual do Office" -ForegroundColor Green
        } catch {
            Write-Host "Erro ao executar o instalador do Office: $($_.Exception.Message)" -ForegroundColor Red
        }
    } catch {
        Write-Host "Ocorreu um erro durante a instalacao do Office: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host ""
    pause
    Clear-Host
}

function Activate-WindowsOffice {
    try {
        Write-Host "`nDeseja abrir as ferramentas oficiais de ativacao do Windows e Office? [S/n]" -ForegroundColor Yellow
        $confirm = Read-Host "(Pressione Enter para SIM)"
        if ($confirm -and $confirm -notmatch "^[Ss]$") {
            Write-Host "`nOperacao cancelada pelo usuario. Retornando ao menu principal..." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Clear-Host
            return
        }

        Write-Host "`nAbrindo Configuracoes de Ativacao do Windows..." -ForegroundColor Cyan
        Start-Process "ms-settings:activation"

        Write-Host "Abrindo portal de conta do Microsoft 365..." -ForegroundColor Cyan
        Start-Process "https://portal.office.com/account/"

        Write-Host "`nFinalize manualmente a ativacao conforme o licenciamento da organizacao." -ForegroundColor Yellow
        Write-Host "`nPressione qualquer tecla para retornar ao menu principal..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Write-Host "Ocorreu um erro durante a abertura das ferramentas de ativacao: $($_.Exception.Message)" -ForegroundColor Red
    }
    Clear-Host
}

function Remove-Apps {
    if (-not (Check-WingetInstalled)) {
        Write-Host "`nO Winget nao esta instalado neste computador. Por favor, instale-o antes de prosseguir." -ForegroundColor Red
        pause
        Clear-Host
        return
    }

    Accept-WingetTerms

    $appsToRemove = @(
        "Tactical RMM Agent",
        "Agente Milvus",
        "Mesh Agent",
        "Solitaire",
        "Jogo de Copas",
        "Spades",
        "Solitaire & Casual Games",
        "Game Bar",
        "Game Speech Window",
        "Noticias",
        "Xbox",
        "Xbox TCUI",
        "Xbox Game Bar Plugin",
        "Xbox Console Companion",
        "Xbox Game Speech Window",
        "Xbox Identity Provider",
        "WebAdvisor da McAfee",
        "Dropbox - promocao",
        "Noticias",
        "Microsoft Bing"
    )

    $wingetOutput = winget list | ForEach-Object {
        $fields = $_ -split '\s{2,}'
        [PSCustomObject]@{
            Name = $fields[0]
            Id = if ($fields.Count -gt 1) { $fields[1] } else { "" }
        }
    }

    $installedApps = $wingetOutput | Where-Object {
        $appsToRemove -contains $_.Name -or ($appsToRemove | ForEach-Object { $_.ToLower() }) -contains $_.Name.ToLower()
    }

    if ($installedApps.Count -eq 0) {
        Write-Host "`nNenhum dos aplicativos especificados esta instalado no sistema.`n" -ForegroundColor Yellow
        pause
        Clear-Host
        return
    }

    Write-Host "`nOs seguintes aplicativos serao removidos:`n" -ForegroundColor Cyan
    $installedApps | ForEach-Object { Write-Host $_.Name }

    $confirmation = Read-Host "Pressione [Enter] para continuar ou digite 'N' para cancelar"
    if ($confirmation -eq 'N') {
        Write-Host "`nA operacao foi cancelada.`n" -ForegroundColor Yellow
        pause
        Clear-Host
        return
    }

    foreach ($app in $installedApps) {
        Write-Host "Removendo $($app.Name)..." -ForegroundColor Green
        try {
            winget uninstall --name "$($app.Name)" --silent
            Write-Host "$($app.Name) foi removido com sucesso.`n" -ForegroundColor Green
        } catch {
            Write-Host "Falha ao remover $($app.Name). Verifique manualmente." -ForegroundColor Red
        }
    }

    Write-Host "Todos os aplicativos especificados foram processados." -ForegroundColor Cyan
    pause
    Clear-Host
}

# =========================
# USUARIOS LOCAIS
# =========================

function Show-ActiveLocalUsers {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "         Lista dos Usuarios Locais e Ativos do Windows" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    $localUsers = Get-LocalUser | Where-Object { $_.Enabled -eq $true }

    if ($localUsers) {
        foreach ($user in $localUsers) {
            Write-Host $user.Name -ForegroundColor Cyan
        }
    } else {
        Write-Host "Nenhum usuario ativo encontrado." -ForegroundColor Yellow
    }

    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Host ""
    pause
    Manage-LocalUsers
}

function Show-LocalAdminUsers {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "            Usuarios do Grupo Administradores" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    $domain = (Get-WmiObject Win32_ComputerSystem).Domain
    $isDomainJoined = $domain -ne (Get-WmiObject Win32_ComputerSystem).Name

    if ($isDomainJoined) {
        Write-Host "`nO computador esta no dominio: $domain" -ForegroundColor Yellow
        Write-Host "Listando membros usando 'net localgroup'`n" -ForegroundColor Green
        cmd /c "net localgroup Administradores"
    } else {
        Write-Host "`nO computador esta em um grupo de trabalho: $domain" -ForegroundColor Yellow
        try {
            $admins = Get-LocalGroupMember -Group "Administradores"

            if ($admins) {
                Write-Host "`nMembros do grupo Administradores:" -ForegroundColor Green
                foreach ($admin in $admins) {
                    Write-Host $admin.Name -ForegroundColor Cyan
                }
            } else {
                Write-Host "Nenhum usuario administrador encontrado." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "`nErro ao acessar os membros do grupo Administradores." -ForegroundColor Red
            Write-Host "Verifique se voce tem permissoes administrativas." -ForegroundColor Yellow
        }
    }

    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Host ""
    pause
    Manage-LocalUsers
}

function Create-NewLocalUser {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "          CRIAR NOVO USUARIO LOCAL" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green

    do {
        Write-Host ""
        $username = Read-Host "Digite o nome do novo usuario (ou pressione '0' para voltar ao menu anterior)"

        if (-not $username) {
            Write-Host "`nO nome do usuario nao pode estar em branco." -ForegroundColor Red
            continue
        }

        if ($username -eq "0") {
            Write-Host "`nRetornando ao menu anterior..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Manage-LocalUsers
            return
        }

        if ($username -notmatch '^[a-zA-Z0-9._-]+$') {
            Write-Host "`nO nome do usuario contem caracteres invalidos." -ForegroundColor Red
            continue
        }

        Write-Host "`nO nome do novo usuario sera: $username" -ForegroundColor Yellow
        $confirm = Read-Host "Confirme o nome do usuario (S/N)"
        if ($confirm -eq "S" -or $confirm -eq "s") {
            $existingUser = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
            if ($existingUser) {
                Write-Host "`nO usuario '$username' ja existe." -ForegroundColor Red
                $createNew = Read-Host "Deseja criar o usuario com outro nome? (S/N)"
                if ($createNew -eq "N" -or $createNew -eq "n") {
                    Clear-Host
                    Manage-LocalUsers
                    return
                }
            } else {
                break
            }
        }
    } while ($true)

    do {
        $password = Read-Host "`nDigite a senha do novo usuario" -AsSecureString
        $confirmPassword = Read-Host "Confirme a senha digitada" -AsSecureString

        if (([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))) -ne
            ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirmPassword)))) {
            Write-Host "`nAs senhas digitadas nao coincidem. Tente novamente." -ForegroundColor Red
        } else {
            Write-Host "`nSenha confirmada com sucesso." -ForegroundColor Green
            break
        }
    } while ($true)

    New-LocalUser -Name $username -Password $password | Out-Null

    $addToAdminsGroup = Read-Host "Deseja adicionar o usuario ao grupo Administradores? (S/N)"
    if ($addToAdminsGroup -eq "S" -or $addToAdminsGroup -eq "s") {
        Add-LocalGroupMember -Group "Administradores" -Member $username
        Write-Host "`nUsuario '$username' criado e adicionado ao grupo Administradores." -ForegroundColor Green
    } else {
        $GUsers = Get-LocalGroup | Where-Object { $_.SID -eq 'S-1-5-32-545' }
        Add-LocalGroupMember -Group $GUsers -Member $username
        Write-Host "`nUsuario '$username' criado no grupo Local padrao." -ForegroundColor Green
    }

    Set-LocalUser -Name $username -PasswordNeverExpires 1

    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "             PROCESSO FINALIZADO" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green

    pause
    Manage-LocalUsers
}

function Create-HtiUser {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "            CRIAR USUARIO HTI ADMINISTRADOR" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    $username = "HTI"
    $existingUser = Get-LocalUser -Name $username -ErrorAction SilentlyContinue

    if ($existingUser) {
        Write-Host "`nO usuario '$username' ja existe neste Windows." -ForegroundColor Red
        Write-Host "`nRetornando ao menu anterior ..."
        Start-Sleep -Seconds 2
        Manage-LocalUsers
        return
    }

    do {
        $password = Read-Host "Digite a senha do usuario HTI" -AsSecureString
        $passwordConfirmation = Read-Host "Confirme a senha digitada" -AsSecureString

        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
        $plainPasswordConfirmation = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordConfirmation))

        if ($plainPassword -ne $plainPasswordConfirmation) {
            Write-Host "`nAs senhas digitadas nao coincidem. Tente novamente." -ForegroundColor Red
        } else {
            Write-Host "`nSenha confirmada com sucesso!" -ForegroundColor Green
            break
        }
    } while ($true)

    New-LocalUser -Name $username -Password $password | Out-Null
    Add-LocalGroupMember -Group "Administradores" -Member $username
    Write-Host "`nUsuario $username criado no grupo Administradores`n" -ForegroundColor Green
    Set-LocalUser -Name $username -PasswordNeverExpires 1

    $enableRDP = Read-Host "Deseja ativar a Area de Trabalho Remota do Windows? (S/N)"
    if ($enableRDP -match "^[sS]$") {
        try {
            Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0
            $rdpRules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "*Remote Desktop*" }
            if ($rdpRules) {
                $rdpRules | ForEach-Object { Enable-NetFirewallRule -Name $_.Name }
                Write-Host "`nRegras de firewall do RDP ativadas com sucesso!" -ForegroundColor Green
            } else {
                Write-Host "`nNao foi possivel localizar as regras de firewall do RDP." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "`nErro ao ativar a Area de Trabalho Remota." -ForegroundColor Red
        }
    } else {
        Write-Host "`nArea de Trabalho Remota nao foi ativada." -ForegroundColor Yellow
    }

    Start-Sleep -Seconds 2
    pause
    Manage-LocalUsers
}

function Change-LocalUserPassword {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "            Trocar a Senha de um Usuario Local" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Output "`nQual usuario abaixo deseja trocar a senha?"

    $localUsers = Get-LocalUser | Where-Object { $_.Enabled -eq $true }
    $localUsers | Select-Object -ExpandProperty Name
    Write-Host "`nSe deseja remover a senha de um usuario, eh so deixar em branco e confirmar" -ForegroundColor Yellow

    $username = Read-Host "`nDigite o nome do usuario para alterar a senha"
    $userExists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
    if (-not $userExists) {
        Write-Host "O usuario '$username' nao existe." -ForegroundColor Yellow
        Write-Host "`nRetornando ao menu anterior ..."
        Start-Sleep -Seconds 3
        Manage-LocalUsers
        return
    }

    do {
        $newPassword = Read-Host "Digite a nova senha" -AsSecureString
        $passwordConfirmation = Read-Host "Confirme a nova senha" -AsSecureString

        $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringUni([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPassword))
        $plainConfirmation = [System.Runtime.InteropServices.Marshal]::PtrToStringUni([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordConfirmation))

        if ($plainPassword -ne $plainConfirmation) {
            Write-Host "`nAs senhas digitadas nao coincidem. Tente novamente." -ForegroundColor Red
        } else {
            Write-Host "`nSenha confirmada com sucesso!" -ForegroundColor Green
            break
        }
    } while ($true)

    Set-LocalUser -Name $username -Password $newPassword
    Write-Host "`nSenha do usuario '$username' alterada com sucesso!" -ForegroundColor Green
    pause
    Manage-LocalUsers
}

function Enable-Disable-Remove-LocalUser {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "    Habilitar, Desabilitar ou Remover um Usuario Local" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Output "`nSelecione um dos usuarios abaixo:`n"

    $allUsers = Get-LocalUser
    foreach ($user in $allUsers) {
        $status = if ($user.Enabled) { "Habilitado" } else { "Desabilitado" }
        Write-Host "$($user.Name) ... $status"
    }

    $username = Read-Host "`nDigite o nome do usuario para habilitar, desabilitar ou remover"
    $userExists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
    if (-not $userExists) {
        Write-Host "`nO usuario '$username' nao existe." -ForegroundColor Yellow
        Write-Host "`nRetornando ao menu anterior ..."
        Start-Sleep -Seconds 2
        Manage-LocalUsers
        return
    }

    $choice = Read-Host "`nPressione 'H' para Habilitar, 'D' para Desabilitar ou 'R' para Remover"
    switch ($choice.ToUpper()) {
        "H" {
            Enable-LocalUser -Name "$username"
            Write-Host "`nO usuario '$username' foi Habilitado com sucesso!`n" -ForegroundColor Green
            pause
            Manage-LocalUsers
        }
        "D" {
            Disable-LocalUser -Name "$username"
            Write-Host "`nO usuario '$username' foi Desabilitado com sucesso!`n" -ForegroundColor Green
            pause
            Manage-LocalUsers
        }
        "R" {
            Write-Host "`nATENCAO: Remover um usuario e uma acao irreversivel." -ForegroundColor Red
            Write-Host "Isso pode causar perda de informacoes associadas ao usuario '$username'."
            $confirmRemove = Read-Host "`nTem certeza que deseja remover o usuario '$username'? Digite 'Sim' para confirmar ou pressione Enter para cancelar"
            if ($confirmRemove -ne "Sim") {
                Write-Host "`nOperacao cancelada. O usuario '$username' NAO foi removido." -ForegroundColor Yellow
                pause
                Manage-LocalUsers
                return
            }

            Remove-LocalUser -Name "$username"
            Write-Host "`nO usuario '$username' foi removido com sucesso!" -ForegroundColor Green
            pause
            Manage-LocalUsers
        }
        default {
            Write-Host "`nOpcao invalida." -ForegroundColor Yellow
            Write-Host "`nRetornando ao menu anterior ..."
            Start-Sleep -Seconds 3
            Manage-LocalUsers
        }
    }
}

function Add-Remove-UserFromAdminGroup {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "Adicionar ou Remover Usuario do Grupo Administradores" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    $domain = (Get-WmiObject Win32_ComputerSystem).Domain
    $isDomainJoined = $domain -ne (Get-WmiObject Win32_ComputerSystem).Name

    if ($isDomainJoined) {
        Write-Host "`nO computador esta no dominio: $domain" -ForegroundColor Yellow
        Write-Host "`nListando os Usuarios no Grupo Administradores'`n" -ForegroundColor Green
        cmd /c "net localgroup Administradores"
    } else {
        Write-Host "`nO computador esta em um grupo de trabalho: $domain" -ForegroundColor Yellow
        try {
            $admins = Get-LocalGroupMember -Group "Administradores"

            if ($admins) {
                Write-Host "`nMembros do grupo Administradores:" -ForegroundColor Green
                foreach ($admin in $admins) {
                    Write-Host $admin.Name -ForegroundColor Cyan
                }
            } else {
                Write-Host "`nNenhum usuario administrador encontrado." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "`nErro ao acessar os membros do grupo Administradores." -ForegroundColor Red
        }
    }

    $username = Read-Host "`nDigite o nome do usuario"

    if (-not $isDomainJoined) {
        $userExists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
        if (-not $userExists) {
            Write-Host "`nO usuario '$username' nao existe.`n" -ForegroundColor Red
            pause
            return
        }
    }

    $hostname = $env:COMPUTERNAME
    $choice = Read-Host "Pressione 'A' para Adicionar ou 'R' para Remover"
    switch ($choice.ToUpper()) {
        "A" {
            if ($isDomainJoined) {
                cmd /c "net localgroup Administradores $username /add"
            } else {
                Add-LocalGroupMember -Group "Administradores" -Member "$hostname\$username"
            }
            Write-Host "Operacao concluida." -ForegroundColor Green
            pause
            Manage-LocalUsers
        }
        "R" {
            if ($isDomainJoined) {
                cmd /c "net localgroup Administradores $username /delete"
            } else {
                Remove-LocalGroupMember -Group "Administradores" -Member "$hostname\$username"
            }
            Write-Host "Operacao concluida." -ForegroundColor Green
            pause
            Manage-LocalUsers
        }
        default {
            Write-Host "Opcao invalida." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Manage-LocalUsers
        }
    }
}

function Manage-LocalUsers {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "              Gerenciador de Usuarios Locais" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Green

    Write-Host "1) Mostrar os Usuarios locais do Windows que estao ativos"
    Write-Host "2) Mostrar os usuarios locais do grupo administradores"
    Write-Host "3) Criar um novo usuario local"
    Write-Host "4) Criar um novo usuario HTI como Administrador"
    Write-Host "5) Trocar a senha de um usuario local"
    Write-Host "6) Habilitar, Desabilitar ou Remover um usuario local"
    Write-Host "7) Adicionar ou remover um usuario do grupo Administradores"
    Write-Host ""
    Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen

    $userOption = Read-Host "`nEscolha uma opcao"

    switch ($userOption) {
        "1" { Show-ActiveLocalUsers }
        "2" { Show-LocalAdminUsers }
        "3" { Create-NewLocalUser }
        "4" { Create-HtiUser }
        "5" { Change-LocalUserPassword }
        "6" { Enable-Disable-Remove-LocalUser }
        "7" { Add-Remove-UserFromAdminGroup }
        "0" { Clear-Host; return }
        default { Write-Host "Opcao invalida. Por favor, escolha novamente." }
    }
}

# =========================
# DISCO / MONITORAMENTO
# =========================

function Check-SSDHealth {
    Clear-Host
    try {
        $physicalDisks = Get-PhysicalDisk
        foreach ($disk in $physicalDisks) {
            Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host "Informacoes do Disco Fisico: $($disk.DeviceID)" -ForegroundColor Green
            Write-Host "Modelo: $($disk.Model)" -ForegroundColor Green
            Write-Host "Saude Operacional: $($disk.HealthStatus)" -ForegroundColor Green
            Write-Host "Status Operacional: $($disk.OperationalStatus)" -ForegroundColor Green
            Write-Host "Tipo de Midia: $($disk.MediaType)" -ForegroundColor Green
            Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
        }
    } catch {
        Write-Host "Nao foi possivel obter informacoes sobre os discos fisicos: $($_.Exception.Message)" -ForegroundColor Red
    }
    pause
    Disk-ManagementMenu
}

function Check-DiskUsage {
    param (
        [string]$DriveLetter = "C:",
        [int]$Threshold = 0
    )

    try {
        $drive = Get-PSDrive -Name $DriveLetter.TrimEnd(':') -ErrorAction Stop
        $computerName = $env:COMPUTERNAME
        $ipAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike '*Loopback*' }).IPAddress
        $currentUser = $env:USERNAME

        $usedSpacePercent = [math]::Round((($drive.Used / ($drive.Free + $drive.Used)) * 100), 2)

        Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Informacoes sobre a unidade: $DriveLetter" -ForegroundColor Green
        Write-Host "Nome do Computador: $computerName" -ForegroundColor Green
        Write-Host "Usuario Atual: $currentUser" -ForegroundColor Green
        Write-Host "Endereco IP: $(Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 | Select-Object -ExpandProperty IPV4Address)" -ForegroundColor Green
        Write-Host "`nPercentual de Uso: ${usedSpacePercent}%" -ForegroundColor Green
        Write-Host "`nEspaco Total: $([math]::Round(($drive.Free + $drive.Used) / 1GB, 2)) GB" -ForegroundColor Green
        Write-Host "Espaco Livre: $([math]::Round($drive.Free / 1GB, 2)) GB" -ForegroundColor Green
        Write-Host "Espaco Usado: $([math]::Round($drive.Used / 1GB, 2)) GB" -ForegroundColor Green
        Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan

        if ($usedSpacePercent -ge $Threshold) {
            Write-Host "ALERTA: O disco $DriveLetter esta com ${usedSpacePercent}% de uso, acima do limite de ${Threshold}%!" -ForegroundColor Red
        }

        $sendEmail = Read-Host "Deseja enviar essas informacoes por e-mail? (S/N - padrao: N)"
        if ([string]::IsNullOrWhiteSpace($sendEmail) -or $sendEmail -match "^[Nn]$") {
            Write-Host "Relatorio nao sera enviado por e-mail." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Disk-ManagementMenu
        } elseif ($sendEmail -match "^[Ss]$") {
            $subject = "Relatorio de Uso do Disco - $DriveLetter ($computerName)"
            $body = @"
Relatorio de Uso do Disco - Unidade $DriveLetter

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
            Send-Email -Subject $subject -Body $body
        } else {
            Write-Host "Relatorio nao sera enviado por e-mail." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Disk-ManagementMenu
        }

    } catch {
        Write-Host "A unidade '$DriveLetter' nao foi encontrada." -ForegroundColor Red
    }
}

function Select-Drive {
    Write-Host "`nUnidades disponiveis no sistema:" -ForegroundColor Cyan

    $drives = Get-PSDrive | Where-Object { $_.Provider -like "*FileSystem*" }

    if ($drives.Count -eq 0) {
        Write-Host "Nenhuma unidade de disco foi encontrada no sistema." -ForegroundColor Red
        return $null
    }

    $drives | ForEach-Object {
        Write-Host " - $($_.Name):\ ($([math]::Round($_.Used / 1GB, 2)) GB usados / $([math]::Round($_.Free / 1GB, 2)) GB livres)" -ForegroundColor Green
    }

    $defaultDrive = "C"
    $driveLetter = Read-Host "`nDigite a letra da unidade que deseja verificar (Exemplo: C, D) [Default: $defaultDrive]"

    if ([string]::IsNullOrWhiteSpace($driveLetter)) {
        $driveLetter = $defaultDrive
    }

    $drive = $drives | Where-Object { $_.Name -eq $driveLetter }

    if ($drive) {
        Write-Host "Unidade encontrada: $($drive.Name)" -ForegroundColor Green
    } else {
        Write-Host "Unidade '$driveLetter' nao foi encontrada." -ForegroundColor Red
        return $null
    }

    return $driveLetter
}

function Schedule-DiskMonitoring {
    try {
        $scriptPath = Join-Path $ScriptsPath "MonitorDisk.ps1"

        if (-not (Test-Path $scriptPath)) {
            Write-Host "`nScript MonitorDisk.ps1 nao encontrado. Baixando do GitHub..." -ForegroundColor Yellow
            try {
                Download-FromGit -RelativePath "Scripts/MonitorDisk.ps1" -DestinationPath $scriptPath
                Write-Host "Script MonitorDisk.ps1 baixado e salvo em '$scriptPath'." -ForegroundColor Green
            } catch {
                Write-Host "Falha ao baixar o script MonitorDisk.ps1: $($_.Exception.Message)" -ForegroundColor Red
                return
            }
        }

        Write-Host "`nDeseja criar um agendamento para monitorar o uso do HD? [S/n]" -ForegroundColor Yellow
        $confirm = Read-Host "(Pressione Enter para SIM)"
        if ($confirm -and $confirm -notmatch "^[Ss]$") {
            Write-Host "`nOperacao cancelada pelo usuario. Retornando ao menu principal..." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Disk-ManagementMenu
            return
        }

        Write-Host "`nEscolha a frequencia do agendamento:"
        Write-Host "1) Diario"
        Write-Host "2) Semanal"
        Write-Host "3) Mensal"
        $frequencyChoice = Read-Host "`nDigite o numero da opcao desejada"

        $frequency = switch ($frequencyChoice) {
            "1" { "DAILY" }
            "2" { "WEEKLY" }
            "3" { "MONTHLY" }
            default {
                Write-Host "Opcao invalida. Tente novamente" -ForegroundColor Red
                Start-Sleep -Seconds 2
                Disk-ManagementMenu
                return
            }
        }

        Write-Host "`nATENCAO" -ForegroundColor Yellow
        Write-Host "O monitoramento esta pre-configurado para o disco C: com limite de 90%." -ForegroundColor Cyan
        Write-Host "Se deseja alterar esses valores, edite diretamente o script em $scriptPath." -ForegroundColor Cyan

        $taskName = "Monitoramento de Disco"
        $taskFolder = "_HTI"
        $folderPath = "\$taskFolder"

        $service = New-Object -ComObject "Schedule.Service"
        $service.Connect()
        $rootFolder = $service.GetFolder("\")
        try {
            $rootFolder.GetFolder($folderPath) | Out-Null
        } catch {
            $rootFolder.CreateFolder($folderPath) | Out-Null
        }

        $taskPath = "\$taskFolder\$taskName"

        if (Get-ScheduledTask -TaskName $taskName -TaskPath "\$taskFolder\" -ErrorAction SilentlyContinue) {
            Write-Host "`nJa existe uma tarefa chamada '$taskName' na pasta '$taskFolder'." -ForegroundColor Yellow
            $overwrite = Read-Host "Deseja substituir a tarefa existente? (S/N)"
            if ($overwrite -notmatch "^[Ss]$") {
                Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Disk-ManagementMenu
                return
            }

            Unregister-ScheduledTask -TaskName $taskName -TaskPath "\$taskFolder\" -Confirm:$false
        }

        switch ($frequency) {
            "DAILY" {
                schtasks.exe /Create /F /TN "$taskPath" /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" /SC DAILY /ST 12:00 /RL HIGHEST /RU "SYSTEM"
            }
            "WEEKLY" {
                schtasks.exe /Create /F /TN "$taskPath" /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" /SC WEEKLY /D MON /ST 12:00 /RL HIGHEST /RU "SYSTEM"
            }
            "MONTHLY" {
                schtasks.exe /Create /F /TN "$taskPath" /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" /SC MONTHLY /D MON /MO FIRST /ST 12:00 /RL HIGHEST /RU "SYSTEM"
            }
        }

        Write-Host "`nTarefa '$taskName' criada com sucesso na pasta '$taskFolder' no Agendador de Tarefas do Windows!" -ForegroundColor Green
    } catch {
        Write-Host "`nErro ao criar o agendamento: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host ""
    pause
    Disk-ManagementMenu
}

function Disk-ManagementMenu {
    Clear-Host
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "           Status, Monitoramento e Alertas de Disco           " -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "Escolha uma opcao:`n" -ForegroundColor Yellow
    Write-Host "1) Verificar a saude dos discos" -ForegroundColor Cyan
    Write-Host "2) Mostrar o Uso do HD e enviar por e-mail" -ForegroundColor Cyan
    Write-Host "3) Criar um agendamento no Windows para monitorar o uso do HD" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    $choice = Read-Host "Digite o numero da opcao desejada"

    switch ($choice) {
        "1" { Check-SSDHealth }
        "2" {
            $driveLetter = Select-Drive
            if ($driveLetter -eq $null) {
                Write-Host "Operacao cancelada." -ForegroundColor Yellow
                return
            }
            $threshold = 90
            Check-DiskUsage -DriveLetter $driveLetter -Threshold $threshold
        }
        "3" { Schedule-DiskMonitoring }
        "0" { Clear-Host; return }
        default { Write-Host "Opcao invalida." -ForegroundColor Red }
    }
}

# =========================
# OTIMIZACOES
# =========================

function PowerUP {
    $monitorTimeoutDC = 5
    $standbyTimeoutDC = 10
    $monitorTimeoutAC = 60
    $standbyTimeoutAC = 0

    $isLaptop = (Get-WmiObject -Class Win32_SystemEnclosure).ChassisTypes -match "8|9|10|11|14|18|21|30"

    if ($isLaptop) {
        Write-Output "Dispositivo identificado como Notebook."
        Start-Process -FilePath "powercfg" -ArgumentList "/change monitor-timeout-dc $monitorTimeoutDC" -Verb RunAs -Wait
        Start-Process -FilePath "powercfg" -ArgumentList "/change standby-timeout-dc $standbyTimeoutDC" -Verb RunAs -Wait
        Start-Process -FilePath "powercfg" -ArgumentList "/change monitor-timeout-ac $monitorTimeoutAC" -Verb RunAs -Wait
        Start-Process -FilePath "powercfg" -ArgumentList "/change standby-timeout-ac $standbyTimeoutAC" -Verb RunAs -Wait
    } else {
        Write-Output "Dispositivo identificado como Desktop."
        Start-Process -FilePath "powercfg" -ArgumentList "/change monitor-timeout-ac $monitorTimeoutAC" -Verb RunAs -Wait
        Start-Process -FilePath "powercfg" -ArgumentList "/change standby-timeout-ac $standbyTimeoutAC" -Verb RunAs -Wait
    }

    Write-Output "Configuracoes de energia aplicadas com sucesso!"
    pause
    Otimizacoes-Win
}

function WinUpdate {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "           Verificando atualizacoes do Windows..." -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan

    $updates = Get-CimInstance -Namespace "root\ccm\clientSDK" -ClassName "CCM_SoftwareUpdate" -ErrorAction SilentlyContinue

    if ($updates.Count -eq 0) {
        Write-Host "Nenhuma atualizacao disponivel no momento." -ForegroundColor Green
    } else {
        Write-Host "Atualizacoes disponiveis:" -ForegroundColor Yellow
        $updates | ForEach-Object { Write-Host "- $($_.ArticleID)" -ForegroundColor White }

        $install = Read-Host "`nDeseja instalar todas as atualizacoes agora? (S/N)"
        if ($install -match "[Ss]") {
            Write-Host "`nIniciando a instalacao das atualizacoes..." -ForegroundColor Cyan
            Start-Process -FilePath "C:\Windows\System32\UsoClient.exe" -ArgumentList "StartScan" -Wait
            Write-Host "`nAs atualizacoes foram aplicadas com sucesso!" -ForegroundColor Green
            Write-Host "Reinicie o computador para concluir a instalacao, se necessario." -ForegroundColor Red
        } else {
            Write-Host "`nAs atualizacoes nao foram instaladas." -ForegroundColor Yellow
        }
    }

    pause
    Otimizacoes-Win
}

function Clear-PrintSpoolerQueue {
    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host "INICIANDO LIMPEZA DA FILA DE IMPRESSAO" -ForegroundColor Yellow
    Write-Host "==========================================================="

    try {
        Write-Host "`n[ETAPA 1] Parando o servico de Spooler (Spooler)..." -ForegroundColor Cyan
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Write-Host "Servico de Spooler parado com sucesso." -ForegroundColor Green

        $spoolPath = Join-Path -Path $env:SystemRoot -ChildPath "System32\spool\PRINTERS"
        Write-Host "`n[ETAPA 2] Limpando arquivos da fila em '$spoolPath'..." -ForegroundColor Cyan
        Remove-Item -Path "$spoolPath\*" -Recurse -Force -ErrorAction Stop
        Write-Host "Fila de impressao limpa com sucesso." -ForegroundColor Green
    }
    catch {
        Write-Host "Ocorreu um erro durante a limpeza da fila de impressao: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        Write-Host "`n[ETAPA 3] Iniciando o servico de Spooler..." -ForegroundColor Cyan
        $spoolerService = Get-Service -Name Spooler
        if ($spoolerService.Status -ne 'Running') {
            Start-Service -Name Spooler
            Write-Host "Servico de Spooler iniciado com sucesso." -ForegroundColor Green
        } else {
            Write-Host "O servico de Spooler ja estava em execucao." -ForegroundColor Yellow
        }

        Write-Host "`n===========================================================" -ForegroundColor Yellow
        Write-Host "PROCESSO DE LIMPEZA CONCLUIDO." -ForegroundColor Green
        Write-Host "==========================================================="
    }
}

function Limpar-CacheTeams {
    $confirmacao = Read-Host "`nEsta acao vai limpar o cache do Teams e desconectar a conta. Deseja continuar? (S/N) [S]"

    if ($confirmacao -eq "N" -or $confirmacao -eq "n") {
        Write-Host "Operacao cancelada pelo usuario." -ForegroundColor Yellow
        pause
        Otimizacoes-Win
    }

    Write-Host "`nIniciando limpeza de cache do Microsoft Teams..." -ForegroundColor Cyan

    $processNames = @("Teams", "ms-teams")
    foreach ($proc in $processNames) {
        $teamsProcess = Get-Process -Name $proc -ErrorAction SilentlyContinue
        if ($teamsProcess) {
            Write-Host "Encerrando o processo: $proc ..." -ForegroundColor Yellow
            Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
        }
    }
    Start-Sleep -Seconds 2

    $legacyPath = "$env:APPDATA\Microsoft\Teams"
    $newTeamsPath = "$env:LocalAppData\Packages\MSTeams_8wekyb3d8bbwe\LocalCache"
    $newTeamsAltPath = "$env:LocalAppData\Microsoft\Teams"

    $found = $false

    if (Test-Path $newTeamsPath) {
        Remove-Item "$newTeamsPath\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Cache da nova versao limpo com sucesso." -ForegroundColor Green
        $found = $true
    }
    elseif (Test-Path $newTeamsAltPath) {
        Remove-Item "$newTeamsAltPath\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Cache da nova versao (alternativa) limpo com sucesso." -ForegroundColor Green
        $found = $true
    }
    elseif (Test-Path $legacyPath) {
        $foldersToDelete = @(
            "Application Cache\Cache",
            "blob_storage",
            "Cache",
            "databases",
            "GPUCache",
            "IndexedDB",
            "Local Storage",
            "tmp"
        )

        foreach ($folder in $foldersToDelete) {
            $fullPath = Join-Path $legacyPath $folder
            if (Test-Path $fullPath) {
                Remove-Item $fullPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        Write-Host "Cache da versao classica limpo com sucesso." -ForegroundColor Green
        $found = $true
    }

    if (-not $found) {
        Write-Host "Nao foi possivel localizar os arquivos de cache do Teams." -ForegroundColor Red
    } else {
        $teamsPaths = @(
            "$env:LocalAppData\Microsoft\WindowsApps\ms-teams.exe",
            "$env:LocalAppData\Microsoft\Teams\current\Teams.exe"
        )

        $started = $false
        foreach ($path in $teamsPaths) {
            if (Test-Path $path) {
                Start-Process $path
                Write-Host "Teams iniciado com sucesso a partir de: $path" -ForegroundColor Green
                $started = $true
                break
            }
        }

        if (-not $started) {
            Write-Host "Nao foi possivel iniciar o Teams. Inicie manualmente, se necessario." -ForegroundColor Red
        }
    }

    Write-Host "`nLimpeza concluida!" -ForegroundColor Green
    pause
    Otimizacoes-Win
}

function Gerenciar-AnyDesk {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "           Verificacao do AnyDesk ..." -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
    Write-Host "`nATENCAO: Tanto para Instalar ou Remover sera aberta janela do Anydesk para interacao`n" -ForegroundColor Yellow

    $servico = Get-Service -Name "AnyDesk" -ErrorAction SilentlyContinue
    $processo = Get-Process -Name "AnyDesk" -ErrorAction SilentlyContinue

    $uninstalls = @()
    $uninstalls += Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
    $uninstalls += Get-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue

    $anydeskPath = ($uninstalls | Where-Object { $_.DisplayName -like "*AnyDesk*" }).InstallLocation
    if (-not $anydeskPath -or -not (Test-Path $anydeskPath)) {
        $anydeskPath = "C:\Program Files (x86)\AnyDesk"
    }

    $versao = $null
    if (Test-Path (Join-Path $anydeskPath "AnyDesk.exe")) {
        try { $versao = (Get-Item (Join-Path $anydeskPath "AnyDesk.exe")).VersionInfo.ProductVersion.Trim() } catch {}
    }

    if (-not $versao -and $processo) {
        try { $versao = ($processo | Select-Object -First 1).FileVersion.Trim() } catch {}
    }

    if ($servico -or $processo) {
        Write-Host "`nAnyDesk encontrado no sistema!" -ForegroundColor Green
        if ($servico) { Write-Host "Status do Servico: $($servico.Status)" -ForegroundColor Green }
        if ($versao) { Write-Host "Versao do AnyDesk: $versao" -ForegroundColor Green }
        else { Write-Host "Versao do AnyDesk: Nao foi possivel obter" -ForegroundColor Yellow }
    } else {
        Write-Host "`nAnyDesk nao encontrado no sistema." -ForegroundColor Yellow
    }

    Write-Host "`nSelecione uma opcao:"
    Write-Host "`n1) Instala AnyDesk (ultima versao via Winget)"
    Write-Host "2) Baixa e executa AnyDesk versao 7 do GitHub"
    Write-Host "3) Remover AnyDesk do computador"
    Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    $opcao = Read-Host "`nDigite o numero da opcao desejada"

    switch ($opcao) {
        "1" {
            Write-Host "`nInstalando ultima versao do AnyDesk via Winget..." -ForegroundColor Cyan
            winget install AnyDesk -e --silent
            pause
            Otimizacoes-Win
        }
        "2" {
            $arquivo = Join-Path $UtilsPath "AnyDesk-v7.exe"

            Write-Host "`nBaixando o instalador do AnyDesk-v7..." -ForegroundColor Cyan
            try {
                Download-FromGit -RelativePath "Utilitarios/AnyDesk-v7.exe" -DestinationPath $arquivo
                Write-Host "Instalador do AnyDesk-v7 baixado com sucesso para '$arquivo'." -ForegroundColor Green
            } catch {
                Write-Host "Erro ao baixar o instalador do AnyDesk-v7: $($_.Exception.Message)" -ForegroundColor Red
                pause
                Gerenciar-AnyDesk
            }

            Write-Host "`nInstalando o AnyDesk v7..." -ForegroundColor Cyan
            try {
                $argumentos = "/S /U=1"
                Start-Process -FilePath $arquivo -ArgumentList $argumentos -Wait -NoNewWindow
                Write-Host "AnyDesk instalado com sucesso!" -ForegroundColor Green
                pause
                Gerenciar-AnyDesk
            } catch {
                Write-Host "Erro ao instalar o AnyDesk: $($_.Exception.Message)" -ForegroundColor Red
                pause
                Gerenciar-AnyDesk
            }
        }
        "3" {
            if ($servico) { Stop-Service -Name "AnyDesk" -Force -ErrorAction SilentlyContinue }
            if ($processo) { Stop-Process -Name "AnyDesk" -Force -ErrorAction SilentlyContinue }

            winget uninstall AnyDesk -e --silent --force
            Remove-Item -Recurse -Force "$env:APPDATA\AnyDesk" -ErrorAction SilentlyContinue
            Remove-Item -Recurse -Force "$env:LOCALAPPDATA\AnyDesk" -ErrorAction SilentlyContinue
            Remove-Item -Recurse -Force "C:\ProgramData\AnyDesk" -ErrorAction SilentlyContinue

            Write-Host "`nRemocao concluida!" -ForegroundColor Green
            pause
            Otimizacoes-Win
        }
        "0" { Clear-Host; Otimizacoes-Win }
        default { Write-Host "`nOpcao invalida. Nenhuma acao realizada." -ForegroundColor Red }
    }
}

function LogsEventViewer {
    Clear-Host
    Write-Host "=================================================="
    Write-Host "     Analisador de Logs do Windows para TI"
    Write-Host "=================================================="
    Write-Host
    Write-Host "Essa rotina foi mantida simplificada nesta versao." -ForegroundColor Yellow
    Write-Host "Abra o Visualizador de Eventos manualmente se precisar de analise detalhada." -ForegroundColor Cyan
    pause
    Clear-Host
    Otimizacoes-Win
}

function CleanTemp {
    Clear-Host
    Write-Host "===================================================" -ForegroundColor Green
    Write-Host " LIMPEZA DE ARQUIVOS TEMPORARIOS" -ForegroundColor Green
    Write-Host "===================================================" -ForegroundColor Green

    try {
        Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Limpeza basica concluida com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro durante a limpeza: $($_.Exception.Message)" -ForegroundColor Red
    }

    pause
    Otimizacoes-Win
}

function SystemIntegrityCheck {
    Clear-Host
    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host "FUNCAO DE VERIFICACAO DE INTEGRIDADE DO SISTEMA" -ForegroundColor Yellow
    Write-Host "==========================================================="

    $title = "Confirmar Execucao"
    $question = "Deseja iniciar o processo completo de verificacao e reparo?"
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim", "&Nao")
    $userResponse = $host.UI.PromptForChoice($title, $question, $choices, 0)

    if ($userResponse -eq 1) {
        Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Yellow
        Clear-Host
        Otimizacoes-Win
    }

    try {
        Write-Host "`n[ETAPA 1 de 2] Executando o Reparo da Imagem do Windows..." -ForegroundColor Cyan
        Repair-WindowsImage -Online -RestoreHealth -ErrorAction Stop
        Write-Host "[SUCESSO] Etapa 1 concluida." -ForegroundColor Green
        Start-Sleep -Seconds 3

        Write-Host "`n[ETAPA 2 de 2] Executando o Verificador de Arquivos do Sistema (SFC)..." -ForegroundColor Cyan
        sfc.exe /scannow

        Write-Host "`n===========================================================" -ForegroundColor Yellow
        Write-Host "VERIFICACAO DE INTEGRIDADE CONCLUIDA!" -ForegroundColor Green
        Write-Host "==========================================================="

    } catch {
        Write-Host "Ocorreu um erro critico durante o processo de verificacao: $($_.Exception.Message)" -ForegroundColor Red
    }

    pause
    Otimizacoes-Win
}

function NetworkStackReset {
    Clear-Host
    Write-Host "================== Reset de Rede ==================" -ForegroundColor Yellow

    $title = "Confirmar Reset da Rede"
    $question = "Prosseguir com a redefinicao de rede?"
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim, continuar", "&Nao, cancelar")
    $finalResponse = $host.UI.PromptForChoice($title, $question, $choices, 1)

    if ($finalResponse -ne 0) {
        Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Yellow
        return
    }

    try {
        Clear-DnsClientCache
        ipconfig /release | Out-Null
        ipconfig /renew | Out-Null
        netsh.exe winsock reset | Out-Null
        netsh.exe int ip reset | Out-Null

        Write-Host "`nPROCESSO DE RESET DE REDE CONCLUIDO COM SUCESSO!" -ForegroundColor Green
    } catch {
        Write-Host "Ocorreu um erro inesperado durante a redefinicao de rede: $($_.Exception.Message)" -ForegroundColor Red
    }

    pause
    Otimizacoes-Win
}

function Otimizacoes-Win {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "           Ferramentas e Otimizacoes do Windows" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Green

    Write-Host "1) Gerenc. Energia - Sempre ativo" -ForegroundColor Green
    Write-Host "2) Atualizacao do Windows" -ForegroundColor Green
    Write-Host "3) Instalar o Microsoft PowerToys - PENDENTE" -ForegroundColor DarkGray
    Write-Host "4) Atualizacao dos Drivers - PENDENTE" -ForegroundColor DarkGray
    Write-Host "5) Limpeza Fila de Impressao" -ForegroundColor Green
    Write-Host "6) Atualiza todos os programas instalados pelo Winget - PENDENTE" -ForegroundColor DarkGray
    Write-Host "7) Limpeza de arquivos temporarios" -ForegroundColor Green
    Write-Host "8) Verificacao do disco - PENDENTE" -ForegroundColor DarkGray
    Write-Host "9) Checagem da integridade do sistema" -ForegroundColor Green
    Write-Host "10) Reset de Rede" -ForegroundColor Green
    Write-Host "11) Visualizador Logs EventViewer" -ForegroundColor Green
    Write-Host "12) Todas as opcoes acima - PENDENTE" -ForegroundColor DarkGray
    Write-Host "13) Limpeza inteligente de cache do Teams" -ForegroundColor Green
    Write-Host "14) Anydesk: Status / Instala / Remover" -ForegroundColor Green
    Write-Host ""
    Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen

    $userOption = Read-Host "`nEscolha uma opcao"

    switch ($userOption) {
        "1" { PowerUP }
        "2" { WinUpdate }
        "3" { InstallPowerToys }
        "4" { UpdateDrivers }
        "5" { Clear-PrintSpoolerQueue }
        "6" { UpdatePrograms }
        "7" { CleanTemp }
        "8" { CheckDisk }
        "9" { SystemIntegrityCheck }
        "10" { NetworkStackReset }
        "11" { LogsEventViewer }
        "12" { TodasFuncoes }
        "13" { Limpar-CacheTeams }
        "14" { Gerenciar-AnyDesk }
        "0" { Clear-Host; Main }
        default { Write-Host "Opcao invalida. Por favor, escolha novamente." }
    }
}

# =========================
# PENDENTES / PLACEHOLDERS
# =========================

function Install-Bitdefender { Write-Host "Funcao pendente." -ForegroundColor Yellow; pause; Main }
function Install-LegalPackage { Write-Host "Funcao pendente." -ForegroundColor Yellow; pause; Main }
function Install-AccountingPackage { Write-Host "Funcao pendente." -ForegroundColor Yellow; pause; Main }
function InstallPowerToys { Write-Host "Funcao pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
function UpdateDrivers { Write-Host "Funcao pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
function UpdatePrograms { Write-Host "Funcao pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
function CheckDisk { Write-Host "Funcao pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
function TodasFuncoes { Write-Host "Funcao pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }

# =========================
# LIMPEZA / CHANGELOG / MAIN
# =========================

function Remove-HTIFiles {
    $files = @(
        (Join-Path $ScriptsPath "Script-Inicial.ps1"),
        (Join-Path $ScriptsPath "clientes.csv")
    )

    foreach ($file in $files) {
        if (Test-Path $file) {
            Remove-Item -Path $file -Force -ErrorAction SilentlyContinue
        }
    }
}

function Show-Changelog {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "                  HISTORICO DE ALTERACOES" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
    Write-Host "[v0.8.0] - Migrado para GitHub raw nas rotinas de download"
    Write-Host "[v0.7.5] - Adicionado reset de rede do Windows"
    Write-Host "[v0.7.4] - Adicionada checagem de integridade do sistema"
    Write-Host "[v0.7.3] - Adicionada limpeza da fila de impressao"
    Write-Host "[v0.7.2] - Adicionada analise do WinSAT"
    Write-Host "[v0.7.1] - Melhorias gerais"
    Write-Host "`nPressione qualquer tecla para voltar ao menu principal..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Main
}

function Main {
    Clear-Host
    Initialize-HTIStructure
    Remove-HTIFiles

    if (-not (Test-Admin)) {
        Write-Host "ATENCAO: Este script precisa ser executado como administrador.`n" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para sair..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        return
    }

    do {
        Write-Host "`n--------------------------------------------------------------" -ForegroundColor Blue
        Write-Host "             SETUP-01 - HORUS TI SOLUCOES" -ForegroundColor Cyan
        Write-Host "--------------------------------------------------------------`n" -ForegroundColor Blue
        Write-Host "Versao: 0.8.0`n" -ForegroundColor DarkGray
        Write-Host "1) Exibir Informacoes do Windows" -ForegroundColor Green
        Write-Host "2) Renomear o computador" -ForegroundColor Green
        Write-Host "3) Instalacao do agente TiFlux" -ForegroundColor Green
        Write-Host "4) Instalacao do Endpoint Bitdefender - PENDENTE" -ForegroundColor DarkGray
        Write-Host "5) Instalacao do pacote basico de programas" -ForegroundColor Green
        Write-Host "6) Instalacao do pacote de programas juridicos - PENDENTE" -ForegroundColor DarkGray
        Write-Host "7) Instalacao do pacote de programas contabeis - PENDENTE" -ForegroundColor DarkGray
        Write-Host "8) Instalacao do Office" -ForegroundColor Green
        Write-Host "9) Ativacao do Windows e Office (oficial)" -ForegroundColor Green
        Write-Host "10) Remove Apps desnecessarios" -ForegroundColor Green
        Write-Host "11) Baixar aplicativos utilitarios {DiskUse, BSOD, Sysinternals...}" -ForegroundColor Green
        Write-Host "12) Gerenciador de Usuarios Locais" -ForegroundColor Green
        Write-Host "13) Ferramentas e Otimizacoes" -ForegroundColor Green
        Write-Host "14) Monitorar uso do HD e envio de alertas" -ForegroundColor Green
        Write-Host ""
        Write-Host "0) Sair - Limpa historico"

        $mainOption = Read-Host "`nEscolha uma opcao"

        switch ($mainOption) {
            "1"  { Write-Host "Aguarde, coletando informacoes do Windows..." -ForegroundColor Yellow; Info-Windows }
            "2"  { Rename-ComputerCustom }
            "3"  { Install-Agente }
            "4"  { Install-Bitdefender }
            "5"  { Install-BasicPackage }
            "6"  { Install-LegalPackage }
            "7"  { Install-AccountingPackage }
            "8"  { Install-Office }
            "9"  { Activate-WindowsOffice }
            "10" { Remove-Apps }
            "11" { Download-Utilities }
            "12" { Manage-LocalUsers }
            "13" { Otimizacoes-Win }
            "14" { Disk-ManagementMenu }
            "99" { Show-Changelog }
            "0"  {
                Write-Host "Limpando historico powershell..." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Remove-HTIFiles
                try {
                    [System.IO.File]::WriteAllText((Get-PSReadLineOption).HistorySavePath, "")
                } catch {}
                Clear-History -ErrorAction SilentlyContinue
                Clear-Host
                return
            }
            default { Write-Host "Opcao invalida. Tente novamente." }
        }
    } while ($mainOption -ne "0")
}

Main
