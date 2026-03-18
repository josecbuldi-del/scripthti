# ==============================================================================
# SETUP-01 - HORUS TI SOLUCOES
# Versao: 0.0.3
# Requer: PowerShell como Administrador
# Antes de executar: Set-ExecutionPolicy Bypass -Scope Process -Force
# ==============================================================================

Clear-Host

# ==============================================================================
# CONFIGURACAO CENTRAL
# Altere apenas esta secao para ajustar caminhos e URLs base
# ==============================================================================
$GitBaseUrl   = "https://raw.githubusercontent.com/josecbuldi-del/scripthti/main"
$BasePath     = "C:\_HTI"
$ScriptsPath  = Join-Path $BasePath "Scripts"
$UtilsPath    = Join-Path $BasePath "Utilitarios"
$ProgramsPath = Join-Path $BasePath "Programas"
$CleanupPath  = Join-Path $ScriptsPath "Limpeza"

# ==============================================================================
# SENHA DE ACESSO
# Altere $ScriptPassword para trocar a senha de acesso ao script
# ==============================================================================
$ScriptPassword = "suroh"

function Request-AccessPassword {
    Clear-Host
    Write-Host ""
    Write-Host "  ██╗  ██╗████████╗██╗" -ForegroundColor DarkBlue
    Write-Host "  ██║  ██║╚══██╔══╝██║" -ForegroundColor DarkBlue
    Write-Host "  ███████║   ██║   ██║" -ForegroundColor DarkBlue
    Write-Host "  ██╔══██║   ██║   ██║" -ForegroundColor DarkBlue
    Write-Host "  ██║  ██║   ██║   ██║" -ForegroundColor DarkBlue
    Write-Host "  ╚═╝  ╚═╝   ╚═╝   ╚═╝" -ForegroundColor DarkBlue
    Write-Host ""
    Write-Host "--------------------------------------------------------------" -ForegroundColor Blue
    Write-Host "             HORUS TI SOLUCOES" -ForegroundColor White
    Write-Host "--------------------------------------------------------------" -ForegroundColor Blue
    Write-Host ""

    $attempts = 0
    do {
        $attempts++
        $inputPassword = Read-Host "Digite a senha de acesso"
        if ($inputPassword -eq $ScriptPassword) {
            Write-Host "`nAcesso autorizado!" -ForegroundColor Green
            Start-Sleep -Seconds 1
            return $true
        } else {
            Write-Host "Senha incorreta. Tente novamente." -ForegroundColor Red
            if ($attempts -ge 3) {
                Write-Host "`nNumero maximo de tentativas atingido. Encerrando." -ForegroundColor Red
                Start-Sleep -Seconds 2
                exit
            }
        }
    } while ($true)
}

# ==============================================================================
# INICIALIZACAO DE ESTRUTURA
# ==============================================================================
function Initialize-HTIStructure {
    @($BasePath, $ScriptsPath, $UtilsPath, $ProgramsPath, $CleanupPath) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -Path $_ -ItemType Directory -Force | Out-Null
        }
    }
}

function Get-GitRawUrl {
    param([Parameter(Mandatory=$true)][string]$RelativePath)
    return "$GitBaseUrl/$RelativePath"
}

function Download-FromGit {
    param(
        [Parameter(Mandatory=$true)][string]$RelativePath,
        [Parameter(Mandatory=$true)][string]$DestinationPath
    )
    $url = Get-GitRawUrl -RelativePath $RelativePath
    Invoke-WebRequest -Uri $url -OutFile $DestinationPath -UseBasicParsing -ErrorAction Stop
}

# ==============================================================================
# FUNCOES UTILITARIAS GERAIS
# ==============================================================================
function Check-WingetInstalled {
    $wingetPackage = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*Microsoft.DesktopAppInstaller*" }
    if ($null -eq $wingetPackage -or -not $wingetPackage.InstallLocation) {
        Write-Host "Winget nao esta instalado ou nao foi encontrado." -ForegroundColor Yellow
        return $false
    }
    $wingetExePath = Join-Path $wingetPackage.InstallLocation "winget.exe"
    if (Test-Path $wingetExePath) {
        $env:Path = $env:Path + ";" + $wingetPackage.InstallLocation
    } else {
        Write-Host "Winget.exe nao encontrado: $wingetExePath" -ForegroundColor Red
        return $false
    }
    try { $null = Get-Command winget -ErrorAction Stop; return $true } catch { return $false }
}

function Show-LoadingAnimation {
    $animation = @("|", "/", "-", "\")
    $i = 0; $count = 0
    while ($count -lt 5) {
        Write-Host -NoNewline "$($animation[$i % $animation.Length])" -ForegroundColor Green
        Start-Sleep -Milliseconds 200
        Write-Host -NoNewline "`r"
        $i++; $count++
    }
    Write-Host ""
}

function Accept-WingetTerms { echo Y | winget list > $null }

function Get-ScoreColor {
    param([double]$Score)
    if ($Score -ge 1 -and $Score -le 3)        { return "Red" }
    elseif ($Score -ge 4 -and $Score -le 5)    { return "DarkYellow" }
    elseif ($Score -ge 5.1 -and $Score -le 6.9){ return "Yellow" }
    elseif ($Score -ge 7 -and $Score -le 8.9)  { return "Green" }
    elseif ($Score -ge 9)                       { return "Blue" }
    else                                        { return "White" }
}

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ==============================================================================
# LIMPEZA COMPLETA DO HISTORICO DO POWERSHELL
# ==============================================================================
function Clear-AllPSHistory {
    try {
        $histPath = (Get-PSReadLineOption).HistorySavePath
        if (Test-Path $histPath) {
            [System.IO.File]::WriteAllText($histPath, "")
        }
    } catch {}
    Clear-History -ErrorAction SilentlyContinue
    [Microsoft.PowerShell.PSConsoleReadLine]::ClearHistory()
}

# ==============================================================================
# ENVIO DE E-MAIL
# ==============================================================================
function Send-Email {
    param(
        [Parameter(Mandatory=$true)][string]$Subject,
        [Parameter(Mandatory=$true)][string]$Body
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
        Send-MailMessage -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $credential `
            -From $sender -To $recipient -Subject $Subject -Body $Body
        Write-Host "E-mail enviado com sucesso!" -ForegroundColor Green
        Start-Sleep -Seconds 2; Clear-Host
    } catch {
        Write-Host "Erro ao enviar o e-mail: $($_.Exception.Message)" -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

# ==============================================================================
# OPCAO 1 - INFORMACOES DO WINDOWS
# ==============================================================================
function Info-Windows {
    Clear-Host
    Write-Host "`n`n`n`n`n`nAguarde, coletando informacoes do computador..." -ForegroundColor Yellow
    Show-LoadingAnimation

    $computerInfo = Get-ComputerInfo
    Clear-Host

    Write-Output "--------------------------------------------------------------"
    Write-Host "             -------   INFORMACOES DO WINDOWS   -------         " -ForegroundColor Blue
    Write-Output "--------------------------------------------------------------`n"
    Write-Output "Data de Execucao: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')"
    Write-Output "Nome do Computador: $env:COMPUTERNAME"
    Write-Output "Versao do Windows: $($computerInfo.OsName)"

    $winVersion = $null
    try {
        $winVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion -EA SilentlyContinue).DisplayVersion
        if (-not $winVersion) {
            $winVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId -EA SilentlyContinue).ReleaseId
        }
    } catch { $winVersion = "Nao disponivel" }

    Write-Output "Release: $winVersion"
    Write-Output "Data Instalacao: $($computerInfo.OsInstallDate)"

    $psVersion = $PSVersionTable.PSVersion
    Write-Output "Versao do PowerShell: $psVersion"
    if ($psVersion.Major -lt 5) {
        Write-Host "ATENCAO: PowerShell desatualizado!" -ForegroundColor Red
    }

    if (Check-WingetInstalled) { Write-Host "Winget: Instalado" -ForegroundColor Green }
    else { Write-Host "Winget: Nao Instalado" -ForegroundColor Red }

    Write-Output "Dominio/Workgroup: $($computerInfo.CsDomain)"
    Write-Output "Fabricante: $($computerInfo.CsManufacturer)"
    Write-Output "Modelo: $($computerInfo.CsModel)"
    Write-Output "Familia: $($computerInfo.CsSystemFamily)"
    Write-Output "Processador: $((Get-WmiObject Win32_Processor).Name)"

    $totalMemoryGB = [math]::Round($computerInfo.OsTotalVisibleMemorySize / 1MB, 2)
    Write-Output "Memoria (GB): $totalMemoryGB"

    $uptime = $computerInfo.OsUptime
    $uptimeFormatted = $uptime.ToString('d\.hh\:mm\:ss')
    if ($uptime.Days -ge 3)     { Write-Host "Uptime: $uptimeFormatted" -ForegroundColor Red }
    elseif ($uptime.Days -ge 1) { Write-Host "Uptime: $uptimeFormatted" -ForegroundColor Yellow }
    else                        { Write-Output "Uptime: $uptimeFormatted" }

    Write-Output "Usuario: $env:USERNAME"

    try {
        $localIP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 -EA Stop).IPV4Address.IPAddressToString
        Write-Output "IP Local: $localIP"
    } catch { $localIP = "Nao disponivel"; Write-Output "IP Local: $localIP" }

    function Get-PublicIP {
        param([string]$url)
        try {
            if (-not ("System.Net.Http.HttpClient" -as [Type])) { Add-Type -AssemblyName "System.Net.Http" }
            $client = New-Object System.Net.Http.HttpClient
            $client.Timeout = New-TimeSpan -Seconds 10
            $task = $client.GetStringAsync($url)
            if ($task.Wait(10000)) { return $task.Result } else { throw "Timeout" }
        } catch { return "Nao disponivel" }
        finally { if ($client) { $client.Dispose() } }
    }

    $publicIPv4 = Get-PublicIP "https://api.ipify.org?format=text"
    $publicIPv6 = Get-PublicIP "https://api64.ipify.org?format=text"

    if ($publicIPv4 -like "*Nao disponivel*" -and $publicIPv6 -like "*Nao disponivel*") {
        Write-Output "IP Publico: Nao foi possivel obter."
    } elseif ($publicIPv4 -eq $publicIPv6 -or $publicIPv6 -like "*Nao disponivel*") {
        Write-Output "IP Publico (IPv4): $publicIPv4"
    } else {
        Write-Output "IP Publico (IPv4): $publicIPv4"
        Write-Output "IP Publico (IPv6): $publicIPv6"
    }

    Write-Output "--------------------------------------------------------------"
    Write-Host "## Desempenho (WinSAT)"
    try {
        $winSAT = Get-CimInstance Win32_WinSAT -EA Stop
        $scores = @{ "Processador"="$($winSAT.CPUScore)"; "Graficos3D"="$($winSAT.D3DScore)";
                     "Disco"="$($winSAT.DiskScore)"; "Graficos"="$($winSAT.GraphicsScore)"; "Memoria"="$($winSAT.MemoryScore)" }
        $minScore=10.0; $maxScore=0.0; $minName=""; $maxName=""
        Write-Host "`nPontuacoes WinSAT:" -ForegroundColor White
        foreach ($c in $scores.GetEnumerator()) {
            $s = [double]$c.Value
            Write-Host "  $($c.Key): $s" -ForegroundColor (Get-ScoreColor $s)
            if ($s -lt $minScore) { $minScore=$s; $minName=$c.Key }
            if ($s -gt $maxScore) { $maxScore=$s; $maxName=$c.Key }
        }
        Write-Host "  Menor: $minName ($minScore)" -ForegroundColor Yellow
        Write-Host "  Maior: $maxName ($maxScore)" -ForegroundColor Green
        Write-Host "  WinSPRLevel: $($winSAT.WinSPRLevel)" -ForegroundColor (Get-ScoreColor $winSAT.WinSPRLevel)
    } catch { Write-Host "WinSAT nao disponivel." -ForegroundColor Red }

    Write-Output "--------------------------------------------------------------`n"

    $tiService = Get-Service -Name "TiService" -EA SilentlyContinue
    if ($tiService) {
        if ($tiService.Status -eq "Running") { Write-Host "Agente TiFlux: Instalado e em Execucao" -ForegroundColor Green }
        else { Write-Host "Agente TiFlux: Instalado e Parado" -ForegroundColor Red }
    } else { Write-Host "Agente TiFlux: Nao Instalado" -ForegroundColor Red }

    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    Write-Host "--------------------------------------------------------------" -ForegroundColor Yellow

    $sendToSupport = Read-Host "`nDeseja adicionar AnyDesk e enviar por e-mail? (S/[N])"
    if ($sendToSupport -notmatch "^[Ss](im)?$") {
        Write-Host "Voltando ao menu..." -ForegroundColor Yellow; Start-Sleep -Seconds 1; Clear-Host; return
    }

    $anydeskAddress   = Read-Host "Endereco do AnyDesk"
    $clientName       = Read-Host "Nome do Cliente (Enter para pular)"
    $additionalNotes  = Read-Host "Observacoes (Enter para nao)"

    $emailContent = @(
        "Computador: $env:COMPUTERNAME", "IP Local: $localIP",
        "IP Publico (IPv4): $publicIPv4", "IP Publico (IPv6): $publicIPv6",
        "AnyDesk: $anydeskAddress", "",
        "Data: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')",
        "Windows: $($computerInfo.OsName)", "Release: $winVersion",
        "Fabricante: $($computerInfo.CsManufacturer)", "Modelo: $($computerInfo.CsModel)"
    )
    if (-not [string]::IsNullOrWhiteSpace($clientName))     { $emailContent += "Cliente: $clientName" }
    if (-not [string]::IsNullOrWhiteSpace($additionalNotes)) { $emailContent += "Obs: $additionalNotes" }

    $subject = "Info Windows - $env:COMPUTERNAME"
    if (-not [string]::IsNullOrWhiteSpace($clientName)) { $subject += " - $clientName" }
    Send-Email -Subject $subject -Body ($emailContent -join "`n")

    Write-Host "Informacoes enviadas." -ForegroundColor Green
    Start-Sleep -Seconds 2; Clear-Host
}

# ==============================================================================
# OPCAO 2 - RENOMEAR COMPUTADOR
# ==============================================================================
function Rename-ComputerCustom {
    Write-Host "Nome atual: $env:COMPUTERNAME"
    do {
        $novoNome = Read-Host "Novo nome (Enter para cancelar)"
        if ([string]::IsNullOrWhiteSpace($novoNome)) {
            Write-Host "Cancelado."; Start-Sleep -Seconds 1; Clear-Host; return
        }
        if ($novoNome -match "^[a-zA-Z0-9\.\-_]+$") {
            try {
                Rename-Computer -NewName $novoNome -Force
                Write-Host "Renomeado para '$novoNome'. Reinicie para aplicar." -ForegroundColor Yellow; return
            } catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red; pause }
        } else { Write-Host "Nome invalido. Use letras, numeros, '-', '.' e '_'." -ForegroundColor Red }
    } while ($true)
}

# ==============================================================================
# OPCAO 3 - INSTALACAO DO AGENTE TIFLUX
# ==============================================================================
function Get-ClientLinksFromCSV {
    $csvFilePath = Join-Path $ScriptsPath "clientes.csv"
    if (-not (Test-Path $csvFilePath)) {
        Write-Host "Baixando clientes.csv..." -ForegroundColor Yellow
        try {
            Download-FromGit "Scripts/clientes.csv" $csvFilePath
            Write-Host "Baixado com sucesso!" -ForegroundColor Green
        } catch {
            Write-Host "Erro ao baixar CSV: $($_.Exception.Message)" -ForegroundColor Red; return $null
        }
    }
    try { return Import-Csv -Path $csvFilePath }
    catch { Write-Host "Erro ao ler CSV: $($_.Exception.Message)" -ForegroundColor Red; return $null }
}

function Download-And-InstallAgent {
    param([string]$clientName, [string]$downloadLink)
    $fileName = Join-Path $BasePath "$clientName.msi"
    Write-Host "Baixando agente para $clientName..."
    try {
        Invoke-WebRequest -Uri $downloadLink -OutFile $fileName
        Write-Host "Baixado!" -ForegroundColor Green
    } catch { Write-Host "Erro ao baixar: $($_.Exception.Message)" -ForegroundColor Red; return }
    Write-Host "Instalando..."
    Start-Process msiexec.exe -ArgumentList "/i", $fileName, "/quiet", "/norestart" -NoNewWindow -Wait
    Write-Host "Instalacao concluida para $clientName."
    pause; Clear-Host
}

function Install-Agente {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "Instalar Agente TiFlux" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    $service = Get-Service -Name "TiService" -EA SilentlyContinue
    if ($service -and $service.Status -eq "Running") {
        Write-Host "`nAgente TiFlux ja instalado." -ForegroundColor Yellow
        $removeConfirmation = Read-Host "Deseja remover antes de prosseguir? (S/N)"
        if ($removeConfirmation -match "S") {
            foreach ($p in @("Ti Service And Agent","TiPeerToPeer","timessenger")) {
                try { Write-Host "Removendo $p..." -ForegroundColor Cyan; winget uninstall --name $p -e }
                catch { Write-Host "Nao encontrado: $p" -ForegroundColor Yellow }
            }
            $installConfirmation = Read-Host "Continuar com a instalacao? (S/N)"
            if ($installConfirmation -notmatch "S") {
                Write-Host "Cancelado." -ForegroundColor Yellow; Start-Sleep -Seconds 2; Clear-Host; return
            }
        } else { Write-Host "Cancelado." -ForegroundColor Yellow; Start-Sleep -Seconds 2; Clear-Host; return }
    }

    $csvPath = Join-Path $ScriptsPath "clientes.csv"
    if (-not (Test-Path $csvPath)) {
        Write-Host "Baixando clientes.csv..." -ForegroundColor Yellow
        try { Download-FromGit "Scripts/clientes.csv" $csvPath; Write-Host "OK!" -ForegroundColor Green }
        catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red; Clear-Host; return }
    }

    $clients = Get-ClientLinksFromCSV
    if (-not $clients) { Write-Host "Erro ao carregar clientes." -ForegroundColor Red; Clear-Host; return }

    while ($true) {
        $nomeCliente = Read-Host "`nNome do cliente (ou '0' para sair)"
        if ($nomeCliente -eq "0") {
            Remove-Item $csvPath -Force -EA SilentlyContinue
            Write-Host "CSV removido." -ForegroundColor Yellow; Start-Sleep -Seconds 2; Clear-Host; return
        }

        $resultados = $clients | Where-Object { $_.Cliente -like "*$nomeCliente*" }
        if ($resultados.Count -eq 0) {
            Write-Host "Nenhum cliente encontrado." -ForegroundColor Red
            $listarTodos = Read-Host "Listar todos? (S/N)"
            if ($listarTodos -match "S") {
                $clients | ForEach-Object { Write-Host "$($_.Cliente) - Indice: $([array]::IndexOf($clients, $_) + 1)" }
            }
        } else {
            Write-Host "`nClientes encontrados:" -ForegroundColor Cyan
            $resultados | ForEach-Object { Write-Host "$($_.Cliente) - Indice: $([array]::IndexOf($clients, $_) + 1)" }
            $input = Read-Host "`nNumero do cliente (ou '0' para cancelar)"
            if ($input -eq "0") { pause; Clear-Host; return }
            $clientIndex = 0
            if ([int]::TryParse($input,[ref]$clientIndex) -and $clientIndex -ge 1 -and $clientIndex -le $clients.Count) {
                $sel = $clients[$clientIndex - 1]
                Write-Host "`nSelecionado: $($sel.Cliente)" -ForegroundColor Green
                if ((Read-Host "Confirmar? (S/N)") -match "S") {
                    Download-And-InstallAgent $sel.Cliente $sel.'Link-Agente'; return
                } else { Write-Host "Cancelado." -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
            } else { Write-Host "Indice invalido." -ForegroundColor Red }
        }
    }
}

# ==============================================================================
# OPCAO 5 - PACOTE BASICO DE PROGRAMAS
# ==============================================================================
function Install-BasicPackage {
    if (-not (Check-WingetInstalled)) {
        Write-Host "`nWinget nao instalado." -ForegroundColor Red; pause; Clear-Host; return
    }
    Accept-WingetTerms

    $packages = @(
        @{ Name = "7-Zip";                     Command = "winget install --id=7zip.7zip -e" },
        @{ Name = "Google Chrome";             Command = "winget install --id=Google.Chrome -e" },
        @{ Name = "AnyDesk";                   Command = "winget install --id=AnyDeskSoftwareGmbH.AnyDesk -e" },
        @{ Name = "Lightshot";                 Command = "winget install --id=Skillbrains.Lightshot -e" },
        @{ Name = "Adobe Acrobat Reader 64bit";Command = "winget install --id=Adobe.Acrobat.Reader.64-bit -e" }
    )

    Write-Host "`nProgramas disponiveis para instalar/atualizar:" -ForegroundColor Cyan
    foreach ($p in $packages) { Write-Host "- $($p.Name)" -ForegroundColor Yellow }

    $confirm = Read-Host "`nConfirmar instalacao/atualizacao? (S/N)"
    if ($confirm -ne "S") { Write-Host "Cancelado." -ForegroundColor Red; Start-Sleep -Seconds 2; Clear-Host; return }

    foreach ($package in $packages) {
        $name = $package.Name; $command = $package.Command
        $installed = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -EA SilentlyContinue |
            Where-Object { $_.DisplayName -like "*$name*" }

        if ($installed) {
            Write-Host "`n$name ja instalado." -ForegroundColor Yellow
            if ((Read-Host "Atualizar $name? (S/N)") -eq "S") {
                Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c $command --silent --force"
                Write-Host "$name atualizado." -ForegroundColor Green
            } else { Write-Host "Atualizacao de $name cancelada." -ForegroundColor Red }
        } else {
            Write-Host "`nInstalando $name..." -ForegroundColor Yellow
            Start-Process -NoNewWindow -Wait -FilePath "cmd.exe" -ArgumentList "/c $command --silent"
            Write-Host "$name instalado." -ForegroundColor Green
        }
    }

    Write-Host "`nConcluido." -ForegroundColor Cyan
    [void][System.Console]::ReadKey($true); Clear-Host
}

# ==============================================================================
# OPCAO 6 - PACOTE JURIDICO (Java, PJe Office, Certisign)
# PARA ADICIONAR NOVOS SOFTWARES: adicione um bloco no switch abaixo
# ==============================================================================
function Install-LegalPackage {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "      Pacote de Programas Juridicos" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Green

    Write-Host "Selecione o que deseja instalar:`n"
    Write-Host "A) Java 32 bits (JRE - necessario para PJe/sistemas juridicos)" -ForegroundColor Green
    Write-Host "B) Java 64 bits (JRE - necessario para PJe/sistemas juridicos)" -ForegroundColor Green
    Write-Host "C) PJe Office Pro (cliente do processo judicial eletronico)"    -ForegroundColor Green
    Write-Host "D) Driver Certificado Digital Certisign (abre site de download)" -ForegroundColor Green

    # ------------------------------------------------------------------
    # PARA ADICIONAR NOVO SOFTWARE JURIDICO:
    # 1) Adicione uma linha Write-Host com a letra e descricao
    # 2) Adicione o case correspondente no switch abaixo
    # ------------------------------------------------------------------

    Write-Host ""
    Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    Write-Host ""

    $choice = Read-Host "Digite a letra da opcao desejada"

    switch ($choice.ToUpper()) {

        # ------------------------------------------------------------------
        # CASE A - Java 32 bits
        # ------------------------------------------------------------------
        "A" {
            Write-Host "`nBaixando e instalando Java 32 bits (JRE)..." -ForegroundColor Cyan
            try {
                $javaPath = Join-Path $ProgramsPath "java32.exe"
                $javaUrl  = "https://javadl.oracle.com/webapps/download/AutoDL?BundleId=249552_4d245f941845490c91360409ecffb3b4"
                Write-Host "Baixando..." -ForegroundColor Yellow
                Invoke-WebRequest -Uri $javaUrl -OutFile $javaPath -UseBasicParsing -EA Stop
                Write-Host "Instalando (silencioso)..." -ForegroundColor Yellow
                Start-Process -FilePath $javaPath -ArgumentList "/s" -Wait -NoNewWindow
                Write-Host "Java 32 bits instalado com sucesso!" -ForegroundColor Green
            } catch {
                Write-Host "Erro ao instalar Java 32 bits: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Tente manualmente em: https://www.java.com/pt-BR/download/" -ForegroundColor Yellow
            }
            pause; Install-LegalPackage
        }

        # ------------------------------------------------------------------
        # CASE B - Java 64 bits
        # ------------------------------------------------------------------
        "B" {
            Write-Host "`nBaixando e instalando Java 64 bits (JRE)..." -ForegroundColor Cyan
            try {
                $javaPath = Join-Path $ProgramsPath "java64.exe"
                $javaUrl  = "https://javadl.oracle.com/webapps/download/AutoDL?BundleId=249554_4d245f941845490c91360409ecffb3b4"
                Write-Host "Baixando..." -ForegroundColor Yellow
                Invoke-WebRequest -Uri $javaUrl -OutFile $javaPath -UseBasicParsing -EA Stop
                Write-Host "Instalando (silencioso)..." -ForegroundColor Yellow
                Start-Process -FilePath $javaPath -ArgumentList "/s" -Wait -NoNewWindow
                Write-Host "Java 64 bits instalado com sucesso!" -ForegroundColor Green
            } catch {
                Write-Host "Erro ao instalar Java 64 bits: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Tente manualmente em: https://www.java.com/pt-BR/download/" -ForegroundColor Yellow
            }
            pause; Install-LegalPackage
        }

        # ------------------------------------------------------------------
        # CASE C - PJe Office Pro
        # ------------------------------------------------------------------
        "C" {
            Write-Host "`nBaixando e instalando PJe Office Pro..." -ForegroundColor Cyan
            try {
                $pjePath = Join-Path $ProgramsPath "PjeOffice.exe"
                $pjeUrl  = "https://www.pje.jus.br/wiki/images/PjeOffice-win64.exe"
                Write-Host "Baixando..." -ForegroundColor Yellow
                Invoke-WebRequest -Uri $pjeUrl -OutFile $pjePath -UseBasicParsing -EA Stop
                Write-Host "Iniciando instalacao..." -ForegroundColor Yellow
                Start-Process -FilePath $pjePath -Wait
                Write-Host "PJe Office Pro instalado com sucesso!" -ForegroundColor Green
            } catch {
                Write-Host "Erro ao baixar PJe Office: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Baixe manualmente em: https://www.pje.jus.br/wiki/index.php/PjeOffice" -ForegroundColor Yellow
            }
            pause; Install-LegalPackage
        }

        # ------------------------------------------------------------------
        # CASE D - Driver Certisign (abre site e pasta de downloads)
        # ------------------------------------------------------------------
        "D" {
            Write-Host "`nAbrindo site de suporte e downloads da Certisign..." -ForegroundColor Cyan
            try {
                Start-Process "https://certisign.com.br/suporte/download"
                Write-Host "Site aberto no navegador padrao." -ForegroundColor Green
                Write-Host "Faca o download do driver diretamente pela pagina." -ForegroundColor Yellow
            } catch {
                Write-Host "Nao foi possivel abrir o navegador: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Acesse manualmente: https://certisign.com.br/suporte/download" -ForegroundColor Yellow
            }
            pause; Install-LegalPackage
        }

        # ------------------------------------------------------------------
        # CASE 0 - Voltar
        # ------------------------------------------------------------------
        "0" { Clear-Host; return }

        default {
            Write-Host "Opcao invalida." -ForegroundColor Red
            Start-Sleep -Seconds 2; Install-LegalPackage
        }
    }
}

# ==============================================================================
# OPCAO 8 - INSTALACAO DO OFFICE
# ==============================================================================
function Install-Office {
    try {
        Write-Host "`nDeseja instalar o Office? [S/n]" -ForegroundColor Yellow
        $confirm = Read-Host "(Enter = SIM)"
        if ($confirm -and $confirm -notmatch "^[Ss]$") {
            Write-Host "Cancelado." -ForegroundColor Red; Start-Sleep -Seconds 2; Clear-Host; return
        }
        try {
            $officeProcesses = Get-Process | Where-Object { $_.Name -like "*office*" }
            if ($officeProcesses) {
                Write-Host "`nProcessos Office em execucao:" -ForegroundColor Yellow
                $officeProcesses | ForEach-Object { Write-Host "- $($_.Name)" -ForegroundColor Cyan }
                if ((Read-Host "Forcar encerramento? [S/n] (Enter=SIM)") -notmatch "^[Ss]$") {
                    Write-Host "Cancelado." -ForegroundColor Red; Start-Sleep -Seconds 2; return
                }
                $officeProcesses | ForEach-Object {
                    try { Stop-Process -Id $_.Id -Force; Write-Host "'$($_.Name)' encerrado." -ForegroundColor Green }
                    catch { Write-Host "Erro ao encerrar '$($_.Name)'." -ForegroundColor Red }
                }
            }
        } catch { Write-Host "Erro ao verificar processos: $($_.Exception.Message)" -ForegroundColor Red }

        $officeInstallerPath = Join-Path $ProgramsPath "OfficeSetup.exe"
        Write-Host "`nBaixando instalador do Office..." -ForegroundColor Cyan
        try {
            Download-FromGit "Programas/OfficeSetup.exe" $officeInstallerPath
            Write-Host "Baixado em '$officeInstallerPath'." -ForegroundColor Green
        } catch { Write-Host "Erro ao baixar Office: $($_.Exception.Message)" -ForegroundColor Red; return }

        Write-Host "`nIniciando instalacao..." -ForegroundColor Yellow
        try { Start-Process -FilePath $officeInstallerPath; Write-Host "Continue com a instalacao manual." -ForegroundColor Green }
        catch { Write-Host "Erro ao executar instalador: $($_.Exception.Message)" -ForegroundColor Red }
    } catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red }
    Write-Host ""; pause; Clear-Host
}

# ==============================================================================
# OPCAO 9 - ATIVACAO DO WINDOWS E OFFICE
# ==============================================================================
function Activate-WindowsOffice {
    try {
        Clear-Host
        Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "         Ativacao do Windows e Office" -ForegroundColor Cyan
        Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
        Write-Host "Esta opcao executa a ferramenta Microsoft Activation Scripts." -ForegroundColor Yellow
        Write-Host "Fonte: https://get.activated.win`n" -ForegroundColor DarkGray

        $confirm = Read-Host "Iniciar ativacao? [S/n] (Enter=SIM)"
        if ($confirm -and $confirm -notmatch "^[Ss]$") {
            Write-Host "Cancelado." -ForegroundColor Red; Start-Sleep -Seconds 2; Clear-Host; return
        }

        Write-Host "`nIniciando..." -ForegroundColor Cyan
        irm https://get.activated.win | iex

        Write-Host "`nFinalize conforme as instrucoes na tela." -ForegroundColor Yellow
        Write-Host "Pressione qualquer tecla para voltar ao menu..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red; pause
    }
    Clear-Host
}

# ==============================================================================
# OPCAO 10 - REMOVER APPS DESNECESSARIOS
# ==============================================================================
function Remove-Apps {
    if (-not (Check-WingetInstalled)) {
        Write-Host "Winget nao instalado." -ForegroundColor Red; pause; Clear-Host; return
    }
    Accept-WingetTerms

    # Para adicionar apps: inclua o nome exato na lista abaixo
    $appsToRemove = @(
        "Tactical RMM Agent","Agente Milvus","Mesh Agent","Solitaire","Jogo de Copas",
        "Spades","Solitaire & Casual Games","Game Bar","Game Speech Window","Noticias",
        "Xbox","Xbox TCUI","Xbox Game Bar Plugin","Xbox Console Companion",
        "Xbox Game Speech Window","Xbox Identity Provider","WebAdvisor da McAfee",
        "Dropbox - promocao","Microsoft Bing"
    )

    $wingetOutput = winget list | ForEach-Object {
        $fields = $_ -split '\s{2,}'
        [PSCustomObject]@{ Name = $fields[0]; Id = if ($fields.Count -gt 1) { $fields[1] } else { "" } }
    }

    $installedApps = $wingetOutput | Where-Object {
        $name = $_.Name
        $appsToRemove | Where-Object { $_ -eq $name -or $_.ToLower() -eq $name.ToLower() }
    }

    if ($installedApps.Count -eq 0) {
        Write-Host "`nNenhum app para remover encontrado." -ForegroundColor Yellow; pause; Clear-Host; return
    }

    Write-Host "`nApps a serem removidos:`n" -ForegroundColor Cyan
    $installedApps | ForEach-Object { Write-Host $_.Name }

    if ((Read-Host "Pressione Enter para continuar ou 'N' para cancelar") -eq 'N') {
        Write-Host "Cancelado." -ForegroundColor Yellow; pause; Clear-Host; return
    }

    foreach ($app in $installedApps) {
        Write-Host "Removendo $($app.Name)..." -ForegroundColor Green
        try {
            winget uninstall --name "$($app.Name)" --silent
            Write-Host "$($app.Name) removido.`n" -ForegroundColor Green
        } catch { Write-Host "Falha ao remover $($app.Name)." -ForegroundColor Red }
    }

    Write-Host "Concluido." -ForegroundColor Cyan; pause; Clear-Host
}

# ==============================================================================
# OPCAO 11 - DOWNLOAD DE UTILITARIOS (SELECIONAVEL)
# PARA ADICIONAR NOVO UTILITARIO: inclua um bloco @{} na lista $utils abaixo
# ==============================================================================
function Download-Utilities {
    Clear-Host

    # ------------------------------------------------------------------
    # LISTA DE UTILITARIOS DISPONIVEIS
    # Para adicionar: copie um bloco @{} e ajuste Name, File e Source
    # Source pode ser "git" (baixa do GitHub) ou "url" (URL direta)
    # ------------------------------------------------------------------
    $utils = @(
        @{ Name = "CrystalDiskInfo (Saude do HD/SSD)";  File = "CrystalDiskInfo.exe"; Source = "git"; Path = "Utilitarios/CrystalDiskInfo.exe" },
        @{ Name = "Sysinternals Suite (Microsoft)";     File = "SysinternalsSuite.zip"; Source = "git"; Path = "Utilitarios/SysinternalsSuite.zip" },
        @{ Name = "BlueScreenView (Analise de BSOD)";   File = "BlueScreenView.exe"; Source = "git"; Path = "Utilitarios/BlueScreenView.exe" },
        @{ Name = "WinDirStat (Uso de disco visual)";   File = "WinDirStat.exe"; Source = "git"; Path = "Utilitarios/WinDirStat.exe" },
        @{ Name = "HWiNFO (Info de hardware)";          File = "HWiNFO.exe"; Source = "git"; Path = "Utilitarios/HWiNFO.exe" }
        # Exemplo de URL direta:
        # @{ Name = "Exemplo URL Direta"; File = "app.exe"; Source = "url"; Url = "https://site.com/app.exe" }
    )

    Write-Host "`n--------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "         Download de Aplicativos Utilitarios de Suporte" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor DarkGray
    Write-Host "Utilitarios disponiveis:`n" -ForegroundColor Cyan

    for ($i = 0; $i -lt $utils.Count; $i++) {
        $num = $i + 1
        Write-Host "$num) $($utils[$i].Name)" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "T) Baixar TODOS os utilitarios" -ForegroundColor Green
    Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    Write-Host ""

    $choice = Read-Host "Digite o numero do utilitario desejado (ou T para todos)"

    if ($choice -eq "0") { Clear-Host; return }

    # Monta lista do que vai baixar
    $toDownload = @()
    if ($choice.ToUpper() -eq "T") {
        $toDownload = $utils
    } else {
        $idx = 0
        if ([int]::TryParse($choice, [ref]$idx) -and $idx -ge 1 -and $idx -le $utils.Count) {
            $toDownload = @($utils[$idx - 1])
        } else {
            Write-Host "Opcao invalida." -ForegroundColor Red
            Start-Sleep -Seconds 2; Download-Utilities; return
        }
    }

    Write-Host ""
    foreach ($util in $toDownload) {
        $destPath = Join-Path $UtilsPath $util.File

        # Verifica se ja existe
        if (Test-Path $destPath) {
            $overwrite = Read-Host "$($util.Name) ja existe. Sobrescrever? (S/N)"
            if ($overwrite -notmatch "^[Ss]$") {
                Write-Host "Pulando $($util.Name)." -ForegroundColor Yellow; continue
            }
        }

        Write-Host "Baixando: $($util.Name)..." -ForegroundColor Cyan
        try {
            if ($util.Source -eq "git") {
                Download-FromGit $util.Path $destPath
            } elseif ($util.Source -eq "url") {
                Invoke-WebRequest -Uri $util.Url -OutFile $destPath -UseBasicParsing -EA Stop
            }
            Write-Host "$($util.Name) baixado com sucesso!" -ForegroundColor Green
        } catch {
            Write-Host "Falha ao baixar $($util.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "`nArquivos salvos em: $UtilsPath" -ForegroundColor Green
    Write-Host ""
    pause
    Clear-Host
}

# ==============================================================================
# OPCAO 12 - GERENCIADOR DE USUARIOS LOCAIS
# ==============================================================================
function Show-ActiveLocalUsers {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "         Usuarios Locais Ativos" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    $localUsers = Get-LocalUser | Where-Object { $_.Enabled -eq $true }
    if ($localUsers) { foreach ($u in $localUsers) { Write-Host $u.Name -ForegroundColor Cyan } }
    else { Write-Host "Nenhum usuario ativo." -ForegroundColor Yellow }
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    pause; Manage-LocalUsers
}

function Show-LocalAdminUsers {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "            Usuarios do Grupo Administradores" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    $domain = (Get-WmiObject Win32_ComputerSystem).Domain
    $isDomain = $domain -ne (Get-WmiObject Win32_ComputerSystem).Name
    if ($isDomain) {
        Write-Host "Dominio: $domain" -ForegroundColor Yellow; cmd /c "net localgroup Administradores"
    } else {
        try {
            $admins = Get-LocalGroupMember -Group "Administradores"
            if ($admins) { foreach ($a in $admins) { Write-Host $a.Name -ForegroundColor Cyan } }
            else { Write-Host "Nenhum administrador." -ForegroundColor Yellow }
        } catch { Write-Host "Erro ao acessar grupo." -ForegroundColor Red }
    }
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    pause; Manage-LocalUsers
}

function Create-NewLocalUser {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "       CRIAR NOVO USUARIO LOCAL" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green
    do {
        $username = Read-Host "Nome do usuario (ou '0' para voltar)"
        if (-not $username) { Write-Host "Nome nao pode ser vazio." -ForegroundColor Red; continue }
        if ($username -eq "0") { Manage-LocalUsers; return }
        if ($username -notmatch '^[a-zA-Z0-9._-]+$') { Write-Host "Caracteres invalidos." -ForegroundColor Red; continue }
        Write-Host "Nome: $username" -ForegroundColor Yellow
        if ((Read-Host "Confirmar? (S/N)") -in @("S","s")) {
            $existing = Get-LocalUser -Name $username -EA SilentlyContinue
            if ($existing) {
                Write-Host "Usuario '$username' ja existe." -ForegroundColor Red
                if ((Read-Host "Outro nome? (S/N)") -in @("N","n")) { Manage-LocalUsers; return }
            } else { break }
        }
    } while ($true)

    do {
        $password = Read-Host "Senha" -AsSecureString
        $confirmPassword = Read-Host "Confirme a senha" -AsSecureString
        $p1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
        $p2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirmPassword))
        if ($p1 -ne $p2) { Write-Host "Senhas nao coincidem." -ForegroundColor Red }
        else { Write-Host "Senha confirmada." -ForegroundColor Green; break }
    } while ($true)

    New-LocalUser -Name $username -Password $password | Out-Null
    if ((Read-Host "Adicionar ao grupo Administradores? (S/N)") -in @("S","s")) {
        Add-LocalGroupMember -Group "Administradores" -Member $username
        Write-Host "Usuario '$username' criado como Administrador." -ForegroundColor Green
    } else {
        $gu = Get-LocalGroup | Where-Object { $_.SID -eq 'S-1-5-32-545' }
        Add-LocalGroupMember -Group $gu -Member $username
        Write-Host "Usuario '$username' criado no grupo padrao." -ForegroundColor Green
    }
    Set-LocalUser -Name $username -PasswordNeverExpires 1
    Write-Host "==========================================" -ForegroundColor Green
    pause; Manage-LocalUsers
}

function Create-HtiUser {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "            CRIAR USUARIO HTI ADMINISTRADOR" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    $username = "HTI"
    if (Get-LocalUser -Name $username -EA SilentlyContinue) {
        Write-Host "Usuario '$username' ja existe." -ForegroundColor Red
        Start-Sleep -Seconds 2; Manage-LocalUsers; return
    }
    do {
        $password = Read-Host "Senha do usuario HTI" -AsSecureString
        $pc = Read-Host "Confirme a senha" -AsSecureString
        $p1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
        $p2 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pc))
        if ($p1 -ne $p2) { Write-Host "Senhas nao coincidem." -ForegroundColor Red }
        else { Write-Host "Senha confirmada!" -ForegroundColor Green; break }
    } while ($true)
    New-LocalUser -Name $username -Password $password | Out-Null
    Add-LocalGroupMember -Group "Administradores" -Member $username
    Set-LocalUser -Name $username -PasswordNeverExpires 1
    Write-Host "Usuario HTI criado como Administrador." -ForegroundColor Green
    if ((Read-Host "Ativar Area de Trabalho Remota (RDP)? (S/N)") -match "^[sS]$") {
        try {
            Set-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server" "fDenyTSConnections" 0
            $rdp = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "*Remote Desktop*" }
            if ($rdp) { $rdp | ForEach-Object { Enable-NetFirewallRule -Name $_.Name }; Write-Host "RDP ativado." -ForegroundColor Green }
            else { Write-Host "Regras de firewall RDP nao encontradas." -ForegroundColor Yellow }
        } catch { Write-Host "Erro ao ativar RDP." -ForegroundColor Red }
    }
    Start-Sleep -Seconds 2; pause; Manage-LocalUsers
}

function Change-LocalUserPassword {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "            Trocar Senha de Usuario Local" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Get-LocalUser | Where-Object { $_.Enabled } | Select-Object -ExpandProperty Name
    Write-Host "(Deixe em branco para remover a senha)" -ForegroundColor Yellow
    $username = Read-Host "Nome do usuario"
    if (-not (Get-LocalUser -Name $username -EA SilentlyContinue)) {
        Write-Host "Usuario '$username' nao existe." -ForegroundColor Yellow; Start-Sleep -Seconds 3; Manage-LocalUsers; return
    }
    do {
        $newPass = Read-Host "Nova senha" -AsSecureString
        $conf    = Read-Host "Confirme" -AsSecureString
        $p1 = [Runtime.InteropServices.Marshal]::PtrToStringUni([Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPass))
        $p2 = [Runtime.InteropServices.Marshal]::PtrToStringUni([Runtime.InteropServices.Marshal]::SecureStringToBSTR($conf))
        if ($p1 -ne $p2) { Write-Host "Senhas nao coincidem." -ForegroundColor Red }
        else { Write-Host "Senha confirmada!" -ForegroundColor Green; break }
    } while ($true)
    Set-LocalUser -Name $username -Password $newPass
    Write-Host "Senha de '$username' alterada!" -ForegroundColor Green
    pause; Manage-LocalUsers
}

function Enable-Disable-Remove-LocalUser {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "    Habilitar / Desabilitar / Remover Usuario" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Get-LocalUser | ForEach-Object {
        $s = if ($_.Enabled) { "Habilitado" } else { "Desabilitado" }
        Write-Host "$($_.Name) ... $s"
    }
    $username = Read-Host "Nome do usuario"
    if (-not (Get-LocalUser -Name $username -EA SilentlyContinue)) {
        Write-Host "Usuario nao existe." -ForegroundColor Yellow; Start-Sleep -Seconds 2; Manage-LocalUsers; return
    }
    $choice = Read-Host "H=Habilitar / D=Desabilitar / R=Remover"
    switch ($choice.ToUpper()) {
        "H" { Enable-LocalUser $username; Write-Host "Habilitado!" -ForegroundColor Green; pause; Manage-LocalUsers }
        "D" { Disable-LocalUser $username; Write-Host "Desabilitado!" -ForegroundColor Green; pause; Manage-LocalUsers }
        "R" {
            Write-Host "ATENCAO: acao irreversivel!" -ForegroundColor Red
            if ((Read-Host "Digite 'Sim' para confirmar") -eq "Sim") {
                Remove-LocalUser $username; Write-Host "Removido!" -ForegroundColor Green
            } else { Write-Host "Cancelado." -ForegroundColor Yellow }
            pause; Manage-LocalUsers
        }
        default { Write-Host "Opcao invalida." -ForegroundColor Yellow; Start-Sleep -Seconds 3; Manage-LocalUsers }
    }
}

function Add-Remove-UserFromAdminGroup {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "Adicionar / Remover do Grupo Administradores" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    $domain = (Get-WmiObject Win32_ComputerSystem).Domain
    $isDomain = $domain -ne (Get-WmiObject Win32_ComputerSystem).Name
    if ($isDomain) { cmd /c "net localgroup Administradores" }
    else {
        try {
            $admins = Get-LocalGroupMember -Group "Administradores"
            if ($admins) { $admins | ForEach-Object { Write-Host $_.Name -ForegroundColor Cyan } }
        } catch { Write-Host "Erro ao listar grupo." -ForegroundColor Red }
    }
    $username = Read-Host "Nome do usuario"
    if (-not $isDomain -and -not (Get-LocalUser -Name $username -EA SilentlyContinue)) {
        Write-Host "Usuario nao existe." -ForegroundColor Red; pause; return
    }
    $hostname = $env:COMPUTERNAME
    switch ((Read-Host "A=Adicionar / R=Remover").ToUpper()) {
        "A" {
            if ($isDomain) { cmd /c "net localgroup Administradores $username /add" }
            else { Add-LocalGroupMember -Group "Administradores" -Member "$hostname\$username" }
            Write-Host "Concluido." -ForegroundColor Green; pause; Manage-LocalUsers
        }
        "R" {
            if ($isDomain) { cmd /c "net localgroup Administradores $username /delete" }
            else { Remove-LocalGroupMember -Group "Administradores" -Member "$hostname\$username" }
            Write-Host "Concluido." -ForegroundColor Green; pause; Manage-LocalUsers
        }
        default { Write-Host "Opcao invalida." -ForegroundColor Red; Start-Sleep -Seconds 2; Manage-LocalUsers }
    }
}

function Manage-LocalUsers {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "              Gerenciador de Usuarios Locais" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Green
    Write-Host "1) Usuarios locais ativos"
    Write-Host "2) Usuarios do grupo Administradores"
    Write-Host "3) Criar novo usuario local"
    Write-Host "4) Criar usuario HTI Administrador"
    Write-Host "5) Trocar senha de usuario local"
    Write-Host "6) Habilitar / Desabilitar / Remover usuario"
    Write-Host "7) Adicionar / Remover do grupo Administradores"
    Write-Host ""; Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    switch (Read-Host "`nOpcao") {
        "1" { Show-ActiveLocalUsers }
        "2" { Show-LocalAdminUsers }
        "3" { Create-NewLocalUser }
        "4" { Create-HtiUser }
        "5" { Change-LocalUserPassword }
        "6" { Enable-Disable-Remove-LocalUser }
        "7" { Add-Remove-UserFromAdminGroup }
        "0" { Clear-Host; return }
        default { Write-Host "Opcao invalida." }
    }
}

# ==============================================================================
# OPCAO 13 - FERRAMENTAS E OTIMIZACOES
# ==============================================================================
function PowerUP {
    $monAC=60; $stbAC=0; $monDC=5; $stbDC=10
    $isLaptop = (Get-WmiObject Win32_SystemEnclosure).ChassisTypes -match "8|9|10|11|14|18|21|30"
    if ($isLaptop) {
        Write-Output "Notebook detectado."
        powercfg /change monitor-timeout-dc $monDC
        powercfg /change standby-timeout-dc $stbDC
        powercfg /change monitor-timeout-ac $monAC
        powercfg /change standby-timeout-ac $stbAC
    } else {
        Write-Output "Desktop detectado."
        powercfg /change monitor-timeout-ac $monAC
        powercfg /change standby-timeout-ac $stbAC
    }
    Write-Output "Configuracoes de energia aplicadas!"
    pause; Otimizacoes-Win
}

function WinUpdate {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "           Verificando atualizacoes do Windows..." -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
    $updates = Get-CimInstance -Namespace "root\ccm\clientSDK" -ClassName "CCM_SoftwareUpdate" -EA SilentlyContinue
    if (-not $updates -or $updates.Count -eq 0) {
        Write-Host "Nenhuma atualizacao disponivel." -ForegroundColor Green
    } else {
        $updates | ForEach-Object { Write-Host "- $($_.ArticleID)" -ForegroundColor White }
        if ((Read-Host "Instalar atualizacoes? (S/N)") -match "[Ss]") {
            Start-Process "C:\Windows\System32\UsoClient.exe" -ArgumentList "StartScan" -Wait
            Write-Host "Atualizacoes aplicadas. Reinicie se necessario." -ForegroundColor Green
        }
    }
    pause; Otimizacoes-Win
}

function Clear-PrintSpoolerQueue {
    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host "LIMPEZA DA FILA DE IMPRESSAO" -ForegroundColor Yellow
    Write-Host "==========================================================="
    try {
        Write-Host "`n[1] Parando Spooler..." -ForegroundColor Cyan
        Stop-Service -Name Spooler -Force -EA Stop
        $spoolPath = Join-Path $env:SystemRoot "System32\spool\PRINTERS"
        Write-Host "[2] Limpando: $spoolPath..." -ForegroundColor Cyan
        Remove-Item "$spoolPath\*" -Recurse -Force -EA Stop
        Write-Host "Fila limpa!" -ForegroundColor Green
    } catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red }
    finally {
        Write-Host "[3] Reiniciando Spooler..." -ForegroundColor Cyan
        if ((Get-Service Spooler).Status -ne 'Running') {
            Start-Service Spooler; Write-Host "Spooler reiniciado." -ForegroundColor Green
        }
        Write-Host "==========================================================="
    }
}

function Limpar-CacheTeams {
    if ((Read-Host "`nLimpar cache do Teams e desconectar conta? (S/N)") -in @("N","n")) {
        Write-Host "Cancelado." -ForegroundColor Yellow; pause; Otimizacoes-Win; return
    }
    Write-Host "Encerrando Teams..." -ForegroundColor Cyan
    @("Teams","ms-teams") | ForEach-Object {
        if (Get-Process -Name $_ -EA SilentlyContinue) { Stop-Process -Name $_ -Force -EA SilentlyContinue; Start-Sleep -Seconds 3 }
    }
    Start-Sleep -Seconds 2

    $found = $false
    $newPath    = "$env:LocalAppData\Packages\MSTeams_8wekyb3d8bbwe\LocalCache"
    $newAlt     = "$env:LocalAppData\Microsoft\Teams"
    $legacyPath = "$env:APPDATA\Microsoft\Teams"

    if (Test-Path $newPath)    { Remove-Item "$newPath\*" -Recurse -Force -EA SilentlyContinue; Write-Host "Cache nova versao limpo." -ForegroundColor Green; $found=$true }
    elseif (Test-Path $newAlt) { Remove-Item "$newAlt\*" -Recurse -Force -EA SilentlyContinue; Write-Host "Cache alternativo limpo." -ForegroundColor Green; $found=$true }
    elseif (Test-Path $legacyPath) {
        @("Application Cache\Cache","blob_storage","Cache","databases","GPUCache","IndexedDB","Local Storage","tmp") | ForEach-Object {
            $fp = Join-Path $legacyPath $_
            if (Test-Path $fp) { Remove-Item $fp -Recurse -Force -EA SilentlyContinue }
        }
        Write-Host "Cache versao classica limpo." -ForegroundColor Green; $found=$true
    }

    if (-not $found) { Write-Host "Cache do Teams nao encontrado." -ForegroundColor Red }
    else {
        $started = $false
        @("$env:LocalAppData\Microsoft\WindowsApps\ms-teams.exe","$env:LocalAppData\Microsoft\Teams\current\Teams.exe") | ForEach-Object {
            if (-not $started -and (Test-Path $_)) { Start-Process $_; Write-Host "Teams reiniciado." -ForegroundColor Green; $started=$true }
        }
        if (-not $started) { Write-Host "Inicie o Teams manualmente." -ForegroundColor Red }
    }
    Write-Host "Limpeza concluida!" -ForegroundColor Green; pause; Otimizacoes-Win
}

function Gerenciar-AnyDesk {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "           AnyDesk: Status / Instalar / Remover" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan

    $servico = Get-Service -Name "AnyDesk" -EA SilentlyContinue
    $processo = Get-Process -Name "AnyDesk" -EA SilentlyContinue
    $uninstalls = @()
    $uninstalls += Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -EA SilentlyContinue
    $uninstalls += Get-ItemProperty "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -EA SilentlyContinue
    $anydeskPath = ($uninstalls | Where-Object { $_.DisplayName -like "*AnyDesk*" }).InstallLocation
    if (-not $anydeskPath -or -not (Test-Path $anydeskPath)) { $anydeskPath = "C:\Program Files (x86)\AnyDesk" }
    $versao = $null
    if (Test-Path (Join-Path $anydeskPath "AnyDesk.exe")) {
        try { $versao = (Get-Item (Join-Path $anydeskPath "AnyDesk.exe")).VersionInfo.ProductVersion.Trim() } catch {}
    }
    if (-not $versao -and $processo) { try { $versao = ($processo | Select-Object -First 1).FileVersion.Trim() } catch {} }

    if ($servico -or $processo) {
        Write-Host "AnyDesk encontrado!" -ForegroundColor Green
        if ($servico) { Write-Host "Servico: $($servico.Status)" -ForegroundColor Green }
        if ($versao) { Write-Host "Versao: $versao" -ForegroundColor Green }
        else { Write-Host "Versao: Nao disponivel" -ForegroundColor Yellow }
    } else { Write-Host "AnyDesk nao encontrado no sistema." -ForegroundColor Yellow }

    Write-Host "`n1) Instalar ultima versao (Winget)"
    Write-Host "2) Baixar e instalar AnyDesk v7 (GitHub)"
    Write-Host "3) Remover AnyDesk"
    Write-Host "0) Voltar" -ForegroundColor DarkGreen

    switch ((Read-Host "`nOpcao")) {
        "1" {
            Write-Host "Instalando via Winget..." -ForegroundColor Cyan
            winget install AnyDesk -e --silent; pause; Otimizacoes-Win
        }
        "2" {
            $arquivo = Join-Path $UtilsPath "AnyDesk-v7.exe"
            Write-Host "Baixando AnyDesk v7..." -ForegroundColor Cyan
            try {
                Download-FromGit "Utilitarios/AnyDesk-v7.exe" $arquivo
                Start-Process $arquivo -ArgumentList "/S /U=1" -Wait -NoNewWindow
                Write-Host "AnyDesk v7 instalado!" -ForegroundColor Green
            } catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red }
            pause; Gerenciar-AnyDesk
        }
        "3" {
            if (-not ($servico -or $processo)) { Write-Host "AnyDesk nao esta instalado." -ForegroundColor Yellow; pause; Gerenciar-AnyDesk; return }
            if ((Read-Host "Confirmar remocao? (S/N)").ToUpper() -ne "S") { Write-Host "Cancelado." -ForegroundColor Yellow; return }
            if ($servico) { Stop-Service "AnyDesk" -Force -EA SilentlyContinue }
            if ($processo) { Stop-Process -Name "AnyDesk" -Force -EA SilentlyContinue }
            winget uninstall AnyDesk -e --silent --force
            @("$env:APPDATA\AnyDesk","$env:LOCALAPPDATA\AnyDesk","C:\ProgramData\AnyDesk") |
                ForEach-Object { Remove-Item $_ -Recurse -Force -EA SilentlyContinue }
            Write-Host "AnyDesk removido!" -ForegroundColor Green; pause; Otimizacoes-Win
        }
        "0" { Clear-Host; Otimizacoes-Win }
        default { Write-Host "Opcao invalida." -ForegroundColor Red }
    }
}

function LogsEventViewer {
    Clear-Host
    Write-Host "=================================================="
    Write-Host "     Analisador de Logs do Windows para TI"
    Write-Host "=================================================="
    Write-Host "Simplificado nesta versao. Abra o Visualizador de Eventos manualmente." -ForegroundColor Yellow
    pause; Clear-Host; Otimizacoes-Win
}

function CleanTemp {
    Clear-Host
    Write-Host "===================================================" -ForegroundColor Green
    Write-Host " LIMPEZA DE ARQUIVOS TEMPORARIOS" -ForegroundColor Green
    Write-Host "===================================================" -ForegroundColor Green
    try {
        Remove-Item "$env:TEMP\*" -Recurse -Force -EA SilentlyContinue
        Remove-Item "C:\Windows\Temp\*" -Recurse -Force -EA SilentlyContinue
        Write-Host "Limpeza concluida!" -ForegroundColor Green
    } catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red }
    pause; Otimizacoes-Win
}

function SystemIntegrityCheck {
    Clear-Host
    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host "VERIFICACAO DE INTEGRIDADE DO SISTEMA" -ForegroundColor Yellow
    Write-Host "==========================================================="
    $r = $host.UI.PromptForChoice("Confirmar","Iniciar verificacao e reparo?",[System.Management.Automation.Host.ChoiceDescription[]]@("&Sim","&Nao"),0)
    if ($r -eq 1) { Write-Host "Cancelado." -ForegroundColor Yellow; Clear-Host; Otimizacoes-Win; return }
    try {
        Write-Host "[1/2] DISM RestoreHealth..." -ForegroundColor Cyan
        Repair-WindowsImage -Online -RestoreHealth -EA Stop
        Write-Host "DISM concluido." -ForegroundColor Green; Start-Sleep -Seconds 3
        Write-Host "[2/2] SFC /scannow..." -ForegroundColor Cyan
        sfc.exe /scannow
        Write-Host "Verificacao concluida!" -ForegroundColor Green
    } catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red }
    pause; Otimizacoes-Win
}

function NetworkStackReset {
    Clear-Host
    Write-Host "================== Reset de Rede ==================" -ForegroundColor Yellow
    $r = $host.UI.PromptForChoice("Confirmar","Prosseguir?",[System.Management.Automation.Host.ChoiceDescription[]]@("&Sim","&Nao"),1)
    if ($r -ne 0) { Write-Host "Cancelado." -ForegroundColor Yellow; return }
    try {
        Clear-DnsClientCache
        ipconfig /release | Out-Null; ipconfig /renew | Out-Null
        netsh.exe winsock reset | Out-Null; netsh.exe int ip reset | Out-Null
        Write-Host "Reset de rede concluido!" -ForegroundColor Green
    } catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red }
    pause; Otimizacoes-Win
}

function Otimizacoes-Win {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "           Ferramentas e Otimizacoes do Windows" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Green
    Write-Host "1)  Gerenc. Energia - Sempre ativo"           -ForegroundColor Green
    Write-Host "2)  Atualizacao do Windows"                   -ForegroundColor Green
    Write-Host "3)  Microsoft PowerToys - PENDENTE"           -ForegroundColor DarkGray
    Write-Host "4)  Atualizacao de Drivers - PENDENTE"        -ForegroundColor DarkGray
    Write-Host "5)  Limpeza da Fila de Impressao"             -ForegroundColor Green
    Write-Host "6)  Atualizar programas (Winget) - PENDENTE"  -ForegroundColor DarkGray
    Write-Host "7)  Limpeza de arquivos temporarios"          -ForegroundColor Green
    Write-Host "8)  Verificacao do disco - PENDENTE"          -ForegroundColor DarkGray
    Write-Host "9)  Integridade do sistema (DISM + SFC)"      -ForegroundColor Green
    Write-Host "10) Reset de Rede"                            -ForegroundColor Green
    Write-Host "11) Logs do EventViewer"                      -ForegroundColor Green
    Write-Host "12) Todas as opcoes - PENDENTE"               -ForegroundColor DarkGray
    Write-Host "13) Limpeza de cache do Teams"                -ForegroundColor Green
    Write-Host "14) AnyDesk: Status / Instalar / Remover"     -ForegroundColor Green
    Write-Host ""; Write-Host "0) Voltar ao menu principal"   -ForegroundColor DarkGreen

    switch (Read-Host "`nOpcao") {
        "1"  { PowerUP }
        "2"  { WinUpdate }
        "3"  { Write-Host "Pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
        "4"  { Write-Host "Pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
        "5"  { Clear-PrintSpoolerQueue; pause; Otimizacoes-Win }
        "6"  { Write-Host "Pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
        "7"  { CleanTemp }
        "8"  { Write-Host "Pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
        "9"  { SystemIntegrityCheck }
        "10" { NetworkStackReset }
        "11" { LogsEventViewer }
        "12" { Write-Host "Pendente." -ForegroundColor Yellow; pause; Otimizacoes-Win }
        "13" { Limpar-CacheTeams }
        "14" { Gerenciar-AnyDesk }
        "0"  { Clear-Host; Main }
        default { Write-Host "Opcao invalida." }
    }
}

# ==============================================================================
# OPCAO 14 - MONITORAMENTO DE DISCO
# ==============================================================================
function Check-SSDHealth {
    Clear-Host
    try {
        Get-PhysicalDisk | ForEach-Object {
            Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host "Disco: $($_.DeviceID) | Modelo: $($_.Model)" -ForegroundColor Green
            Write-Host "Saude: $($_.HealthStatus) | Status: $($_.OperationalStatus) | Tipo: $($_.MediaType)" -ForegroundColor Green
            Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
        }
    } catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red }
    pause; Disk-ManagementMenu
}

function Check-DiskUsage {
    param([string]$DriveLetter="C:", [int]$Threshold=0)
    try {
        $drive = Get-PSDrive -Name $DriveLetter.TrimEnd(':') -EA Stop
        $used  = [math]::Round((($drive.Used / ($drive.Free + $drive.Used)) * 100), 2)
        Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Unidade: $DriveLetter | Computador: $env:COMPUTERNAME | Usuario: $env:USERNAME" -ForegroundColor Green
        Write-Host "Uso: ${used}% | Total: $([math]::Round(($drive.Free+$drive.Used)/1GB,2))GB | Livre: $([math]::Round($drive.Free/1GB,2))GB | Usado: $([math]::Round($drive.Used/1GB,2))GB" -ForegroundColor Green
        Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
        if ($used -ge $Threshold) { Write-Host "ALERTA: ${used}% de uso!" -ForegroundColor Red }
        if ((Read-Host "Enviar por e-mail? (S/N)") -match "^[Ss]$") {
            Send-Email -Subject "Disco $DriveLetter - $env:COMPUTERNAME" -Body "Uso: ${used}%`nTotal: $([math]::Round(($drive.Free+$drive.Used)/1GB,2))GB`nLivre: $([math]::Round($drive.Free/1GB,2))GB"
        } else { Start-Sleep -Seconds 2; Disk-ManagementMenu }
    } catch { Write-Host "Unidade '$DriveLetter' nao encontrada." -ForegroundColor Red }
}

function Select-Drive {
    $drives = Get-PSDrive | Where-Object { $_.Provider -like "*FileSystem*" }
    if ($drives.Count -eq 0) { Write-Host "Nenhuma unidade encontrada." -ForegroundColor Red; return $null }
    $drives | ForEach-Object { Write-Host " - $($_.Name):\ ($([math]::Round($_.Used/1GB,2))GB usados / $([math]::Round($_.Free/1GB,2))GB livres)" -ForegroundColor Green }
    $dl = Read-Host "Letra da unidade (Enter=C)"
    if ([string]::IsNullOrWhiteSpace($dl)) { $dl = "C" }
    if ($drives | Where-Object { $_.Name -eq $dl }) { return $dl }
    Write-Host "Unidade '$dl' nao encontrada." -ForegroundColor Red; return $null
}

function Schedule-DiskMonitoring {
    $scriptPath = Join-Path $ScriptsPath "MonitorDisk.ps1"
    if (-not (Test-Path $scriptPath)) {
        Write-Host "Baixando MonitorDisk.ps1..." -ForegroundColor Yellow
        try { Download-FromGit "Scripts/MonitorDisk.ps1" $scriptPath; Write-Host "OK!" -ForegroundColor Green }
        catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red; return }
    }
    if ((Read-Host "Criar agendamento? [S/n] (Enter=SIM)") -and (Read-Host) -notmatch "^[Ss]$") {
        Disk-ManagementMenu; return
    }
    Write-Host "1) Diario  2) Semanal  3) Mensal"
    $freq = switch (Read-Host "Frequencia") { "1"{"DAILY"} "2"{"WEEKLY"} "3"{"MONTHLY"} default { Write-Host "Invalido."; Disk-ManagementMenu; return } }
    $taskName = "Monitoramento de Disco"; $taskFolder = "_HTI"
    $svc = New-Object -ComObject "Schedule.Service"; $svc.Connect()
    $root = $svc.GetFolder("\")
    try { $root.GetFolder("\$taskFolder") | Out-Null } catch { $root.CreateFolder("\$taskFolder") | Out-Null }
    if (Get-ScheduledTask -TaskName $taskName -TaskPath "\$taskFolder\" -EA SilentlyContinue) {
        if ((Read-Host "Tarefa ja existe. Substituir? (S/N)") -notmatch "^[Ss]$") { Disk-ManagementMenu; return }
        Unregister-ScheduledTask -TaskName $taskName -TaskPath "\$taskFolder\" -Confirm:$false
    }
    $taskPath = "\$taskFolder\$taskName"; $cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    switch ($freq) {
        "DAILY"   { schtasks.exe /Create /F /TN "$taskPath" /TR $cmd /SC DAILY /ST 12:00 /RL HIGHEST /RU "SYSTEM" }
        "WEEKLY"  { schtasks.exe /Create /F /TN "$taskPath" /TR $cmd /SC WEEKLY /D MON /ST 12:00 /RL HIGHEST /RU "SYSTEM" }
        "MONTHLY" { schtasks.exe /Create /F /TN "$taskPath" /TR $cmd /SC MONTHLY /D MON /MO FIRST /ST 12:00 /RL HIGHEST /RU "SYSTEM" }
    }
    Write-Host "Agendamento criado!" -ForegroundColor Green; pause; Disk-ManagementMenu
}

function Disk-ManagementMenu {
    Clear-Host
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "       Status, Monitoramento e Alertas de Disco" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "1) Saude dos discos"  -ForegroundColor Cyan
    Write-Host "2) Uso do HD (com e-mail)" -ForegroundColor Cyan
    Write-Host "3) Agendar monitoramento" -ForegroundColor Cyan
    Write-Host ""; Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    switch (Read-Host "Opcao") {
        "1" { Check-SSDHealth }
        "2" { $dl = Select-Drive; if ($dl) { Check-DiskUsage -DriveLetter $dl -Threshold 90 } }
        "3" { Schedule-DiskMonitoring }
        "0" { Clear-Host; return }
        default { Write-Host "Opcao invalida." -ForegroundColor Red }
    }
}

# ==============================================================================
# REMOCAO DE ARQUIVOS TEMPORARIOS DO SCRIPT
# ==============================================================================
function Remove-HTIFiles {
    @((Join-Path $ScriptsPath "Script-Inicial.ps1"), (Join-Path $ScriptsPath "clientes.csv")) |
        ForEach-Object { if (Test-Path $_) { Remove-Item $_ -Force -EA SilentlyContinue } }
}

# ==============================================================================
# CHANGELOG
# Adicione entradas aqui a cada nova versao
# ==============================================================================
function Show-Changelog {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "                  HISTORICO DE ALTERACOES" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
    Write-Host "[v0.0.3] - Senha de acesso adicionada (suroh)"
    Write-Host "[v0.0.3] - Opcao 11 agora permite selecionar utilitario individual"
    Write-Host "[v0.0.3] - Opcao 6 implementada: Java 32/64, PJe Office Pro, Certisign"
    Write-Host "[v0.0.3] - Logo HTI em DarkBlue, cabecalho simplificado para HORUS TI SOLUCOES"
    Write-Host "[v0.0.3] - Limpeza completa do historico PS ao sair (opcao 0)"
    Write-Host "[v0.0.2] - Reset de rede e integridade do sistema adicionados"
    Write-Host "[v0.0.1] - Versao inicial com estrutura base GitHub"
    Write-Host "`nPressione qualquer tecla para voltar..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); Main
}

# ==============================================================================
# MENU PRINCIPAL
# Para adicionar uma nova opcao:
# 1) Adicione um Write-Host com o numero e descricao abaixo
# 2) Adicione o case no switch com a chamada da funcao
# ==============================================================================
function Main {
    Clear-Host
    Initialize-HTIStructure
    Remove-HTIFiles

    if (-not (Test-Admin)) {
        Write-Host "ATENCAO: Execute como Administrador.`n" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para sair..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown'); return
    }

    # Solicita senha antes de exibir o menu
    if (-not (Request-AccessPassword)) { return }

    do {
        Clear-Host
        Write-Host ""
        Write-Host "  ██╗  ██╗████████╗██╗" -ForegroundColor DarkBlue
        Write-Host "  ██║  ██║╚══██╔══╝██║" -ForegroundColor DarkBlue
        Write-Host "  ███████║   ██║   ██║" -ForegroundColor DarkBlue
        Write-Host "  ██╔══██║   ██║   ██║" -ForegroundColor DarkBlue
        Write-Host "  ██║  ██║   ██║   ██║" -ForegroundColor DarkBlue
        Write-Host "  ╚═╝  ╚═╝   ╚═╝   ╚═╝" -ForegroundColor DarkBlue
        Write-Host ""
        Write-Host "--------------------------------------------------------------" -ForegroundColor Blue
        Write-Host "                  HORUS TI SOLUCOES" -ForegroundColor White
        Write-Host "--------------------------------------------------------------`n" -ForegroundColor Blue
        Write-Host "Versao: 0.0.3`n" -ForegroundColor DarkGray

        Write-Host "1)  Exibir Informacoes do Windows"                              -ForegroundColor Green
        Write-Host "2)  Renomear o computador"                                      -ForegroundColor Green
        Write-Host "3)  Instalacao do agente TiFlux"                                -ForegroundColor Green
        Write-Host "4)  Instalacao do Endpoint Bitdefender - PENDENTE"              -ForegroundColor DarkGray
        Write-Host "5)  Instalacao do pacote basico de programas"                   -ForegroundColor Green
        Write-Host "6)  Pacote juridico (Java, PJe Office, Certisign)"              -ForegroundColor Green
        Write-Host "7)  Pacote contabil - PENDENTE"                                 -ForegroundColor DarkGray
        Write-Host "8)  Instalacao do Office"                                       -ForegroundColor Green
        Write-Host "9)  Ativacao do Windows e Office"                               -ForegroundColor Green
        Write-Host "10) Remove Apps desnecessarios"                                 -ForegroundColor Green
        Write-Host "11) Baixar aplicativos utilitarios (selecionar)"                -ForegroundColor Green
        Write-Host "12) Gerenciador de Usuarios Locais"                             -ForegroundColor Green
        Write-Host "13) Ferramentas e Otimizacoes"                                  -ForegroundColor Green
        Write-Host "14) Monitorar uso do HD e envio de alertas"                     -ForegroundColor Green
        Write-Host ""
        Write-Host "99) Changelog"                                                  -ForegroundColor DarkGray
        Write-Host "0)  Sair - Limpa historico"

        $mainOption = Read-Host "`nEscolha uma opcao"

        switch ($mainOption) {
            "1"  { Write-Host "Coletando informacoes..." -ForegroundColor Yellow; Info-Windows }
            "2"  { Rename-ComputerCustom }
            "3"  { Install-Agente }
            "4"  { Write-Host "Pendente." -ForegroundColor Yellow; pause }
            "5"  { Install-BasicPackage }
            "6"  { Install-LegalPackage }
            "7"  { Write-Host "Pendente." -ForegroundColor Yellow; pause }
            "8"  { Install-Office }
            "9"  { Activate-WindowsOffice }
            "10" { Remove-Apps }
            "11" { Download-Utilities }
            "12" { Manage-LocalUsers }
            "13" { Otimizacoes-Win }
            "14" { Disk-ManagementMenu }
            "99" { Show-Changelog }
            "0"  {
                Write-Host "Limpando historico e saindo..." -ForegroundColor Red
                Start-Sleep -Seconds 1
                Remove-HTIFiles
                Clear-AllPSHistory
                Clear-Host
                return
            }
            default { Write-Host "Opcao invalida." }
        }
    } while ($mainOption -ne "0")
}

Main
