# Atencao: antes de executar este scrip eh necessario ativar a permissao de execucao de scritps
# Abra o PowerShell como Administrador e execute o comando abaixo:
# Set-ExecutionPolicy Bypass
# Pressionar A para todos

Clear-Host

# Funcao para verificar se o Winget esta instalado
function Check-WingetInstalled {
    # Captura o caminho da instalacao do Microsoft.DesktopAppInstaller
    $wingetPackage = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*Microsoft.DesktopAppInstaller*" }

    # Verifica se o pacote foi encontrado
    if ($null -eq $wingetPackage -or -not $wingetPackage.InstallLocation) {
        Write-Host "Winget nao esta instalado ou nao foi encontrado." -ForegroundColor Yellow
        return $false
    }

    $wingetPath = $wingetPackage.InstallLocation

    # Construir o caminho completo para o arquivo winget.exe
    $wingetExePath = Join-Path $wingetPath "winget.exe"

    # Adiciona ao PATH apenas se o diretorio existir
    if (Test-Path $wingetExePath) {
        $env:Path = $env:Path + ";" + $wingetPath
    } else {
        Write-Host "Winget.exe nao encontrado no caminho esperado: $wingetExePath" -ForegroundColor Red
        return $false
    }

    try {
        # Tenta verificar se o comando winget esta disponivel
        $wingetCommand = Get-Command winget -ErrorAction Stop
        return $true
    } catch {
        # Se ocorrer erro, significa que o winget nao esta instalado ou nao esta no PATH
        return $false
    }
}


# Funcao para exibir animacao de carregamento
function Show-LoadingAnimation {
    $animation = @("|", "/", "-", "\")
    $i = 0
    $count = 0
    while ($count -lt 5) {  # Controle de quantas vezes a animacao vai rodar
        Write-Host -NoNewline "$($animation[$i % $animation.Length])" -ForegroundColor Green
        Start-Sleep -Milliseconds 200
        Write-Host -NoNewline "`r"
        $i++
        $count++
    }
    Write-Host ""
    return
}

function Accept-WingetTerms {
    # Forca a aceitacao dos termos da Microsoft Store
    echo Y | winget list > $null
}


# --- Funcoes auxiliares devem estar definidas aqui, fora de Info-Windows ---
# Exemplo:
# function Show-LoadingAnimation {
#    # ... seu codigo de animacao ...
# }
# function Get-ComputerInfo {
#    # ... seu codigo para Get-ComputerInfo ...
# }
# function Check-WingetInstalled {
#    # ... seu codigo para Check-WingetInstalled ...
# }
# function Send-Email {
#    # ... seu codigo para Send-Email ...
# }

# Funcao auxiliar para colorir a saida com base na pontuacao
function Get-ScoreColor {
    param([double]$Score)

    if ($Score -ge 1 -and $Score -le 3) {
        return "Red"
    } elseif ($Score -ge 4 -and $Score -le 5) {
        return "DarkYellow" # Laranja e DarkYellow no PowerShell
    } elseif ($Score -ge 5.1 -and $Score -le 6.9) { # Ajuste para nao ter sobreposicao
        return "Yellow"
    } elseif ($Score -ge 7 -and $Score -le 8.9) {
        return "Green"
    } elseif ($Score -ge 9) {
        return "Blue"
    } else {
        return "White" # Cor padrao para pontuacoes fora da faixa
    }
}

function Info-Windows {
    Clear-Host
    # Obter a versao do Windows apos a animacao
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host "Aguarde, coletando informacoes do computador..." -ForegroundColor Yellow
    # Presumindo que Show-LoadingAnimation esta definida em outro lugar no seu script
    Show-LoadingAnimation
    # Coletando informacoes adicionais
    $computerInfo = Get-ComputerInfo
    Clear-Host
    Write-Output "--------------------------------------------------------------"
    Write-Host "             -------   INFORMACOES DO WINDOWS   -------         " -ForegroundColor Blue
    Write-Output "--------------------------------------------------------------`n"

    Write-Output "Script para o Setup de Windows 10 e 11`n"
    Write-Output "Data de Execucao: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')"
    Write-Output "Nome do Computador: $env:COMPUTERNAME"
    Write-Output "Versao do Windows: $($computerInfo.OsName)"
	# Ajuste para obter a versao correta do Windows
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

#    Write-Output "Release: $($computerInfo.OSDisplayVersion)"
    Write-Output "Data Instalacao do Windows: $($computerInfo.OsInstallDate)"
	
	# Versao do PowerShell
	$psVersion = $PSVersionTable.PSVersion
	Write-Output "Versao do PowerShell: $psVersion"

	# Verificar se a versao e antiga (menor que 5.0 para compatibilidade com Get-ComputerInfo)
	if ($psVersion.Major -lt 5.0) {
		Write-Host "ATENCAO: O PowerShell esta desatualizado! Considere atualizar para a versao mais recente." -ForegroundColor Red
	}
    
	# Verificar se o Winget esta instalado e exibir com cores
    $wingetStatus = Check-WingetInstalled # Presumindo que Check-WingetInstalled esta definida em outro lugar
    if ($wingetStatus) {
        Write-Host "Winget: Instalado" -ForegroundColor Green
    } else {
        Write-Host "Winget: Nao Instalado" -ForegroundColor Red
    }

    Write-Output "Nome do Dominio (ou Workgroup): $($computerInfo.CsDomain)"
    Write-Output "Marca do Computador: $($computerInfo.CsManufacturer)"
    Write-Output "Modelo do Computador: $($computerInfo.CsModel)"
    Write-Output "Serie do Computador: $($computerInfo.CsSystemFamily)"

    # Processador
	$processorName = (Get-WmiObject Win32_Processor).Name
    Write-Output "Processador: $processorName"

    # Memoria total
    $totalMemoryGB = [math]::Round($computerInfo.OsTotalVisibleMemorySize / 1MB, 2)
    Write-Output "Total de Memoria (GB): $totalMemoryGB"

	# Uptime
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

    # IP local
    try {
        $localIP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 -ErrorAction Stop).IPV4Address.IPAddressToString
        Write-Output "IP do computador local: $localIP"
    } catch {
        Write-Output "IP do computador local: Nao foi possivel determinar. Erro: $($_.Exception.Message)"
    }
	
	# Funcao para obter o IP publico com timeout de 10 segundos (Compativel com Windows 10 e 11)
	function Get-PublicIP {
		param (
			[string]$url
		)
		
		try {
			# Carregar o assembly necessario para HttpClient (se nao estiver carregado)
			# Este Add-Type pode ser movido para fora da funcao se for chamado muitas vezes
			# para evitar recarregar desnecessariamente.
			if (-not ("System.Net.Http.HttpClient" -as [Type])) {
				Add-Type -AssemblyName "System.Net.Http"
			}

			# Criando o cliente HTTP
			$client = New-Object System.Net.Http.HttpClient
			if (-not $client) {
				throw "Falha ao criar o cliente HTTP"
			}
			
			$client.Timeout = New-TimeSpan -Seconds 10
			$responseTask = $client.ToArray().GetStringAsync($url) # Corrigido .ToArray() para .GetStringAsync()

			# Aguarda a resposta com timeout
			if ($responseTask.Wait(10000)) {  # 10000 ms = 10 segundos
				return $responseTask.Result
			} else {
				throw "Tempo limite excedido"
			}
		} catch {
			return "Nao foi possivel obter o endereco IP ($url). Erro: $($_.Exception.Message)"
		} finally {
			if ($client) { $client.Dispose() }
		}
	}

	# Obtendo os enderecos IPv4 e IPv6
	$publicIPv4 = Get-PublicIP -url "https://api.ipify.org?format=text"
	$publicIPv6 = Get-PublicIP -url "https://api64.ipify.org?format=text"

	# Verificando se o IPv6 e igual ao IPv4 ou se ambos falharam para evitar mensagens confusas
	if ($publicIPv4 -like "*Nao foi possivel obter*" -and $publicIPv6 -like "*Nao foi possivel obter*") {
		Write-Output "Nao foi possivel obter nenhum endereco IP publico."
	} elseif ($publicIPv4 -eq $publicIPv6 -or $publicIPv6 -like "*Nao foi possivel obter*") { # Se IPv6 falhou ou e igual ao IPv4 (muito raro, mas pode acontecer com CGNAT)
		Write-Output "Endereco IP Publico (IPv4): $publicIPv4"
		Write-Output "Nao ha endereco IP Publico IPv6 disponivel ou identificado."
	} else {
		Write-Output "Endereco IP Publico (IPv4): $publicIPv4"
		Write-Output "Endereco IP Publico (IPv6): $publicIPv6"
	}

    Write-Output "--------------------------------------------------------------"
    Write-Host "## Informacoes de Desempenho (WinSAT)"


    # Coletar as pontuacoes do WinSAT
    try {
        $winSAT = Get-CimInstance Win32_WinSAT -ErrorAction Stop
        $scores = @{
            "Processador" = $winSAT.CPUScore
            "Graficos 3D" = $winSAT.D3DScore
            "Disco"       = $winSAT.DiskScore
            "Graficos"    = $winSAT.GraphicsScore
            "Memoria"     = $winSAT.MemoryScore
        }

        $minScore = 10.0 # Inicializa com um valor alto para encontrar o minimo
        $maxScore = 0.0  # Inicializa com um valor baixo para encontrar o maximo
        $minComponentName = ""
        $maxComponentName = ""

        Write-Host "`nPontuacoes de Desempenho (WinSAT):" -ForegroundColor White
        foreach ($component in $scores.GetEnumerator()) {
            $score = $component.Value
            $name = $component.Key
            $color = Get-ScoreColor -Score $score
            Write-Host "  ${name}: $score" -ForegroundColor $color # Corrigido para ${name}

            # Encontrar a menor e maior nota
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
        Write-Host "Nao foi possivel obter as pontuacoes do WinSAT. Pode ser necessario executar 'winsat formal' no CMD como Administrador ou a ferramenta nao esta disponivel." -ForegroundColor Red
        Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Output "--------------------------------------------------------------`n"

    # Verifica se o servico TiService esta instalado e em execucao
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
    
    # Perguntar se deseja enviar informacoes para o suporte
    $sendToSupport = Read-Host "Deseja adicionar o endereco do AnyDesk e enviar estas informacoes por e-mail? (S/[N])"
    if ($sendToSupport -notmatch "^[Ss](im)?$") { # Usa -notmatch para simplificar a verificacao
        Write-Host "Voltando ao menu principal..." -ForegroundColor Yellow
		Start-Sleep -Seconds 1
		Clear-Host
        return
    }

    # Solicitar endereco do AnyDesk
    $anydeskAddress = Read-Host "Digite o endereco do AnyDesk"

    # Solicitar nome do Cliente
    $clientName = Read-Host "Nome do Cliente? (ou pressione Enter para pular)"
	
	# Perguntar sobre observacoes adicionais
    $additionalNotes = Read-Host "Deseja adicionar alguma observacao? [Enter para nao]"

    # Adicionar uma linha em branco apos o endereco do AnyDesk no corpo do e-mail
    # Adicionando as informacoes do WinSAT ao email tambem
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
        "Nome do usuario: $env:USERNAME",
        "",
        "--- Pontuacoes de Desempenho (WinSAT) ---",
        "Processador Score: $($winSAT.CPUScore)",
        "Graficos 3D Score: $($winSAT.D3DScore)",
        "Disco Score: $($winSAT.DiskScore)",
        "Graficos Score: $($winSAT.GraphicsScore)",
        "Memoria Score: $($winSAT.MemoryScore)",
        "Nivel de Experiencia do Windows (WinSPRLevel): $($winSAT.WinSPRLevel)"
    )
    if (-not [string]::IsNullOrWhiteSpace($clientName)) {
        $emailContent += "Nome do Cliente: $clientName"
    }
	if (-not [string]::IsNullOrWhiteSpace($additionalNotes)) {
        $emailContent += "Observacoes: $additionalNotes"
    }

    # Ajustar o assunto do e-mail
    $subject = "Informacoes do Windows - $env:COMPUTERNAME"
    if (-not [string]::IsNullOrWhiteSpace($clientName)) {
        $subject += " - $clientName"
    }

    # Enviar email usando a funcao Send-Email
    # Presumindo que Send-Email esta definida em outro lugar no seu script
    Send-Email -To "suporte@horusti.com.br" -Subject $subject -Body ($emailContent -join "`n")

    Write-Host "As informacoes foram enviadas para o suporte." -ForegroundColor Green
	Start-Sleep -Seconds 2
    Clear-Host
}



# Funcao para verificar se o script esta sendo executado como administrador
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Rename-Computer {
    Write-Host "Nome atual do computador: $env:COMPUTERNAME"
    
    do {
        $novoNome = Read-Host "Digite o novo nome para o computador (ou pressione Enter para cancelar)"
        
        if ([string]::IsNullOrWhiteSpace($novoNome)) {
            Write-Host "Operacao cancelada."
			Start-Sleep -Seconds 1
			Clear-Host
            return
        }

        # Verifica se o nome contem apenas caracteres validos
        if ($novoNome -match "^[a-zA-Z0-9\.\-_]+$") {
            try {
                Rename-Computer -NewName $novoNome -Force
                Write-Host "Computador renomeado para '$novoNome'. Reinicie para aplicar as mudancas." -ForegroundColor Yellow
                return
            } catch {
                Write-Host "Erro ao renomear o computador: $_" -ForegroundColor Red
				pause
            }
        } else {
            Write-Host "Nome invalido! Use apenas letras, numeros, '-', '.' e '_'." -ForegroundColor Red
        }
    } while ($true)
}

# Funcao para carregar clientes e links do arquivo CSV
function Get-ClientLinksFromCSV {
    $csvFilePath = "C:\_HTI\Scripts\clientes.csv"
    
    # Verifica se o arquivo existe
    if (-Not (Test-Path $csvFilePath)) {
        Write-Host "Baixando o arquivo dos links dos clientes..."
        try {
            # Baixar o arquivo CSV do servidor
            Invoke-WebRequest -Uri "https://setup.horus.net.br/Scripts/clientes.csv" -OutFile $csvFilePath
            Write-Host "Arquivo baixado com sucesso!" -ForegroundColor Green
        } catch {
            Write-Host "Erro ao baixar o arquivo CSV. Verifique a conexao." -ForegroundColor Red
            return $null
        }
    }
    
    # Importa os dados do CSV
    $clients = Import-Csv -Path $csvFilePath
    return $clients
}

# Funcao para instalar o agente
function Install-Agente {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "Instalar Agente" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
	# Verifica se o servico TiService esta instalado e em execucao
		$service = Get-Service -Name "TiService" -ErrorAction SilentlyContinue
		if ($service -ne $null -and $service.Status -eq "Running") {
			Write-Host "`nO agente TiFlux ja esta instalado no computador." -ForegroundColor Yellow
			$removeConfirmation = Read-Host "Deseja remover o agente antes de prosseguir? (S/N)"
			if ($removeConfirmation -match "S") {
				# Remove programas relacionados usando winget
				$programsToRemove = @(
					"Ti Service And Agent",
					"TiPeerToPeer",
					"timessenger"
				)

				foreach ($program in $programsToRemove) {
					try {
						Write-Host "`nRemovendo $program..." -ForegroundColor Cyan
						winget uninstall --name $program -e
					} catch {
						Write-Host "Erro ao tentar remover $program. Pode ser que ele nao esteja instalado." -ForegroundColor Yellow
					}
				}

				# Perguntar se deseja continuar com a instalacao
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
    # Verifica e baixa o arquivo CSV se necessario
    $csvPath = "C:\_HTI\Scripts\clientes.csv"
    if (-Not (Test-Path $csvPath)) {
        Write-Host "Baixando arquivo clientes.csv..." -ForegroundColor Yellow
        try {
            Invoke-WebRequest -Uri "https://setup.horus.net.br/Scripts/clientes.csv" -OutFile $csvPath
            Write-Host "Arquivo clientes.csv baixado com sucesso!" -ForegroundColor Green
        } catch {
            Write-Host "Erro ao baixar o arquivo clientes.csv. Verifique a conexao." -ForegroundColor Red
            Clear-Host
            return
        }
    }

    # Carrega os clientes e links do CSV
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
				Write-Host "indice invalido. Tente novamente." -ForegroundColor Red
			}
		}
	}
}

# Funcao para baixar o agente e iniciar a instalacao
function Download-And-InstallAgent {
    param (
        [string]$clientName,
        [string]$downloadLink
    )

    $fileName = "C:\_HTI\$clientName.msi"
        
    # Baixar o arquivo
    Write-Host "Baixando o agente para $clientName..."
    try {
        Invoke-WebRequest -Uri $downloadLink -OutFile $fileName
        Write-Host "Agente baixado com sucesso!" -ForegroundColor Green
    } catch {
        Write-Host "Erro ao baixar o agente. Verifique o link de download." -ForegroundColor Red
        return
    }

    # Iniciar a instalacao silenciosa
    Write-Host "Instalando o agente de $clientName..."
    Start-Process msiexec.exe -ArgumentList "/i", $fileName, "/quiet", "/norestart" -NoNewWindow -Wait
    Write-Host "Instalacao concluida para $clientName."
	pause
	Clear-Host
	return
}


function Install-BasicPackage {
	# Verifica se o Winget esta instalado
    if (-not (Check-WingetInstalled)) {
        Write-Host "`nO Winget nao esta instalado neste computador. Por favor, instale-o antes de prosseguir." -ForegroundColor Red
        pause
		Clear-Host
		return
    }
	Accept-WingetTerms
    # Lista de pacotes com nome e comandos de instalacao via winget
    $packages = @(
        @{ Name = "7-Zip"; Command = "winget install --id=7zip.7zip -e" },
        @{ Name = "Google Chrome"; Command = "winget install --id=Google.Chrome -e" },
        @{ Name = "AnyDesk"; Command = "winget install --id=AnyDeskSoftwareGmbH.AnyDesk -e" },
        @{ Name = "Lightshot"; Command = "winget install --id=Skillbrains.Lightshot -e" },
		@{ Name = "Adobe Acrobat Reader (64-bit)"; Command = "winget install --id=Adobe.Acrobat.Reader.64-bit -e" }
    )

    # Listar programas para o usuario
    Write-Host "`nOs seguintes programas serao instalados e/ou atualizados:" -ForegroundColor Cyan
    foreach ($package in $packages) {
        Write-Host "- $($package.Name)" -ForegroundColor Yellow
    }

    # Confirmar com o usuario
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

        # Verificar se o programa esta instalado
        $installed = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" |
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

    # Pausar e limpar a tela antes de voltar ao menu principal
    Write-Host "`nOperacao concluida. Pressione qualquer tecla para voltar ao menu principal..." -ForegroundColor Cyan
    [void][System.Console]::ReadKey($true)
    Clear-Host
}





# Funcao para baixar aplicativos utilitarios e salvar em C:\_HTI\Utilitarios - OK
function Download-Utilities {
	Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "         Download de Aplicativos Utilitarios de Suporte" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------" -ForegroundColor DarkGray
	# Confirmar a intencao do usuario
    $confirmDownload = Read-Host "`nTem certeza que deseja baixar os aplicativos utilitarios? (S/N)"
    if ($confirmDownload -notmatch "^(Sim|sim|S|s)$") {
        Write-Host "`nOperacao cancelada. Retornando ao menu principal ..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
		Clear-Host
        Main
        return
    }
	# Caminho do arquivo de lista de URLs
	$fileListUrl = "https://setup.horus.net.br/Scripts/files.txt"
	$UtilitariosFolder = "C:\_HTI\Utilitarios"

	# Verificar se a pasta C:\_HTI\Utilitarios existe, se nao, criar
	if (-not (Test-Path $UtilitariosFolder)) {
		New-Item -Path $UtilitariosFolder -ItemType Directory -Force | Out-Null
		Write-Host "A pasta C:\_HTI\Utilitarios foi criada."
	}

	# Baixar o arquivo de lista de URLs
	Write-Host "Baixando a lista de arquivos..."
	Invoke-WebRequest -Uri $fileListUrl -OutFile "$UtilitariosFolder\files.txt"

	# Ler cada linha do arquivo files.txt e baixar os arquivos
	$files = Get-Content "$UtilitariosFolder\files.txt"
	foreach ($fileUrl in $files) {
		$fileName = [System.IO.Path]::GetFileName($fileUrl)
		$outputPath = Join-Path -Path $UtilitariosFolder -ChildPath $fileName

		Write-Host "Baixando: $fileName..."
		Invoke-WebRequest -Uri $fileUrl -OutFile $outputPath

		Write-Host "$fileName baixado com sucesso."
	}

	Write-Host "`nTodos os aplicativos utilitarios foram baixados com sucesso!`n" -ForegroundColor Green
	Write-Host "Estao disponiveis na pasta C:\_HTI\Utilitarios`n" -ForegroundColor Green
	pause
	Clear-Host
}

function Install-Office {
    try {
        # Perguntar ao usuario se deseja instalar o Office
        Write-Host "`nDeseja instalar o Office neste computador? [S/n]" -ForegroundColor Yellow
        $confirm = Read-Host "(Pressione Enter para SIM)"
        if ($confirm -and $confirm -notmatch "^[Ss]$") {
            Write-Host "`nOperacao cancelada pelo usuario. Retornando ao menu principal..." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Clear-Host
            return
        }

        # Forcar o encerramento dos processos relacionados ao Office
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

                # Encerrar os processos relacionados ao Office
                $officeProcesses | ForEach-Object {
                    try {
                        Stop-Process -Id $_.Id -Force
                        Write-Host "Processo '$($_.Name)' encerrado com sucesso." -ForegroundColor Green
                    } catch {
                        Write-Error "Erro ao encerrar o processo '$($_.Name)': $_"
                    }
                }
            } else {
                Write-Host "Nenhum processo relacionado ao Office foi encontrado." -ForegroundColor Green
            }
        } catch {
            Write-Error "Erro ao verificar ou encerrar processos relacionados ao Office: $_"
        }

        # Verificar e criar o diretorio C:\_HTI\Programas, se necessario
        $programsFolder = "C:\_HTI\Programas"
        if (-not (Test-Path $programsFolder)) {
            Write-Host "`nDiretorio $programsFolder nao encontrado. Criando..." -ForegroundColor Cyan
            New-Item -ItemType Directory -Path $programsFolder -Force | Out-Null
            Write-Host "Diretorio $programsFolder criado com sucesso." -ForegroundColor Green
        }

        # Baixar o instalador do Office
        $officeInstallerPath = Join-Path $programsFolder "OfficeSetup.exe"
        $officeInstallerUrl = "https://setup.horus.net.br/Programas/OfficeSetup.exe"

        Write-Host "`nBaixando o instalador do Office..." -ForegroundColor Cyan
        try {
            Invoke-WebRequest -Uri $officeInstallerUrl -OutFile $officeInstallerPath -UseBasicParsing -ErrorAction Stop
            Write-Host "Instalador do Office baixado com sucesso para '$officeInstallerPath'." -ForegroundColor Green
        } catch {
            Write-Error "Erro ao baixar o instalador do Office: $_"
            return
        }

        # Executar o instalador do Office
        Write-Host "`nIniciando a instalacao do Office..." -ForegroundColor Yellow
        try {
            Start-Process -FilePath $officeInstallerPath
            Write-Host "Continue com a instalacao manual do Office" -ForegroundColor Green
        } catch {
            Write-Error "Erro ao executar o instalador do Office: $_"
        }
    } catch {
        Write-Error "Ocorreu um erro durante a instalacao do Office: $_"
    }

    Write-Host ""
    pause
    Clear-Host
}

function Activate-WindowsOffice {
    try {
        # Perguntar ao usuario se deseja iniciar o processo de ativacao
        Write-Host "`nDeseja iniciar o processo de ativacao do Windows e Office? [S/n]" -ForegroundColor Yellow
        $confirm = Read-Host "(Pressione Enter para SIM)"
        if ($confirm -and $confirm -notmatch "^[Ss]$") {
            Write-Host "`nOperacao cancelada pelo usuario. Retornando ao menu principal..." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Clear-Host
            return
        }

        # Executar o comando de ativacao
        Write-Host "`nIniciando o processo de ativacao do Windows e Office..." -ForegroundColor Cyan
        try {
            irm https://get.activated.win | iex
            Write-Host "`nO processo de ativacao foi iniciado." -ForegroundColor Green
        } catch {
            Write-Error "Erro ao executar o processo de ativacao: $_"
            return
        }

        # Mensagem final ao usuario
        Write-Host "`nFinalize manualmente a ativacao conforme as instrucoes apresentadas." -ForegroundColor Yellow
        Write-Host "`nPressione qualquer tecla para retornar ao menu principal..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Write-Error "Ocorreu um erro durante o processo de ativacao: $_"
    }
    Clear-Host
}

# Funcao para remover aplicativos desnecessarios
function Remove-Apps {
    # Verifica se o Winget esta instalado
    if (-not (Check-WingetInstalled)) {
        Write-Host "`nO Winget nao esta instalado neste computador. Por favor, instale-o antes de prosseguir." -ForegroundColor Red
        pause
		Clear-Host
		return
    }
	Accept-WingetTerms
    # Lista de aplicativos a serem removidos
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

    # Obtem a lista de aplicativos instalados usando Winget
    $wingetOutput = winget list | ForEach-Object {
        $fields = $_ -split '\s{2,}'
        [PSCustomObject]@{
            Name = $fields[0]
            Id = if ($fields.Count -gt 1) { $fields[1] } else { "" }
        }
    }

    # Filtra os aplicativos a serem removidos
    $installedApps = $wingetOutput | Where-Object {
        $appsToRemove -contains $_.Name -or ($appsToRemove | ForEach-Object { $_.ToLower() }) -contains $_.Name.ToLower()
    }

    if ($installedApps.Count -eq 0) {
        Write-Host "`nNenhum dos aplicativos especificados esta instalado no sistema.`n" -ForegroundColor Yellow
		pause
		Clear-Host
        return
    }

    # Exibe os aplicativos que serao removidos
    Write-Host "`nOs seguintes aplicativos serao removidos:`n" -ForegroundColor Cyan
    $installedApps | ForEach-Object { Write-Host $_.Name }

    # Confirmacao do usuario
    $confirmation = Read-Host "Pressione [Enter] para continuar ou digite 'N' para cancelar"
    if ($confirmation -eq 'N') {
        Write-Host "`nA operacao foi cancelada.`n" -ForegroundColor Yellow
		pause
		Clear-Host
        return
    }

    # Remove os aplicativos
    foreach ($app in $installedApps) {
        Write-Host "Removendo $($app.Name)..." -ForegroundColor Green
        try {
            # Remove usando apenas o nome
            winget uninstall --name "$($app.Name)" --silent
            Write-Host "$($app.Name) foi removido com sucesso.`n" -ForegroundColor Green
        } catch {
            Write-Host "Falha ao remover $($app.Name). Verifique manualmente." -ForegroundColor Red
        }
    }

    Write-Host "Todos os aplicativos especificados foram processados." -ForegroundColor Cyan
	pause
	Clear-Host
	return
}




# Funcao para mostrar os usuarios locais ativos do Windows - OK
function Show-ActiveLocalUsers {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "         Lista dos Usuarios Locais e Ativos do Windows" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    # Obter usuarios locais ativos
    $localUsers = Get-LocalUser | Where-Object { $_.Enabled -eq $true }

    # Exibir cada usuario com destaque em azul
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

# Funcao para mostrar os usuarios locais do grupo administradores
function Show-LocalAdminUsers {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "            Usuarios do Grupo Administradores" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    # Verifica se o computador esta em um dominio ou em um grupo de trabalho
    $domain = (Get-WmiObject Win32_ComputerSystem).Domain
    $isDomainJoined = $domain -ne (Get-WmiObject Win32_ComputerSystem).Name

    if ($isDomainJoined) {
        # Se o computador esta no dominio, usa o comando 'net localgroup'
        Write-Host "`nO computador esta no dominio: $domain" -ForegroundColor Yellow
        Write-Host "Listando membros usando 'net localgroup'`n" -ForegroundColor Green
        cmd /c "net localgroup Administradores"
    } else {
        # Caso contrario, usa o metodo tradicional
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


# Funcao para criar um novo usuario local
function Create-NewLocalUser {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "          CRIAR NOVO USUaRIO LOCAL" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green

    # Solicitar o nome do novo usuario
    do {
        Write-Host ""
        $username = Read-Host "Digite o nome do novo usuario (ou pressione '0' para voltar ao menu anterior)"

        if (-not $username) {
            Write-Host "`nO nome do usuario nao pode estar em branco. Tente novamente ou pressione '0' para voltar ao menu anterior." -ForegroundColor Red
            continue
        }

        if ($username -eq "0") {
            Write-Host "`nRetornando ao menu anterior..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Manage-LocalUsers
            return
        }

        if ($username -notmatch '^[a-zA-Z0-9._-]+$') {
            Write-Host "`nO nome do usuario contem caracteres invalidos. Use apenas letras, numeros, '.', '-' ou '_'." -ForegroundColor Red
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

    # Solicitar e confirmar a senha do novo usuario
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

    # Criar o usuario local
    New-LocalUser -Name $username -Password $password | Out-Null

    # Perguntar se deseja adicionar ao grupo Administradores
    $addToAdminsGroup = Read-Host "Deseja adicionar o usuario ao grupo Administradores? (S/N)"
    if ($addToAdminsGroup -eq "S" -or $addToAdminsGroup -eq "s") {
        Add-LocalGroupMember -Group "Administradores" -Member $username
        Write-Host "`nUsuario '$username' criado e adicionado ao grupo Administradores." -ForegroundColor Green
    } else {
        # Adicionar ao grupo local de usuarios padrao
        $GUsers = Get-LocalGroup | Where-Object { $_.SID -eq 'S-1-5-32-545' }
        Add-LocalGroupMember -Group $GUsers -Member $username
        Write-Host "`nUsuario '$username' criado no grupo Local padrao." -ForegroundColor Green
    }

    # Definir para que a senha nunca expire
    Set-LocalUser -Name $username -PasswordNeverExpires 1

    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "             PROCESSO FINALIZADO" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green

    # Retornar para a funcao principal
    pause
    Manage-LocalUsers
}

# Funcao para criar um usuario HTI
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

    # Dupla verificacao da senha
    do {
        $password = Read-Host "Digite a senha do usuario HTI" -AsSecureString
        $passwordConfirmation = Read-Host "Confirme a senha digitada" -AsSecureString

        # Converte as SecureString para texto simples para comparar
        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
        $plainPasswordConfirmation = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordConfirmation))

        if ($plainPassword -ne $plainPasswordConfirmation) {
            Write-Host "`nAs senhas digitadas nao coincidem. Tente novamente." -ForegroundColor Red
        } else {
            Write-Host "`nSenha confirmada com sucesso!" -ForegroundColor Green
            break
        }
    } while ($true)

    # Criacao do usuario e inclusao no grupo Administradores
    $user = New-LocalUser -Name $username -Password $password
    Add-LocalGroupMember -Group "Administradores" -Member $username
    Write-Host "`nUsuario $username criado no grupo Administradores`n" -ForegroundColor Green
    Set-LocalUser -Name $username -PasswordNeverExpires 1

    # Pergunta se deseja ativar a Area de Trabalho Remota
    $enableRDP = Read-Host "Deseja ativar a Area de Trabalho Remota do Windows? (S/N)"
    if ($enableRDP -match "^[sS]$") {
        try {
            # Ativando o servico RDP
            Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0

            # Verificando e habilitando as regras de firewall relacionadas ao RDP
            $rdpRules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "*Remote Desktop*" }
            if ($rdpRules) {
                $rdpRules | ForEach-Object { Enable-NetFirewallRule -Name $_.Name }
                Write-Host "`nRegras de firewall do RDP ativadas com sucesso!" -ForegroundColor Green
            } else {
                Write-Host "`nArea de Trabalho Remota foi ativada, mas..." -ForegroundColor Yellow
				Write-Host "`nNao foi possivel localizar as regras de firewall do RDP. Configure-as manualmente." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "`nErro ao ativar a Area de Trabalho Remota." -ForegroundColor Red
            Write-Host "Verifique suas permissoes ou configuracoes do sistema." -ForegroundColor Yellow
        }
    } else {
        Write-Host "`nArea de Trabalho Remota nao foi ativada." -ForegroundColor Yellow
    }

    # Pausa e retorno ao menu principal
    Start-Sleep -Seconds 2
	pause
    Manage-LocalUsers
}


# Funcao para trocar a senha de um usuario local
function Change-LocalUserPassword {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "            Trocar a Senha de um Usuario Local" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Output "`nQual usuario abaixo deseja trocar a senha?"
    
    # Exibe os usuarios locais habilitados
    $localUsers = Get-LocalUser | Where-Object { $_.Enabled -eq $true }
    $localUsers | Select-Object -ExpandProperty Name
    Write-Host "`nSe deseja remover a senha de um usario, eh so deixar em branco e confirmar" -ForegroundColor Yellow
    # Solicita o nome do usuario
    $username = Read-Host "`nDigite o nome do usuario para alterar a senha"
    $userExists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
    if (-not $userExists) {
        Write-Host "O usuario '$username' nao existe." -ForegroundColor Yellow
        Write-Host "`nRetornando ao menu anterior ..."
        Start-Sleep -Seconds 3
        Manage-LocalUsers
        return
    }
    
    # Dupla verificacao da nova senha
    do {
        $newPassword = Read-Host "Digite a nova senha" -AsSecureString
        $passwordConfirmation = Read-Host "Confirme a nova senha" -AsSecureString

        # Converte as senhas temporariamente para texto simples para comparacao
        $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringUni(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPassword)
        )
        $plainConfirmation = [System.Runtime.InteropServices.Marshal]::PtrToStringUni(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordConfirmation)
        )

        if ($plainPassword -ne $plainConfirmation) {
            Write-Host "`nAs senhas digitadas nao coincidem. Tente novamente." -ForegroundColor Red
        } else {
            Write-Host "`nSenha confirmada com sucesso!" -ForegroundColor Green
            break
        }
    } while ($true)
    
    # Altera a senha do usuario
    Set-LocalUser -Name $username -Password $newPassword
    Write-Host "`nSenha do usuario '$username' alterada com sucesso!" -ForegroundColor Green
    # Retornar para a funcao principal
    pause
    Manage-LocalUsers
}

# Funcao para habilitar, desabilitar ou remover um usuario local
function Enable-Disable-Remove-LocalUser {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "    Habilitar, Desabilitar ou Remover um Usuario Local" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Output "`nSelecione um dos usuarios abaixo:`n"
    
    # Listar usuarios locais com status
    $allUsers = Get-LocalUser
    foreach ($user in $allUsers) {
        $status = if ($user.Enabled) { "Habilitado" } else { "Desabilitado" }
        Write-Host "$($user.Name) ... $status"
    }
    
    # Solicitar nome do usuario
    $username = Read-Host "`nDigite o nome do usuario para habilitar, desabilitar ou remover"
    $userExists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
    if (-not $userExists) {
        Write-Host "`nO usuario '$username' nao existe." -ForegroundColor Yellow
        Write-Host "`nRetornando ao menu anterior ..."
        Start-Sleep -Seconds 2
        Manage-LocalUsers
        return
    }
    
    # Exibir opcoes
    $hostname = $env:COMPUTERNAME
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
            # Confirmacao para remover usuario
            Write-Host "`nATENCAO: Remover um usuario e uma acao irreversivel." -ForegroundColor Red
            Write-Host "Isso pode causar perda de informacoes associadas ao usuario '$username'."
            $confirmRemove = Read-Host "`nTem certeza que deseja remover o usuario '$username'? Digite 'Sim' para confirmar ou pressione Enter para cancelar"
            if ($confirmRemove -ne "Sim") {
                Write-Host "`nOperacao cancelada. O usuario '$username' NAO foi removido." -ForegroundColor Yellow
                pause
                Manage-LocalUsers
                return
            }
            
            # Remover o usuario
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

# Funcao para adicionar ou remover um usuario do grupo Administradores
function Add-Remove-UserFromAdminGroup {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "Adicionar ou Remover Usuario do Grupo Administradores" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    # Verifica se o computador esta em um dominio ou em um grupo de trabalho
    $domain = (Get-WmiObject Win32_ComputerSystem).Domain
    $isDomainJoined = $domain -ne (Get-WmiObject Win32_ComputerSystem).Name

    if ($isDomainJoined) {
        # Se o computador esta no dominio, usa o comando 'net localgroup'
        Write-Host "`nO computador esta no dominio: $domain" -ForegroundColor Yellow
        Write-Host "`nListando os Usuarios no Grupo Administratores'`n" -ForegroundColor Green
        cmd /c "net localgroup Administradores"
    } else {
        # Caso contrario, usa o metodo tradicional
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
            Write-Host "Verifique se voce tem permissoes administrativas." -ForegroundColor Yellow
        }
    }
	
    # Solicita o nome do usuario
    $username = Read-Host "`nDigite o nome do usuario"

    if (-not $isDomainJoined) {
        # Verifica se o usuario local existe
        $userExists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
        if (-not $userExists) {
            Write-Host "`nO usuario '$username' nao existe.`n" -ForegroundColor Red
			pause
            return
        }
    }

    # Adicionar ou remover o usuario do grupo Administradores
    $hostname = $env:COMPUTERNAME
    $choice = Read-Host "Pressione 'A' para Adicionar ou 'R' para Remover"
    switch ($choice.ToUpper()) {
        "A" {
            if ($isDomainJoined) {
                cmd /c "net localgroup Administradores $username /add"
                if ($?) {
                    Write-Host "O usuario '$username' foi adicionado ao grupo Administradores." -ForegroundColor Green
					Write-Host "`nRetornando ao menu anterior ...`n"
					pause
					Manage-LocalUsers
                } else {
                    Write-Host "Erro ao adicionar '$username' ao grupo Administradores." -ForegroundColor Red
					Write-Host "`nRetornando ao menu anterior ...`n"
					pause
					Manage-LocalUsers
                }
            } else {
                # Verifica se o usuario ja e membro
                $isAdmin = Get-LocalGroupMember Administradores | Where-Object { $_.Name -eq "$hostname\$username" }
                if ($isAdmin) {
                    Write-Host "O usuario '$username' ja pertence ao grupo Administradores." -ForegroundColor Yellow
					Write-Host "`nRetornando ao menu anterior ...`n"
					pause
					Manage-LocalUsers
                } else {
                    Add-LocalGroupMember -Group "Administradores" -Member "$hostname\$username"
                    Write-Host "O usuario '$username' foi adicionado ao grupo Administradores." -ForegroundColor Green
					Write-Host "`nRetornando ao menu anterior ...`n"
					pause
					Manage-LocalUsers
                }
            }
            break
        }
        "R" {
            if ($isDomainJoined) {
                cmd /c "net localgroup Administradores $username /delete"
                if ($?) {
                    Write-Host "O usuario '$username' foi removido do grupo Administradores." -ForegroundColor Blue
					Write-Host "`nRetornando ao menu anterior ...`n"
					pause
					Manage-LocalUsers
                } else {
                    Write-Host "Erro ao remover '$username' do grupo Administradores." -ForegroundColor Red
					Write-Host "`nRetornando ao menu anterior ...`n"
					pause
					Manage-LocalUsers
                }
            } else {
                $isAdmin = Get-LocalGroupMember Administradores | Where-Object { $_.Name -eq "$hostname\$username" }
                if ($isAdmin) {
                    Remove-LocalGroupMember -Group "Administradores" -Member "$hostname\$username"
                    Write-Host "O usuario '$username' foi removido do grupo Administradores." -ForegroundColor Blue
					Write-Host "`nRetornando ao menu anterior ...`n"
					pause
					Manage-LocalUsers
                } else {
                    Write-Host "O usuario '$username' ja nao pertence ao grupo Administradores." -ForegroundColor Yellow
					Write-Host "`nRetornando ao menu anterior ...`n"
					pause
					Manage-LocalUsers
                }
            }
            break
        }
        default {
            Write-Host "Opcao invalida." -ForegroundColor Red
			Write-Host "`nRetornando ao menu anterior ..."
			Start-Sleep -Seconds 2
			Manage-LocalUsers
            break
        }
    }
}


# Funcao para adicionar ou remover um usuario do grupo Administradores
function Add-Remove-UserFromAdminGroup22 {
    Clear-Host
	Write-Output "`nQual usuario abaixo deseja Adicionar ou Remover do grupo Administradores?"
    #$localUsers = Get-WmiObject Win32_UserAccount | Where-Object { $_.LocalAccount -eq $true -and $_.Disabled -eq $false }
	$localUsers = Get-LocalUser | Where-Object { $_.Enabled -eq $true }
    $localUsers | Select-Object -ExpandProperty Name
	$username = Read-Host "Digite o nome do usuario"
    $userExists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
    if (-not $userExists) {
        Write-Host "O usuario '$username' nao existe."
        return
    }

    $hostname = $env:COMPUTERNAME
	$choice = Read-Host "Pressione 'A' para Adicionar ou 'R' para Remover"
    switch ($choice.ToUpper()) {
        "A" {
            Add-LocalGroupMember -Group "Administradores" -Member "$hostname\$username"
			Write-Host "O usuario '$username' foi adicionado ao grupo Administradores." -ForegroundColor Green
            break
        }
        "R" {
            $isAdmin = Get-LocalGroupMember Administradores | Where-Object {$_.Name -eq "$hostname\$username"}
            if ($isAdmin) {
                Remove-LocalGroupMember -Group "Administradores" -Member "$hostname\$username"
				Write-Host "O usuario '$username' removido do grupo Administradores." -ForegroundColor Blue
            } else {
                Write-Host "O usuario '$username' ja nao pertence ao grupo Administradores." -ForegroundColor Yellow
            }
            break
        }
        default {
            Write-Host "Opcao invalida."
            break
        }
    }
}

# Funcao para verificar a saude do SSD
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
        Write-Error "Nao foi possivel obter informacoes sobre os discos fisicos: $_"
    }
	pause
	Disk-ManagementMenu
}

# Funcao para verificar e exibir o uso do armazenamento
function Check-DiskUsage {
    param (
        [string]$DriveLetter = "C:",       # Unidade padrao
        [int]$Threshold = 00              # Limite de uso em percentual
    )

    try {
        # Verificar se a unidade existe
        $drive = Get-PSDrive -Name $DriveLetter.TrimEnd(':') -ErrorAction Stop

        # Obter nome do computador, endereco IP e usuario
        $computerName = $env:COMPUTERNAME
        $ipAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike '*Loopback*' }).IPAddress
        $currentUser = $env:USERNAME

        # Calcular o percentual de uso do disco
        $usedSpacePercent = [math]::Round((($drive.Used / ($drive.Free + $drive.Used)) * 100), 2)

        # Exibir informacoes na tela
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

        # Verificar se o percentual de uso esta acima do limite
        if ($usedSpacePercent -ge $Threshold) {
            Write-Host "ALERTA: O disco $DriveLetter esta com ${usedSpacePercent}% de uso, acima do limite de ${Threshold}%!" -ForegroundColor Red
        }

        # Perguntar se deseja enviar por e-mail
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
            # Chamar funcao para enviar e-mail
            Send-Email -Subject $subject -Body $body
        } else {
            Write-Host "Relatorio nao sera enviado por e-mail." -ForegroundColor Yellow
			Start-Sleep -Seconds 2
			Disk-ManagementMenu
        }

    } catch {
        Write-Host "A unidade '$DriveLetter' nao foi encontrada." -ForegroundColor Red
        Write-Host "Unidade invalida. Tente novamente." -ForegroundColor Yellow
    }
}

# Funcao para exibir unidades disponiveis e solicitar entrada do usuario
function Select-Drive {
    Write-Host "`nUnidades disponiveis no sistema:" -ForegroundColor Cyan

    # Obter todas as unidades de sistema de arquivos
    $drives = Get-PSDrive | Where-Object { $_.Provider -like "*FileSystem*" }

    # Verificar se existem unidades disponiveis
    if ($drives.Count -eq 0) {
        Write-Host "Nenhuma unidade de disco foi encontrada no sistema." -ForegroundColor Red
        return $null
    }

    # Exibir unidades disponiveis
    $drives | ForEach-Object {
        Write-Host " - $($_.Name):\ ($([math]::Round($_.Used / 1GB, 2)) GB usados / $([math]::Round($_.Free / 1GB, 2)) GB livres)" -ForegroundColor Green
    }

	# Solicitar entrada do usuario
	$defaultDrive = "C"
	$driveLetter = Read-Host "`nDigite a letra da unidade que deseja verificar (Exemplo: C, D) [Default: $defaultDrive]"

	# Se o usuario pressionar "Enter" sem digitar nada, usa o valor padrao
	if ([string]::IsNullOrWhiteSpace($driveLetter)) {
		$driveLetter = $defaultDrive
	}

    # Garantir que a letra da unidade esteja formatada corretamente
    #$driveLetter = $driveLetter.TrimEnd(':')

    # Verificar se a unidade informada esta nas unidades de sistema de arquivos
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
        # Verificar se a pasta C:\_HTI\Scripts existe; se nao, criar
        $scriptFolder = "C:\_HTI\Scripts"
        if (-not (Test-Path $scriptFolder)) {
            New-Item -ItemType Directory -Path $scriptFolder -Force | Out-Null
        }

        # Verificar se o script MonitorDisk.ps1 existe; se nao, baixar da URL
        $scriptPath = Join-Path $scriptFolder "MonitorDisk.ps1"
        if (-not (Test-Path $scriptPath)) {
            Write-Host "`nScript MonitorDisk.ps1 nao encontrado. Baixando da URL..." -ForegroundColor Yellow
            $url = "https://setup.horus.net.br/Scripts/MonitorDisk.ps1"
            try {
                Invoke-WebRequest -Uri $url -OutFile $scriptPath -UseBasicParsing
                Write-Host "Script MonitorDisk.ps1 baixado e salvo em '$scriptPath'." -ForegroundColor Green
            } catch {
                Write-Error "Falha ao baixar o script MonitorDisk.ps1: $_"
                return
            }
        }

        # Perguntar ao usuario se deseja criar um agendamento
        Write-Host "`nDeseja criar um agendamento para monitorar o uso do HD? [S/n]" -ForegroundColor Yellow
        $confirm = Read-Host "(Pressione Enter para SIM)"
        if ($confirm -and $confirm -notmatch "^[Ss]$") {
            Write-Host "`nOperacao cancelada pelo usuario. Retornando ao menu principal..." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Disk-ManagementMenu
            return
        }

        # Selecionar frequencia de agendamento
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

        # Informar configuracao padrao
        Write-Host "`nATENcaO" -ForegroundColor Yellow
        Write-Host "O monitoramento esta pre-configurado para o disco C: com limite de 90%." -ForegroundColor Cyan
        Write-Host "Se deseja alterar esses valores, edite diretamente o script em $scriptPath." -ForegroundColor Cyan

        # Nome da tarefa
        $taskName = "Monitoramento de Disco"
        $taskFolder = "_HTI" # Nome da pasta no Agendador de Tarefas

        # Criar a pasta no Agendador de Tarefas
        $folderPath = "\$taskFolder"
        $service = New-Object -ComObject "Schedule.Service"
        $service.Connect()
        $rootFolder = $service.GetFolder("\")
        try {
            $rootFolder.GetFolder($folderPath)
        } catch {
            $rootFolder.CreateFolder($folderPath)
        }

        # Caminho completo da tarefa
        $taskPath = "\$taskFolder\$taskName"

        # Verificar se a tarefa ja existe
        if (Get-ScheduledTask -TaskName $taskName -TaskPath "\$taskFolder\" -ErrorAction SilentlyContinue) {
            Write-Host "`nJa existe uma tarefa chamada '$taskName' na pasta '$taskFolder'." -ForegroundColor Yellow
            $overwrite = Read-Host "Deseja substituir a tarefa existente? (S/N)"
            if ($overwrite -notmatch "^[Ss]$") {
                Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Disk-ManagementMenu
                return
            }

            # Remover tarefa existente
            Unregister-ScheduledTask -TaskName $taskName -TaskPath "\$taskFolder\" -Confirm:$false
        }

        # Criar a tarefa com base na frequencia
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
        Write-Error "`nErro ao criar o agendamento: $_"
    }
    Write-Host ""
    pause
    Disk-ManagementMenu
}



# Ajuste na chamada no menu
function Disk-ManagementMenu {
    Clear-Host
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "           Status, Monitoramento e Alertas de Disco           " -ForegroundColor Green
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "Escolha uma opcao:`n" -ForegroundColor Yellow
    Write-Host "1) Verificar a saude dos discos" -ForegroundColor Cyan
    Write-Host "2) Mostrar o Uso do HD e enviar por e-mail" -ForegroundColor Cyan
    Write-Host "3) Criar um agendamento no Windows para monitorar o uso do HD (em breve)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    Write-Host "--------------------------------------------------------------" -ForegroundColor Green

    $choice = Read-Host "Digite o numero da opcao desejada"
    
    switch ($choice) {
        "1" {
            Check-SSDHealth
        }
        "2" {
            # Selecionar unidade
            $driveLetter = Select-Drive
            if ($driveLetter -eq $null) {
                Write-Host "Operacao cancelada." -ForegroundColor Yellow
                return
            }

            # Executar funcao de verificacao
            $threshold = 90  # Limite de uso em percentual
            Check-DiskUsage -DriveLetter $driveLetter -Threshold $threshold
        }
        "3" { Schedule-DiskMonitoring }
        "0" { Clear-Host; return }  # Voltar ao menu principal
        default {
            Write-Host "Opcao invalida." -ForegroundColor Red
        }
    }
}

# Funcao para Enviar E-mail
function Send-Email {
    param (
        [string]$Subject,
        [string]$Body
    )
    
    # Configuracoes do alerta de e-mail
    $smtpServer = "smtp.zoho.com"            # Substitua pelo seu servidor SMTP
    $smtpPort = 587                          # Porta SMTP
    $smtpUser = "prb2020@zohomail.com"       # Seu e-mail
    $smtpPassword = "MM4FT/rq4JaGEu&"        # Sua senha
    $recipient = "suporte@horusti.com.br"    # Destinatario
    $sender = "prb2020@zohomail.com"         # Remetente

    # Criar cliente de SMTP
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
    try {
        $smtpClient.Send($mailMessage)
        Write-Host "E-mail enviado com sucesso!" -ForegroundColor Green
		Start-Sleep -Seconds 2
		Clear-Host
    } catch {
        Write-Error "Erro ao enviar o e-mail: $_"
		Start-Sleep -Seconds 2
		return
    }
}

# Definir a funcao de gerenciamento de usuarios locais
function Manage-LocalUsers {
    Clear-Host
	Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "              Gerenciador de Usuarios Locais" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Green

    # Mostrar opcoes do menu de gerenciamento de usuarios locais
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
        "0" { Clear-Host; return }  # Voltar ao menu principal
        default { Write-Host "Opcao invalida. Por favor, escolha novamente." }
    }
}

# Funcao para ajustar configuracoes de energia (PowerUP)
function PowerUP {
    # Definir variaveis para tempos de energia
    $monitorTimeoutDC = 5   # Tempo para desligar monitor na bateria (minutos)
    $standbyTimeoutDC = 10  # Tempo para suspender na bateria (minutos)
    $monitorTimeoutAC = 60  # Tempo para desligar monitor na energia (minutos)
    $standbyTimeoutAC = 0   # Tempo para suspender na energia (nunca)

    $isLaptop = (Get-WmiObject -Class Win32_SystemEnclosure).ChassisTypes -match "8|9|10|11|14|18|21|30"
    
    if ($isLaptop) {
        Write-Output "Dispositivo identificado como Notebook."
        
        # Configuracoes para quando estiver na bateria
        Start-Process -FilePath "powercfg" -ArgumentList "/change monitor-timeout-dc $monitorTimeoutDC" -Verb RunAs -Wait
        Start-Process -FilePath "powercfg" -ArgumentList "/change standby-timeout-dc $standbyTimeoutDC" -Verb RunAs -Wait
        
        # Configuracoes para quando estiver conectado na energia
        Start-Process -FilePath "powercfg" -ArgumentList "/change monitor-timeout-ac $monitorTimeoutAC" -Verb RunAs -Wait
        Start-Process -FilePath "powercfg" -ArgumentList "/change standby-timeout-ac $standbyTimeoutAC" -Verb RunAs -Wait
    } else {
        Write-Output "Dispositivo identificado como Desktop."
        
        # Ajuste de energia para PC
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

    # Verifica se ha atualizacoes pendentes
    $updates = Get-CimInstance -Namespace "root\ccm\clientSDK" -ClassName "CCM_SoftwareUpdate" -ErrorAction SilentlyContinue

    if ($updates.Count -eq 0) {
        Write-Host "Nenhuma atualizacao disponivel no momento." -ForegroundColor Green
    } else {
        Write-Host "Atualizacoes disponiveis:" -ForegroundColor Yellow
        $updates | ForEach-Object { Write-Host "- $($_.ArticleID)" -ForegroundColor White }

        # Pergunta ao usuario se deseja instalar
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

    # Retorna ao menu
    Pause
    Otimizacoes-Win
}


function Clear-PrintSpoolerQueue {
    [CmdletBinding()]
    param()

    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host "INICIANDO LIMPEZA DA FILA DE IMPRESSAO" -ForegroundColor Yellow
    Write-Host "==========================================================="

    # O bloco "Finally" garante que a tentativa de iniciar o servico sempre ocorra,
    # mesmo que o script encontre um erro. Isso evita deixar o Spooler parado.
    try {
        # Etapa 1: Parar o servico de Spooler de Impressao
        Write-Host "`n[ETAPA 1] Parando o servico de Spooler (Spooler)..." -ForegroundColor Cyan
        # Usamos -Force para ajudar a parar o servico mesmo que ele esteja "travado".
        # -ErrorAction Stop garante que qualquer erro aqui sera capturado pelo bloco Catch.
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Write-Host "Servico de Spooler parado com sucesso." -ForegroundColor Green

        # Etapa 2: Deletar os arquivos da fila
        # O caminho padrao e C:\Windows\System32\spool\PRINTERS
        $spoolPath = Join-Path -Path $env:SystemRoot -ChildPath "System32\spool\PRINTERS"
        Write-Host "`n[ETAPA 2] Limpando arquivos da fila em '$spoolPath'..." -ForegroundColor Cyan
        
        # Remove todos os arquivos e pastas dentro do diretorio PRINTERS
        Remove-Item -Path "$spoolPath\*" -Recurse -Force -ErrorAction Stop
        Write-Host "Fila de impressao limpa com sucesso." -ForegroundColor Green
    }
    catch {
        # Captura qualquer erro que ocorreu no bloco Try (ex: acesso negado)
        Write-Error "Ocorreu um erro durante a limpeza da fila de impressao."
        Write-Error $_.Exception.Message
    }
    finally {
        # Etapa 3: Reiniciar o servico de Spooler
        # Este bloco e executado SEMPRE, tendo o Try funcionado ou nao.
        Write-Host "`n[ETAPA 3] Iniciando o servico de Spooler..." -ForegroundColor Cyan
        
        # Verifica se o servico nao esta rodando antes de tentar inicia-lo
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
	# Confirmacao do usuario
	$confirmacao = Read-Host "`nEsta acao vai limpar o cache do Teams e desconectar a conta. Deseja continuar? (S/N) [S]"

	if ($confirmacao -eq "N" -or $confirmacao -eq "n") {
		Write-Host "Operacao cancelada pelo usuario." -ForegroundColor Yellow
		pause
		Otimizacoes-Win
	}

	Write-Host "`nIniciando limpeza de cache do Microsoft Teams..." -ForegroundColor Cyan

	# Tenta encerrar qualquer processo do Teams (nova e antiga versao)
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

	# Caminhos das versoes
	$legacyPath = "$env:APPDATA\Microsoft\Teams"
	$newTeamsPath = "$env:LocalAppData\Packages\MSTeams_8wekyb3d8bbwe\LocalCache"
	$newTeamsAltPath = "$env:LocalAppData\Microsoft\Teams"

	# Flags
	$found = $false

	# Limpeza da nova versao
	if (Test-Path $newTeamsPath) {
		Write-Host "Nova versao do Teams detectada!" -ForegroundColor Cyan
		Remove-Item "$newTeamsPath\*" -Recurse -Force -ErrorAction SilentlyContinue
		Write-Host "Cache da nova versao limpo com sucesso." -ForegroundColor Green
		$found = $true
	}
	elseif (Test-Path $newTeamsAltPath) {
		Write-Host "Nova versao (alternativa) do Teams detectada!" -ForegroundColor Cyan
		Remove-Item "$newTeamsAltPath\*" -Recurse -Force -ErrorAction SilentlyContinue
		Write-Host "Cache da nova versao (alternativa) limpo com sucesso." -ForegroundColor Green
		$found = $true
	}
	elseif (Test-Path $legacyPath) {
		Write-Host "Versao classica do Teams detectada!" -ForegroundColor Cyan

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

	# Reinicio do Teams
	if (-not $found) {
		Write-Host "Nao foi possivel localizar os arquivos de cache do Teams." -ForegroundColor Red
		Write-Host "Verifique se o Teams esta instalado corretamente." -ForegroundColor DarkRed
	} else {
		Write-Host "`nTentando reiniciar o Microsoft Teams..." -ForegroundColor Cyan

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
			pause
			Otimizacoes-Win
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
	
    # Verifica servico, processo e instalacao no registro
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
        try {
            $versao = (Get-Item (Join-Path $anydeskPath "AnyDesk.exe")).VersionInfo.ProductVersion.Trim()
        } catch { }
    }

    if (-not $versao -and $processo) {
        try {
            $versao = ($processo | Select-Object -First 1).FileVersion.Trim()
        } catch { }
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
    Write-Host "2) Baixa e executa AnyDesk versao 7 - necessita instalar"
    Write-Host "3) Remover AnyDesk do computador"
	Write-Host "0) Voltar ao menu principal" -ForegroundColor DarkGreen
    $opcao = Read-Host "`nDigite o numero da opcao desejada"

    switch ($opcao) {
        "1" {
            if ($servico -or $processo) {
                $confirmacao = Read-Host "`nAnyDesk ja esta instalado e em execucao. Deseja reinstalar? (S/N) [Padrao: N]"
                if ($confirmacao.ToUpper() -ne "S") {
                    Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Yellow
                    pause
					Gerenciar-AnyDesk
                }
            }
            Write-Host "`nInstalando ultima versao do AnyDesk via Winget..." -ForegroundColor Cyan
            winget install AnyDesk -e --silent
            pause
            Otimizacoes-Win
        }
        "2" {
            if ($servico -or $processo) {
                $confirmacao = Read-Host "`nAnyDesk ja esta instalado e em execucao. Deseja reinstalar a versao 7? (S/N) [Padrao: N]"
                if ($confirmacao.ToUpper() -ne "S") {
                    Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Yellow
                    pause
					Gerenciar-AnyDesk
                }
            }

            # Verificar e criar o diretorio C:\_HTI\Utilitarios, se necessario
            $utilFolder = "C:\_HTI\Utilitarios"
            if (-not (Test-Path $utilFolder)) {
                Write-Host "`nDiretorio $utilFolder nao encontrado. Criando..." -ForegroundColor Cyan
                New-Item -ItemType Directory -Path $utilFolder -Force | Out-Null
                Write-Host "Diretorio $utilFolder criado com sucesso." -ForegroundColor Green
            }
            
            $arquivo = Join-Path $utilFolder "AnyDesk-v7.exe"
            $anydeskInstallerUrl = "https://setup.horus.net.br/Utilitarios/AnyDesk-v7.exe"

            Write-Host "`nBaixando o instalador do AnyDesk-v7..." -ForegroundColor Cyan
            try {
                Invoke-WebRequest -Uri $anydeskInstallerUrl -OutFile $arquivo -UseBasicParsing -ErrorAction Stop
                Write-Host "Instalador do AnyDesk-v7 baixado com sucesso para '$arquivo'." -ForegroundColor Green
            } catch {
                Write-Error "Erro ao baixar o instalador do AnyDesk-v7: $_"
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
                Write-Error "Erro ao instalar o AnyDesk: $_"
                pause
                Gerenciar-AnyDesk
            }
        }
        "3" {
            if (-not ($servico -or $processo)) {
                Write-Host "`nAnyDesk nao esta instalado para remover." -ForegroundColor Yellow
                pause
				Gerenciar-AnyDesk
            }

            $confirmacao = Read-Host "`nConfirma a remocao do AnyDesk? (S/N) [Padrao: N]"
            if ($confirmacao.ToUpper() -ne "S") {
                Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Yellow
                return
            }

            Write-Host "`nRemovendo o AnyDesk..." -ForegroundColor Cyan

            if ($servico) { Stop-Service -Name "AnyDesk" -Force -ErrorAction SilentlyContinue }
            if ($processo) { Stop-Process -Name "AnyDesk" -Force -ErrorAction SilentlyContinue }

            & winget uninstall AnyDesk -e --silent --force

            Write-Host "`nRemovendo arquivos residuais..." -ForegroundColor Yellow
            Remove-Item -Recurse -Force "$env:APPDATA\AnyDesk" -ErrorAction SilentlyContinue
            Remove-Item -Recurse -Force "$env:LOCALAPPDATA\AnyDesk" -ErrorAction SilentlyContinue
            Remove-Item -Recurse -Force "C:\ProgramData\AnyDesk" -ErrorAction SilentlyContinue

            Write-Host "`nRemocao concluida!" -ForegroundColor Green
            pause
            Otimizacoes-Win
        }
		"0" { Clear-Host; Otimizacoes-Win }  # Voltar ao menu anterior
        default {
            Write-Host "`nOpcao invalida. Nenhuma acao realizada." -ForegroundColor Red
        }
    }
}

function LogsEventViewer {
	
	# --- Configuracao Inicial ---
Clear-Host
$OutputEncoding = [System.Text.Encoding]::UTF8
[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- Titulo ---
Write-Host "=================================================="
Write-Host "     Analisador de Logs do Windows para TI v8.0"
Write-Host "         (Analisando eventos dos ultimos 7 dias)"
Write-Host "=================================================="
Write-Host

# --- RESUMO EM TEXTO SIMPLES ---
Write-Host "--- Resumo dos Event IDs que serao analisados ---" -ForegroundColor White -BackgroundColor DarkBlue
Write-Host
Write-Host 'Discos:' -ForegroundColor Cyan
Write-Host '  - IDs 7, 11, 51, 55: Indicam erros de bloco/setor defeituoso no disco rigido ou SSD.' -ForegroundColor DarkGray
Write-Host '  - IDs 98, 153: Apontam falhas de comunicacao ou timeout com o dispositivo de armazenamento.' -ForegroundColor DarkGray
Write-Host '  - ID 1001 (Chkdsk): Mostra o resultado da verificacao de integridade do disco (Check Disk).' -ForegroundColor DarkGray
Write-Host
Write-Host 'Drivers:' -ForegroundColor Cyan
Write-Host '  - ID 1 (WHEA-Logger): Erro de hardware generico, frequentemente ligado a drivers instaveis.' -ForegroundColor DarkGray
Write-Host '  - ID 219 (Kernel-PnP): Falha ao carregar um driver para um dispositivo Plug and Play.' -ForegroundColor DarkGray
Write-Host '  - IDs 10110, 10111: Um driver critico parou de funcionar no modo de usuario (Driver Frameworks).' -ForegroundColor DarkGray
Write-Host
Write-Host 'Falhas Criticas e Estabilidade:' -ForegroundColor Cyan
Write-Host '  - ID 41 (Kernel-Power): Sistema reiniciou de forma inesperada (possivel Tela Azul/BSOD). (Critico)' -ForegroundColor DarkGray
Write-Host '  - ID 1001 (BugCheck): Relatorio de erro detalhado que acompanha uma Tela Azul (BSOD).' -ForegroundColor DarkGray
Write-Host '  - ID 1000 (App Error): Uma aplicacao travou e fechou inesperadamente (crash).' -ForegroundColor DarkGray
Write-Host '  - ID 1002 (App Hang): Uma aplicacao parou de responder (congelou).' -ForegroundColor DarkGray
Write-Host
Write-Host 'Saude do Sistema Operacional:' -ForegroundColor Cyan
Write-Host '  - IDs 7000, 7009: Um servico essencial do Windows nao conseguiu iniciar corretamente.' -ForegroundColor DarkGray
Write-Host '  - ID 20 (UpdateClient): Erro durante a tentativa de instalar uma atualizacao do Windows.' -ForegroundColor DarkGray
Write-Host
Write-Host 'Desempenho e Seguranca:' -ForegroundColor Cyan
Write-Host '  - ID 100 (Diagnostics): A inicializacao (boot) do Windows foi considerada lenta.' -ForegroundColor DarkGray
Write-Host '  - ID 4625 (Security): Tentativa de logon com falha (senha ou usuario incorreto). (Requer Admin)' -ForegroundColor DarkGray
Write-Host
Write-Host "-------------------------------------------------" -ForegroundColor White -BackgroundColor DarkBlue
Write-Host

# --- PAUSA PARA CONFIRMAcAO DO USUARIO ---
$escolha = Read-Host -Prompt "Pressione ENTER para continuar a analise ou digite 'N' para sair"
if ($escolha -eq 'n') {
    Write-Host "`nAnalise cancelada pelo usuario." -ForegroundColor Yellow
    return # Encerra o script
}

Write-Host "`nIniciando analise...`n" -ForegroundColor Green

# --- Definicao do Periodo de Analise ---
$dataInicial = (Get-Date).AddDays(-7)

# --- 1. Verificacao de Logs de Disco ---
Write-Host "--- 1. Verificando Logs de Problemas de Disco ---" -ForegroundColor Yellow
$diskProvidersToCheck = @('Disk', 'Microsoft-Windows-Ntfs', 'Chkdsk', 'Wininit')
$availableDiskProviders = $diskProvidersToCheck | Where-Object { Get-WinEvent -ListProvider $_ -ErrorAction SilentlyContinue }
if ($availableDiskProviders) {
    $diskLogs = Get-WinEvent -FilterHashtable @{LogName='System','Application'; ProviderName=$availableDiskProviders; ID=7,11,51,55,98,153,1001; Level=1,2,3; StartTime=$dataInicial} -ErrorAction SilentlyContinue
    if ($diskLogs) { $diskLogs | Format-Table TimeCreated, ProviderName, Id, LevelDisplayName, Message -AutoSize -Wrap }
    else { Write-Host "Nenhum log relevante de problema de disco encontrado nos ultimos 7 dias." -ForegroundColor Green }
} else { Write-Host "Nenhum provedor de log de disco encontrado." -ForegroundColor Gray }
Write-Host

# --- 2. Verificacao de Logs de Drivers ---
Write-Host "--- 2. Verificando Logs de Problemas de Drivers ---" -ForegroundColor Yellow
$driverProvidersToCheck = @('Microsoft-Windows-DriverFrameworks-UserMode', 'Microsoft-Windows-Kernel-PnP', 'WHEA-Logger')
$availableDriverProviders = $driverProvidersToCheck | Where-Object { Get-WinEvent -ListProvider $_ -ErrorAction SilentlyContinue }
if ($availableDriverProviders) {
    $driverLogs = Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName=$availableDriverProviders; ID=10110,10111,219,1; Level=1,2,3; StartTime=$dataInicial} -ErrorAction SilentlyContinue
    if ($driverLogs) { $driverLogs | Format-Table TimeCreated, ProviderName, Id, LevelDisplayName, Message -AutoSize -Wrap }
    else { Write-Host "Nenhum log relevante de problema de driver encontrado nos ultimos 7 dias." -ForegroundColor Green }
} else { Write-Host "Nenhum provedor de log de driver encontrado." -ForegroundColor Gray }
Write-Host

# --- 3. Verificacao de Falhas Criticas e Instabilidade ---
Write-Host "--- 3. Verificando Falhas Criticas (BSOD) e Travamentos de Apps ---" -ForegroundColor Yellow
$instabilityEvents = Get-WinEvent -FilterHashtable @{LogName='System','Application';ProviderName='Microsoft-Windows-Kernel-Power','Microsoft-Windows-WER-SystemErrorReporting','Application Error','Application Hang';ID=41,1001,1000,1002;Level=1,2;StartTime=$dataInicial} -ErrorAction SilentlyContinue
if ($instabilityEvents) { $instabilityEvents | Format-Table TimeCreated, ProviderName, Id, Message -AutoSize -Wrap }
else { Write-Host "Nenhum log de falha critica ou travamento de aplicacao encontrado nos ultimos 7 dias." -ForegroundColor Green }
Write-Host

# --- 4. Verificacao de Saude do Sistema Operacional ---
Write-Host "--- 4. Verificando Saude do SO (Falhas em Servicos e Updates) ---" -ForegroundColor Yellow
$osHealthEvents = Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Service Control Manager','Microsoft-Windows-WindowsUpdateClient';ID=7000,7009,20;Level=1,2;StartTime=$dataInicial} -ErrorAction SilentlyContinue
if ($osHealthEvents) { $osHealthEvents | Format-Table TimeCreated, ProviderName, Id, Message -AutoSize -Wrap }
else { Write-Host "Nenhum log de falha em servicos ou no Windows Update encontrado nos ultimos 7 dias." -ForegroundColor Green }
Write-Host

# --- 5. Verificacao de Desempenho (Lentidao na Inicializacao) ---
Write-Host "--- 5. Verificando Desempenho da Inicializacao (Boot) ---" -ForegroundColor Yellow
$bootPerfEvents = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Diagnostics-Performance/Operational';ID=100;Level=1,2,3;StartTime=$dataInicial} -ErrorAction SilentlyContinue
if ($bootPerfEvents) { $bootPerfEvents | Format-Table TimeCreated, Id, Message -AutoSize -Wrap }
else { Write-Host "Nenhum log de lentidao na inicializacao encontrado nos ultimos 7 dias." -ForegroundColor Green }
Write-Host

# --- 6. Verificacao de Seguranca (Tentativas de Logon Falhas) ---
Write-Host "--- 6. Verificando Tentativas de Logon com Falha (Requer Admin) ---" -ForegroundColor Yellow
$failedLogons = Get-WinEvent -FilterHashtable @{LogName='Security';ID=4625;StartTime=$dataInicial} -MaxEvents 20 -ErrorAction SilentlyContinue
if ($failedLogons) {
    Write-Host "Resumo das ultimas tentativas de logon com falha (nos ultimos 7 dias):" -ForegroundColor Red
    
    $parsedLogs = foreach ($log in $failedLogons) {
        # Converte o evento para XML para extrair dados de forma confiavel
        $xml = [xml]$log.ToXml()
        
        # Extrai os campos de interesse do XML (independente do idioma do SO)
        $userName = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
        $ipAddress = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'
        
        # Extrai o motivo da falha da mensagem (que ja esta no idioma correto)
        $reason = ($log.Message -split '\r?\n' | Where-Object { $_.Trim().StartsWith('Razao da falha:') } | ForEach-Object { ($_ -split ':', 2)[1].Trim() }) -join '; '
        
        # Cria um objeto personalizado com os dados extraidos
        [PSCustomObject]@{
            'Data/Hora'        = $log.TimeCreated
            'Usuario'          = $userName
            'Motivo da Falha'  = $reason
            'Endereco Origem'  = $ipAddress
        }
    }
    
    $parsedLogs | Format-Table -AutoSize -Wrap
    
} else {
    if ($error[0] -and $error[0].FullyQualifiedErrorId -like "*no logs*") {
         Write-Host "Nenhum log de tentativa de logon com falha encontrado nos ultimos 7 dias." -ForegroundColor Green
    } elseif ($error[0] -and $error[0].FullyQualifiedErrorId -like "*AccessDenied*") {
        Write-Host "Acesso ao Log de Seguranca negado. Execute o script como Administrador para esta verificacao." -ForegroundColor Red
    } else {
        Write-Host "Nenhum log de tentativa de logon com falha encontrado nos ultimos 7 dias." -ForegroundColor Green
    }
}
Write-Host

# --- Finalizacao ---
Write-Host "=================================================="
Write-Host "               Analise Concluida"
Write-Host "=================================================="
pause
Clear-Host
Otimizacoes-Win
}

function CleanTemp {

# --- Configuracao Global ---
$RelatorioPath = "C:\_HTI\Scripts\Limpeza"
$RelatorioFile = Join-Path -Path $RelatorioPath -ChildPath "RelatorioLimpezaDisco.csv"

# --- Funcoes de Coleta de Informacao ---

function Get-FreeDiskSpaceGB {
    param($DriveLetter = 'C')
    $drive = Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue
    if ($drive) { return [math]::Round($drive.Free / 1GB, 2) }
    return 0
}

function Get-OSInfo {
    $os = Get-CimInstance Win32_OperatingSystem
    $osType = switch ($os.ProductType) {
        1 { "Windows Cliente" }
        2 { "Windows Servidor (DC)" }
        3 { "Windows Servidor" }
        default { "Desconhecido" }
    }
    return "$($os.Caption) ($($osType))"
}

# --- Funcoes de Limpeza (Modulos Interativos) ---

function Clear-BrowserCaches {
    Write-Host "-> Verificando caches de navegadores (Edge, Chrome)..." -ForegroundColor Yellow
    $browsers = @{
        "Microsoft Edge" = @{ ProcessName = "msedge"; Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data" }
        "Google Chrome"  = @{ ProcessName = "chrome"; Path = "$env:LOCALAPPDATA\Google\Chrome\User Data" }
    }

    foreach ($name in $browsers.Keys) {
        $browser = $browsers[$name]
        if (-not (Test-Path $browser.Path)) { continue }

        if (Get-Process -Name $browser.ProcessName -ErrorAction SilentlyContinue) {
            Write-Warning "   $name esta em execucao. A limpeza do seu cache nao pode ser feita com o navegador aberto."
            
            $title = "Navegador em Execucao"
            $question = "Deseja forcar o encerramento do $name para continuar com a limpeza?"
            $choices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim", "&Nao")
            $userResponse = $host.UI.PromptForChoice($title, $question, $choices, 1)

            if ($userResponse -eq 0) { # 0 = Sim
                Write-Host "   Forcando o encerramento do $name..." -ForegroundColor Red
                Stop-Process -Name $browser.ProcessName -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
            } else {
                Write-Warning "   A limpeza do cache do $name sera ignorada."
                continue
            }
        }
        Write-Host "   Limpando cache do $name..." -ForegroundColor Cyan
        $cacheFolders = @("Cache", "Code Cache", "GPUCache", "Application Cache")
        $profiles = Get-ChildItem -Path $browser.Path -Directory -Filter "Profile *" | Select-Object -ExpandProperty FullName
        $profiles += Join-Path -Path $browser.Path -ChildPath "Default"
        foreach ($profilePath in $profiles) {
            if (Test-Path $profilePath) {
                foreach ($folder in $cacheFolders) {
                    $cachePath = Join-Path -Path $profilePath -ChildPath $folder
                    if (Test-Path $cachePath) { Remove-Item -Path $cachePath -Recurse -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    }
}

function Clear-CommonSystemJunk {
    Write-Host "-> Limpando arquivos temporarios comuns do sistema..." -ForegroundColor Yellow
    $pathsToClear = @(
        "$env:windir\Temp\*",
        "$env:windir\Downloaded Program Files\*",
        "$env:ProgramData\Microsoft\Windows\WER\*",
        "$env:windir\SoftwareDistribution\Download\*",
        "$env:windir\Minidump\*.dmp",
        "$env:windir\MEMORY.DMP"
    )
    foreach ($path in $pathsToClear) {
        if (Test-Path $path) {
            Write-Host "   Limpando: $path" -ForegroundColor DarkGray
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Host "-> Limpando cache de Otimizacao de Entrega..." -ForegroundColor Yellow
    $deliveryOptPath = "$env:windir\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache\*"
    if (Test-Path $deliveryOptPath) {
        Remove-Item -Path $deliveryOptPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Clear-AllUserTempFiles {
    Write-Host "-> Limpando arquivos temporarios de TODOS os perfis de usuario..." -ForegroundColor Yellow
    
    # Lista de pastas relativas que queremos limpar dentro de cada perfil
    $relativePathsToClear = @(
        "AppData\Local\Temp",
        "AppData\Local\D3DSCache",
        "AppData\Local\Microsoft\Windows\Explorer", # A logica para thumbcache sera especifica
        "AppData\Local\NVIDIA\DXCache",
        "AppData\Local\AMD\DxCache"
    )

    Get-ChildItem -Path "$env:SystemDrive\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $userProfile = $_
        Write-Host "   Verificando perfil: $($userProfile.Name)" -ForegroundColor Cyan
        
        foreach ($relativePath in $relativePathsToClear) {
            $fullPath = Join-Path -Path $userProfile.FullName -ChildPath $relativePath
            
            # Tratamento especial para o cache de miniaturas (thumbnails)
            if ($relativePath -eq "AppData\Local\Microsoft\Windows\Explorer") {
                if (Test-Path $fullPath) {
                    Remove-Item -Path "$fullPath\thumbcache_*.db" -Force -ErrorAction SilentlyContinue
                }
                continue # Pula para o proximo item da lista
            }

            # Logica padrao para as outras pastas
            if (Test-Path $fullPath) {
                # Usamos o curinga '*' aqui, no Remove-Item, e nao no Join-Path
                Remove-Item -Path "$fullPath\*" -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Clear-TempFilesByAge {
    param([int]$Days = 7)
    Write-Host "-> Limpando arquivos e logs com mais de $Days dias..." -ForegroundColor Yellow
    $paths = @(
        "$env:SystemDrive\Temp",
        "$env:windir\LiveKernelReports",
        "$env:ProgramFiles\DebugDiag\Logs",
        "$env:windir\ccmcache",
        "$env:windir\Microsoft.NET\Framework64\v4.0.30319\Temporary ASP.NET Files",
        "$env:windir\Microsoft.NET\Framework\v4.0.30319\Temporary ASP.NET Files",
        "$env:SystemDrive\inetpub\logs\LogFiles",
        "$env:SystemDrive\inetpub\mailroot\Badmail",
        "$env:windir\System32\LogFiles\HTTPERR"
    )
    foreach ($path in $paths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$Days) } | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    }
}

function Clear-WindowsUpdateCleanup {
    Write-Host "-> Verificando arquivos de atualizacao do Windows (Component Store)..." -ForegroundColor Yellow
    
    $title = "Limpeza Profunda do Windows Update"
    $question = "DESEJA REALIZAR UMA LIMPEZA PROFUNDA DE ATUALIZACOES (DISM)? Esta acao pode levar muito tempo."
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim", "&Nao")
    $userResponse = $host.UI.PromptForChoice($title, $question, $choices, 1)

    if ($userResponse -eq 0) { # 0 = Sim
        Write-Host "   Iniciando 'Dism.exe'. O progresso sera exibido abaixo..." -ForegroundColor Cyan
        try {
            Dism.exe /online /Cleanup-Image /StartComponentCleanup
            Write-Host "   Limpeza de componentes do Windows concluida." -ForegroundColor Green
        } catch {
            Write-Error "Falha ao executar a limpeza de componentes com o DISM. $_"
        }
    } else {
        Write-Host "   Limpeza profunda de atualizacoes ignorada." -ForegroundColor Yellow
    }
}

# --- Funcoes para Tarefas Agendadas e Individuais ---

function Invoke-SilentTaskWithAnimation {
    param([Parameter(Mandatory = $true)][string]$Message, [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock)
    $spinner = @('|', '/', '-', '\'); $job = Start-Job -ScriptBlock $ScriptBlock
    Write-Host -NoNewline "$Message "; while ($job.State -eq 'Running') { foreach ($char in $spinner) { Write-Host -NoNewline "`b$char"; Start-Sleep -Milliseconds 100 } }
    $result = Receive-Job -Job $job
    if ($job.JobStateInfo.State -eq 'Failed') {
        Write-Host "`b Falhou!" -ForegroundColor Red; $errorMsg = $job.ChildJobs[0].JobStateInfo.Reason.Message; if ($errorMsg) { Write-Warning $errorMsg }
    } else {
        Write-Host "`b Concluido!" -ForegroundColor Green
    }
    Remove-Job -Job $job
}

# VERSaO CORRIGIDA - SUBSTITUA A SUA FUNcaO EXISTENTE
function Register-ScheduledCleanupTask {
    $scriptDir = "C:\_HTI\Scripts"
    $scheduledScriptPath = Join-Path -Path $scriptDir -ChildPath "Limpeza-Temporarios.ps1"

    Write-Host "-> Criando script de limpeza automatica..." -ForegroundColor Yellow
    if (-not (Test-Path $scriptDir)) { New-Item -Path $scriptDir -ItemType Directory -Force | Out-Null }

    # CONTEUDO DO SCRIPT AGENDADO - CORRIGIDO E ATUALIZADO
    $scriptContent = @'
# SCRIPT DE EXECUCAO AUTOMATICA (v5.1) - NAO INTERATIVO
# Funcoes de limpeza serao executadas silenciosamente e o resultado salvo em log.
$RelatorioPath = "C:\_HTI\Scripts\Limpeza"
$RelatorioFile = Join-Path -Path $RelatorioPath -ChildPath "RelatorioLimpezaDisco.csv"
function Get-FreeDiskSpaceGB { param($DriveLetter='C'); $drive = Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue; if ($drive) { return [math]::Round($drive.Free / 1GB, 2) }; return 0 }
function Get-OSInfo { $os=Get-CimInstance Win32_OperatingSystem; $osType=switch($os.ProductType){1{"Cliente"}2{"Servidor(DC)"}3{"Servidor"}}; return "$($os.Caption) ($($osType))" }
function Clear-BrowserCaches-Silent {
    $browsers=@{"Microsoft Edge"=@{ProcessName="msedge";Path="$env:LOCALAPPDATA\Microsoft\Edge\User Data"};"Google Chrome"=@{ProcessName="chrome";Path="$env:LOCALAPPDATA\Google\Chrome\User Data"}}
    foreach($name in $browsers.Keys){
        $browser=$browsers[$name]
        if(-not(Test-Path $browser.Path) -or (Get-Process -Name $browser.ProcessName -ErrorAction SilentlyContinue)){continue}
        $cacheFolders=@("Cache","Code Cache","GPUCache","Application Cache")
        $profiles=Get-ChildItem -Path $browser.Path -Directory -Filter "Profile *" | Select-Object -ExpandProperty FullName
        $profiles+=Join-Path -Path $browser.Path -ChildPath "Default"
        foreach($profilePath in $profiles){ if(Test-Path $profilePath){ foreach($folder in $cacheFolders){ $cachePath=Join-Path -Path $profilePath -ChildPath $folder; if(Test-Path $cachePath){ Remove-Item -Path $cachePath -Recurse -Force -EA 0 }}}}
    }
}
function Clear-CommonSystemJunk-Silent {
    @("$env:windir\Temp\*","$env:windir\Downloaded Program Files\*","$env:ProgramData\Microsoft\Windows\WER\*","$env:windir\SoftwareDistribution\Download\*","$env:windir\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache\*","$env:windir\Minidump\*.dmp","$env:windir\MEMORY.DMP") | ForEach-Object { if(Test-Path $_.psobject.BaseObject){Remove-Item -Path $_ -Recurse -Force -EA 0} }
}
function Clear-AllUserTempFiles-Silent {
    $relativePaths = @("AppData\Local\Temp","AppData\Local\D3DSCache","AppData\Local\Microsoft\Windows\Explorer","AppData\Local\NVIDIA\DXCache","AppData\Local\AMD\DxCache")
    Get-ChildItem -Path "$env:SystemDrive\Users" -Directory -EA 0 | ForEach-Object {
        $userProfile = $_
        foreach($relativePath in $relativePaths){
            $fullPath = Join-Path -Path $userProfile.FullName -ChildPath $relativePath
            if($relativePath -eq "AppData\Local\Microsoft\Windows\Explorer"){ if(Test-Path $fullPath){ Remove-Item -Path "$fullPath\thumbcache_*.db" -Force -EA 0 } }
            else { if(Test-Path $fullPath){ Remove-Item -Path "$fullPath\*" -Recurse -Force -EA 0 } }
        }
    }
}
function Clear-TempFilesByAge-Silent {
    param([int]$Days=7)
    $pathsToClear = @("$env:SystemDrive\Temp","$env:windir\LiveKernelReports","$env:ProgramFiles\DebugDiag\Logs","$env:windir\ccmcache","$env:windir\Microsoft.NET\Framework64\v4.0.30319\Temporary ASP.NET Files","$env:windir\Microsoft.NET\Framework\v4.0.30319\Temporary ASP.NET Files","$env:SystemDrive\inetpub\logs\LogFiles","$env:SystemDrive\inetpub\mailroot\Badmail","$env:windir\System32\LogFiles\HTTPERR")
    foreach($path in $pathsToClear){ if(Test-Path $path){ Get-ChildItem -Path $path -Recurse -File -EA 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$Days) } | Remove-Item -Force -EA 0 }}
}
function Clear-WindowsUpdateCleanup-Silent { Dism.exe /online /Cleanup-Image /StartComponentCleanup /quiet }
function Clear-RecycleBin-Silent { try { Clear-RecycleBin -Force -ErrorAction Stop } catch { Remove-Item -Path "C:\`$Recycle.Bin\*" -Recurse -Force -EA 0 } }
function Clear-NvidiaCache-Silent { if(Test-Path "$env:ProgramData\NVIDIA Corporation\NV_Cache"){ Remove-Item -Path "$env:ProgramData\NVIDIA Corporation\NV_Cache\*" -Recurse -Force -EA 0 } }
function Run-SilentCleanup {
    $hostname=$env:COMPUTERNAME;$so=Get-OSInfo;$espacoAntes=Get-FreeDiskSpaceGB;$data=Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    Clear-RecycleBin-Silent;Clear-BrowserCaches-Silent;Clear-CommonSystemJunk-Silent;Clear-AllUserTempFiles-Silent;Clear-TempFilesByAge-Silent;Clear-WindowsUpdateCleanup-Silent;Clear-NvidiaCache-Silent
    $espacoDepois=Get-FreeDiskSpaceGB
    $resultado=[PSCustomObject]@{Tarefa="Limpeza Agendada";Hostname=$hostname;SistemaOperacional=$so;EspacoLivreAntesGB=$espacoAntes;EspacoLivreDepoisGB=$espacoDepois;EspacoRecuperadoGB=[math]::Round($espacoDepois-$espacoAntes,2);DataExecucao=$data}
    if(-not(Test-Path(Split-Path $RelatorioFile -Parent))){ New-Item -ItemType Directory -Path (Split-Path $RelatorioFile -Parent) -Force | Out-Null }
    $resultado|Export-Csv -Path $RelatorioFile -NoTypeInformation -Encoding UTF8 -Append
}
Run-SilentCleanup
'@
    $scriptContent | Set-Content -Path $scheduledScriptPath -Encoding UTF8
    Write-Host "   Script 'Limpeza-Temporarios.ps1' (v5.1) criado/atualizado em '$scriptDir'." -ForegroundColor Green

    Write-Host "-> Agendando a tarefa de limpeza..." -ForegroundColor Yellow
    $taskName = "Limpeza Periodica HTI"
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scheduledScriptPath`""
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 12:30pm
    $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)
    try {
        Register-ScheduledTask -TaskName "Limpeza Periodica HTI" -TaskPath "\HTI\" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Executa a limpeza periodica de arquivos temporarios." -Force
        Write-Host "   Tarefa '$taskName' agendada com sucesso para toda Segunda-feira as 12:30." -ForegroundColor Green
    } catch { Write-Error "Ocorreu um erro ao agendar a tarefa: $_" }
}

# --- Funcao de Execucao e Relatorio (para modo interativo) ---

function Start-CleanupAndReport {
    param([Parameter(Mandatory=$true)][string]$NomeDaTarefa, [Parameter(Mandatory=$true)][scriptblock]$CleanupScriptBlock)
    $hostname = $env:COMPUTERNAME; $so = Get-OSInfo; $espacoAntes = Get-FreeDiskSpaceGB; $data = Get-Date -f "dd-MM-yyyy HH:mm:ss"
    Write-Host "`n================ INICIO: $NomeDaTarefa EM $hostname ================" -ForegroundColor White -BackgroundColor DarkBlue
    Invoke-Command -ScriptBlock $CleanupScriptBlock
    $espacoDepois = Get-FreeDiskSpaceGB; $recuperado = [math]::Round($espacoDepois - $espacoAntes, 2)
    Write-Host "================ FIM DA TAREFA: $NomeDaTarefa ================" -ForegroundColor White -BackgroundColor DarkGreen
    Write-Host "Espaco em disco recuperado: $recuperado GB" -ForegroundColor Green
    try {
        if (!(Test-Path -Path $RelatorioPath)) { New-Item -ItemType Directory -Path $RelatorioPath -Force | Out-Null }
        $resultado = [PSCustomObject]@{ Tarefa = $NomeDaTarefa; Hostname = $hostname; SistemaOperacional = $so; EspacoLivreAntesGB = $espacoAntes; EspacoLivreDepoisGB = $espacoDepois; EspacoRecuperadoGB = $recuperado; DataExecucao = $data }
        $resultado | Export-Csv -Path $RelatorioFile -NoTypeInformation -Encoding UTF8 -Append
        Write-Host "Relatorio salvo em: $RelatorioFile" -ForegroundColor Cyan
    } catch { Write-Error "Falha ao criar diretorio ou salvar o relatorio CSV: $_" }
}

# --- MENU PRINCIPAL ---

:mainMenu while ($true) {
    Clear-Host
    Write-Host "========= MENU DE LIMPEZA DE DISCO (v5.0) =========" -ForegroundColor Green
    Write-Host "Host: $($env:COMPUTERNAME) | OS: $(Get-OSInfo)" -ForegroundColor DarkGray
    Write-Host "---------------------------------------------------"
    Write-Host "OPCOES DE LIMPEZA COMPLETA (com relatorio):" -ForegroundColor Cyan
    Write-Host "  1) Limpeza Completa de DESKTOP (Windows 10/11)"
    Write-Host "  2) Limpeza Completa de SERVIDOR (Windows Server)"
    Write-Host ""
    Write-Host "OPCOES DE LIMPEZA INDIVIDUAL:" -ForegroundColor Cyan
    Write-Host "  3) Limpar Cache DNS"
    Write-Host "  4) Limpar Cache da Microsoft Store"
    Write-Host "  5) Limpar Lixeira"
    Write-Host "  6) Limpar Cache de Drivers NVIDIA"
    Write-Host "  7) Limpeza de Atualizacoes do Windows (DISM)"
    Write-Host ""
    Write-Host "AUTOMACAO:" -ForegroundColor Cyan
    Write-Host "  8) Agendar/Reagendar limpeza periodica (Segura)"
    Write-Host "---------------------------------------------------"
    Write-Host "  0) Voltar menu anterior" -ForegroundColor Red
    Write-Host "==================================================="
    
    $clearRecycleBinScriptBlock = { try { Clear-RecycleBin -Force -ErrorAction Stop } catch { Write-Warning "Nao foi possivel... Usando metodo alternativo."; Remove-Item -Path "C:\`$Recycle.Bin\*" -Recurse -Force -ErrorAction SilentlyContinue } }
    $productType = (Get-CimInstance Win32_OperatingSystem).ProductType
    $choice = Read-Host "Escolha uma opcao"

    switch ($choice) {
        '1' {
            Write-Host "`nA Limpeza Completa de Desktop realizara as seguintes acoes:" -ForegroundColor Cyan
            Write-Host @"
  - Limpeza da Lixeira
  - Limpeza de caches dos navegadores (Edge e Chrome)
  - Limpeza de arquivos temporarios do sistema (Temp, Logs, Dumps)
  - Limpeza de arquivos temporarios de TODOS os perfis de usuario
  - Limpeza de logs e arquivos antigos com mais de 7 dias
  - Limpeza de atualizacoes do Windows (DISM, sera perguntado)
  - Limpeza do cache de drivers NVIDIA
"@ -ForegroundColor Yellow
            
            $title = "Confirmar Limpeza Completa"
            $question = "Deseja continuar com estas limpezas?"
            $choices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim", "&Nao")
            $userResponse = $host.UI.PromptForChoice($title, $question, $choices, 0)

            if ($userResponse -eq 1) { # 1 = Nao
                Write-Host "Operacao cancelada pelo usuario." -ForegroundColor Red; Start-Sleep -Seconds 2; break
            }
            
            Start-CleanupAndReport -NomeDaTarefa "Limpeza Desktop Completa" -CleanupScriptBlock {
                Invoke-SilentTaskWithAnimation -Message "Limpando Lixeira" -ScriptBlock $clearRecycleBinScriptBlock
                Clear-BrowserCaches
                Clear-CommonSystemJunk
                Clear-AllUserTempFiles
                Clear-TempFilesByAge
                Clear-WindowsUpdateCleanup
                Invoke-SilentTaskWithAnimation -Message "Limpando cache NVIDIA" -ScriptBlock { if(Test-Path "$env:ProgramData\NVIDIA Corporation\NV_Cache"){Remove-Item "$env:ProgramData\NVIDIA Corporation\NV_Cache\*" -Recurse -Force -ErrorAction SilentlyContinue}}
            }
        }
        '2' { # Limpeza de Servidor
            Start-CleanupAndReport -NomeDaTarefa "Limpeza Servidor Completa" -CleanupScriptBlock {
                Invoke-SilentTaskWithAnimation -Message "Limpando Lixeira" -ScriptBlock $clearRecycleBinScriptBlock
                Clear-CommonSystemJunk
                Clear-AllUserTempFiles
                Clear-TempFilesByAge
                Clear-WindowsUpdateCleanup
            }
        }
        '3' { Invoke-SilentTaskWithAnimation -Message "Limpando Cache DNS" -ScriptBlock { Clear-DnsClientCache } }
        '4' { Invoke-SilentTaskWithAnimation -Message "Limpando Cache da Microsoft Store" -ScriptBlock { if ($productType -ne 1) { throw "Opcao apenas para Windows Cliente." }; Start-Process "wsreset.exe" -ArgumentList "-s" -Wait -NoNewWindow } }
        '5' { Invoke-SilentTaskWithAnimation -Message "Limpando Lixeira" -ScriptBlock $clearRecycleBinScriptBlock }
        '6' { Invoke-SilentTaskWithAnimation -Message "Limpando Cache NVIDIA" -ScriptBlock { if (-not (Test-Path "$env:ProgramData\NVIDIA Corporation\NV_Cache")) { throw "Diretorio nao encontrado." }; Remove-Item -Path "$env:ProgramData\NVIDIA Corporation\NV_Cache\*" -Recurse -Force -ErrorAction SilentlyContinue } }
        '7' { Clear-WindowsUpdateCleanup }
        '8' { Register-ScheduledCleanupTask }
        '0' { Otimizacoes-Win } # Retorna ao menu anterior
        default { Write-Host "Opcao invalida. Tente novamente." -ForegroundColor Red }
    }
    if ($choice -ne '0') { Write-Host "`nPressione ENTER para voltar ao menu..."; Read-Host | Out-Null }
}
}


function SystemIntegrityCheck {
    [CmdletBinding()]
    param()
	Clear-Host
    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host "FUNCAO DE VERIFICAcaO DE INTEGRIDADE DO SISTEMA" -ForegroundColor Yellow
    Write-Host "==========================================================="
    Write-Host "Esta funcao executara um reparo completo do sistema em duas etapas:"
    Write-Host " 1. Reparo da Imagem do Windows (DISM): Corrige a 'caixa de pecas' do sistema."
    Write-Host " 2. Verificador de Arquivos (SFC): Usa as pecas corrigidas para reparar o sistema."
    Write-Host "`nAVISO: Este processo pode levar mais de uma hora para ser concluido." -ForegroundColor Red
    
    # Criando o prompt de confirmacao (Sim/Nao)
    $title = "Confirmar Execucao"
    $question = "Deseja iniciar o processo completo de verificacao e reparo?"
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim", "&Nao")
    $defaultChoice = 0 # "Sim" e a opcao padrao (indice 0)

    # Exibe o prompt para o usuario e captura a resposta
    $userResponse = $host.UI.PromptForChoice($title, $question, $choices, $defaultChoice)

    # Se o usuario escolheu "Nao" (indice 1), a funcao encerra.
    if ($userResponse -eq 1) {
        Write-Host "`nOperacao cancelada pelo usuario. Retornando ao menu." -ForegroundColor Yellow
        Clear-Host
		Otimizacoes-Win
    }

    # Se o usuario confirmou, o processo continua dentro de um bloco Try/Catch
    try {
        # Etapa 1: Reparar a Imagem do Windows (Component Store)
        Write-Host "`n[ETAPA 1 de 2] Executando o Reparo da Imagem do Windows..." -ForegroundColor Cyan
        Repair-WindowsImage -Online -RestoreHealth -ErrorAction Stop
        Write-Host "[SUCESSO] Etapa 1 concluida." -ForegroundColor Green
        Start-Sleep -Seconds 3

        # Etapa 2: Verificar e Reparar Arquivos de Sistema (SFC)
        Write-Host "`n[ETAPA 2 de 2] Executando o Verificador de Arquivos do Sistema (SFC)..." -ForegroundColor Cyan
        sfc.exe /scannow

        Write-Host "`n===========================================================" -ForegroundColor Yellow
        Write-Host "VERIFICACAO DE INTEGRIDADE CONCLUIDA!" -ForegroundColor Green
        Write-Host "Analise a saida dos comandos para detalhes sobre as correcoes." -ForegroundColor Green
        Write-Host "==========================================================="
    }
    catch {
        Write-Error "Ocorreu um erro critico durante o processo de verificacao."
        Write-Error $_.Exception.Message
    }
}


# Requer que o script principal seja executado como Administrador
#requires -RunAsAdministrator

# Requer que o script principal seja executado como Administrador
#requires -RunAsAdministrator

function NetworkStackReset {
    [CmdletBinding()]
    param()
	Clear-Host
    Write-Host "================== Reset de Rede ==================" -ForegroundColor Yellow
    
    # ETAPA 1: VERIFICACAO DE CONFIGURACAO MANUAL/ESTATICA (LOGICA CORRIGIDA E DEFINITIVA)
    # -------------------------------------------------------------------------------------
    Write-Host "`n[PASSO 1] Verificando configuracoes de rede existentes..." -ForegroundColor Cyan
    
    $manualConfigs = @()
    $activeInterfaces = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }

    foreach ($interface in $activeInterfaces) {
        $ipIsStatic = $false
        $dnsIsStatic = $false

        # Verifica a ORIGEM do endereco IP (Manual ou DHCP)
        $ipAddresses = Get-NetIPAddress -InterfaceIndex $interface.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
        if ($ipAddresses | Where-Object { $_.PrefixOrigin -eq 'Manual' }) {
            $ipIsStatic = $true
        }

        # Verifica se existem servidores DNS configurados MANUALMENTE
        $dnsServers = Get-DnsClientServerAddress -InterfaceIndex $interface.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
        if ($dnsServers) {
            $dnsIsStatic = $true
        }

        # Se o IP ou o DNS for estatico, adiciona a interface à lista de alerta
        if ($ipIsStatic -or $dnsIsStatic) {
            $configDetails = Get-NetIPConfiguration -InterfaceIndex $interface.ifIndex
            $manualConfigs += $configDetails
        }
    }

    if ($manualConfigs) {
        Write-Host "!!! ALERTA DE CONFIGURACAO MANUAL/ESTATICA DETECTADA !!!" -ForegroundColor Red
        Write-Host "A(s) seguinte(s) interface(s) possui(em) configuracoes que nao sao 100% automaticas:"
        
        foreach ($config in $manualConfigs) {
            Write-Host "  - Interface: $($config.InterfaceAlias)" -ForegroundColor Yellow
            
            $ipInfo = Get-NetIPAddress -InterfaceIndex $config.InterfaceIndex -AddressFamily IPv4 | Select-Object -First 1
            if ($ipInfo.PrefixOrigin -eq 'Manual') {
                Write-Host "    IP Estatico: $($ipInfo.IPAddress)" -ForegroundColor Red
            } else {
                Write-Host "    IP (via DHCP): $($ipInfo.IPAddress)" -ForegroundColor Green
            }
            
            $dnsInfo = Get-DnsClientServerAddress -InterfaceIndex $config.InterfaceIndex -AddressFamily IPv4
            if ($dnsInfo) {
                Write-Host "    Servidores DNS Manuais/Estaticos: $($dnsInfo.ServerAddresses -join ', ')" -ForegroundColor Red
            }
        }
        
        Write-Host "`nO comando 'netsh int ip reset' ira APAGAR essas configuracoes, revertendo a interface para 100% DHCP." -ForegroundColor Red
        
        $title = "Confirmar Acao Destrutiva"
        $question = "TEM CERTEZA que deseja continuar e apagar as configuracoes manuais listadas acima?"
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim, apagar configuracoes", "&Nao, cancelar AGORA")
        $response = $host.UI.PromptForChoice($title, $question, $choices, 1) # Padrao e "Nao"
        
        if ($response -ne 0) {
            Write-Host "`nOperacao cancelada com seguranca pelo usuario. Nenhuma alteracao foi feita." -ForegroundColor Green
            pause
			clear-host
			return
        }
        Write-Host "`nConfirmacao explicita recebida. Prosseguindo com o reset..." -ForegroundColor Red
    } else {
        Write-Host "   Nenhuma configuracao manual ou estatica encontrada. Prossiga com seguranca." -ForegroundColor Green
    }

    # ETAPA 2: AVISO PADRAO E CONFIRMACAO FINAL
    # (O restante da funcao continua exatamente igual)
    # ----------------------------------------------------
    Write-Host "`nAVISO: A redefinicao de rede desconectara temporariamente a internet." -ForegroundColor Yellow
    Write-Host "Uma REINICIALIZACAO sera necessaria no final." -ForegroundColor Yellow

    $finalTitle = "Confirmar Reset da Rede"
    $finalQuestion = "Prosseguir com a redefinicao de rede?"
    $finalChoices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim, continuar", "&Nao, cancelar")
    $finalResponse = $host.UI.PromptForChoice($finalTitle, $finalQuestion, $finalChoices, 1)

    if ($finalResponse -ne 0) {
        Write-Host "`nOperacao cancelada pelo usuario." -ForegroundColor Yellow
        return
    }

    # ETAPA 3: EXECUCAO DOS COMANDOS
    # ---------------------------------
    try {
        Write-Host "`nIniciando o processo de redefinicao de rede..." -ForegroundColor Green
        Write-Host "[PASSO 1/4] Limpando o cache de DNS do cliente..." -ForegroundColor Cyan
        Clear-DnsClientCache
        
        Write-Host "[PASSO 2/4] Liberando e Renovando o endereco IP..." -ForegroundColor Cyan
        ipconfig /release | Out-Null
        ipconfig /renew | Out-Null
        Write-Host "   Endereco IP renovado." -ForegroundColor Green

        Write-Host "[PASSO 3/4] Redefinindo o Catalogo Winsock..." -ForegroundColor Cyan
        netsh.exe winsock reset | Out-Null
        Write-Host "   Winsock redefinido com sucesso." -ForegroundColor Green

        Write-Host "[PASSO 4/4] Redefinindo a pilha TCP/IP..." -ForegroundColor Cyan
        netsh.exe int ip reset | Out-Null
        Write-Host "   TCP/IP redefinido com sucesso." -ForegroundColor Green

        Write-Host "`n========================================================" -ForegroundColor Green
        Write-Host "PROCESSO DE RESET DE REDE CONCLUIDO COM SUCESSO!"
        Write-Host "E fundamental que voce reinicie o computador agora."
        Write-Host "========================================================"
        
        $rebootTitle = "Reinicializacao Necessaria"
        $rebootQuestion = "Deseja reiniciar o computador agora para aplicar todas as alteracoes?"
        $rebootChoices = [System.Management.Automation.Host.ChoiceDescription[]]@("&Sim, reiniciar agora", "&Nao, eu reiniciarei manualmente")
        $rebootResponse = $host.UI.PromptForChoice($rebootTitle, $rebootQuestion, $rebootChoices, 0)

        if ($rebootResponse -eq 0) {
            Write-Host "Reiniciando o computador em 5 segundos..." -ForegroundColor Yellow
            Restart-Computer -Force
        } else {
            Write-Host "Lembre-se de reiniciar o computador o mais breve possivel." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Ocorreu um erro inesperado durante a redefinicao de rede."
        Write-Error $_.Exception.Message
        Write-Error "Por favor, certifique-se de que o script foi executado como Administrador."
    }
}

# Definir a funcao de ferramentas e otimizacoes do Windows
function Otimizacoes-Win {
    Clear-Host
	Write-Host "`n--------------------------------------------------------------" -ForegroundColor Green
    Write-Host "           Ferramentas e Otimizacoes do Windows" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Green

    # Mostrar opcoes do menu de gerenciamento de usuarios locais
    Write-Host "1) Gerenc. Energia - Sempre ativo" -ForegroundColor Green
    Write-Host "2) Atualizacao do Windows - Em Desenv" -ForegroundColor Yellow
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
        "0" { Clear-Host; Main }  # Voltar ao menu principal
        default { Write-Host "Opcao invalida. Por favor, escolha novamente." }
    }
}

# Funcao principal
function Main {
    Clear-Host
    Remove-HTIFiles  # Verifica e remove os arquivos no inicio
    
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
        Write-Host "Versao: 0.7.5 Beta`n" -ForegroundColor DarkGray
        Write-Host "1) Exibir Informacoes do Windows" -ForegroundColor Green
        Write-Host "2) Renomear o computador" -ForegroundColor Green
        Write-Host "3) Instalacao do agente TiFlux" -ForegroundColor Green
        Write-Host "4) Instalacao do Endpoint Bitdefender - PENDENTE" -ForegroundColor DarkGray
        Write-Host "5) Instalacao do pacote basico de programas" -ForegroundColor Green
        Write-Host "6) Instalacao do pacote de programas juridicos - PENDENTE" -ForegroundColor DarkGray
        Write-Host "7) Instalacao do pacote de programas contabeis - PENDENTE" -ForegroundColor DarkGray
        Write-Host "8) Instalacao do Office" -ForegroundColor Green
        Write-Host "9) Ativacao do Windows e Office" -ForegroundColor Green
        Write-Host "10) Remove Apps desnecessarios" -ForegroundColor Green
        Write-Host "11) Baixar aplicativos utilitarios {DiskUse, BSOD, Sysinternals...}" -ForegroundColor Green
        Write-Host "12) Gerenciador de Usuarios Locais" -ForegroundColor Green
        Write-Host "13) Ferramentas e Otimizacoes - Em Desenv." -ForegroundColor Yellow
        Write-Host "14) Monitorar uso do HD e envio de alertas" -ForegroundColor Green
        Write-Host ""
        Write-Host "0) Sair - Limpa historico"
        
        $mainOption = Read-Host "`nEscolha uma opcao"

        switch ($mainOption) {
            "1" { Write-Host "Aguarde, coletando informacoes do Windows..." -ForegroundColor Yellow; Info-Windows }
            "2" { Rename-Computer }
            "3" { Install-Agente }
            "4" { Install-Bitdefender }
            "5" { Install-BasicPackage }
            "6" { Install-LegalPackage }
            "7" { Install-AccountingPackage }
            "8" { Install-Office }
            "9" { Activate-WindowsOffice }
            "10" { Remove-Apps }
            "11" { Download-Utilities }
            "12" { Manage-LocalUsers }
            "13" { Otimizacoes-Win }
            "14" { Disk-ManagementMenu }
            "99" { Show-Changelog }
            "0" { 
                Write-Host "limpando historico powershell..." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Remove-HTIFiles  # Remove os arquivos antes de sair
				[System.IO.File]::WriteAllText((Get-PSReadLineOption).HistorySavePath, "")  # Limpa o historico de comandos
                Clear-History
				Clear-Host
                return
            }
            default { Write-Host "Opcao invalida. Tente novamente." }
        }
    } while ($mainOption -ne "0")
}

# Funcao para remover os arquivos especificados
function Remove-HTIFiles {
    $files = @(
        "C:\\_HTI\\Scripts\\Script-Inicial.ps1",
        "C:\\_HTI\\Scripts\\clientes.csv"
    )
    
    foreach ($file in $files) {
        if (Test-Path $file) {
            Remove-Item -Path $file -Force -ErrorAction SilentlyContinue
#            Write-Host "Arquivo removido: $file" -ForegroundColor Yellow#
        }
    }
}

# Funcao para exibir o changelog do script
function Show-Changelog {
    Clear-Host
    Write-Host "`n--------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "                  HISTORICO DE ALTERACOES" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------`n" -ForegroundColor Cyan
    Write-Host "[v0.7.5] - Adicionado a funcao de reset da rede do Windows"
    Write-Host "[v0.7.4] - Adicionado a funcao para analise de integridade profunda (imagem) do Windows"
	Write-Host "[v0.7.3] - Adicionado a funcao para Limpeza da Fila de Impressao"
    Write-Host "[v0.7.2] - Adicionado a analise do WinSat nas informacoes iniciais do computador"
    Write-Host "[v0.7.1] - Analise dos principais IDs de alerta do EventViewer e adicional o app CrystalInfo para analise de SSD"
    Write-Host "[v0.7.0] - Adicionado o recurso de limpeza de disco"
    Write-Host "[v0.6.9] - Aprimorado selecao do cliente na opcao 3 do TiFlux"
    Write-Host "[v0.6.8] - Adicionado a funcao de Instalar e Remover o Anydesk"
    Write-Host "[v0.6.7] - Adicionado a funcao de Limpeza inteligente de cache do Teams"    
    Write-Host "[v0.6.6] - Correcoes e melhorias na opcao 1 Informacoes do Windows"    
    Write-Host "[v0.6.5] - Adicionado o comando para limpar o historico de comandos do powershell ao sair do script"
    Write-Host "[v0.6.4] - Adicionado uma funcao para excluir os arquivos Script-Inicial e clientes.csv da pasta HTI\Scripts"
    Write-Host "[v0.6.3] - Criado o ChangeLog - Historico de Alteracoes"
    Write-Host "[v0.6.2] - Criado a funcao: 1) Gerenc. Energia - Sempre ativo."
    Write-Host "[v0.6.1] - Melhorado a funcao do Winget para ser utilizado dentro do DW-Service Shell Background."
    Write-Host "[v0.6.0] - Implementacao inicial das ferramentas de otimizacao."
    
    Write-Host "`nPressione qualquer tecla para voltar ao menu principal..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Main
}

Main