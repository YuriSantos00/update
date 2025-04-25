# Atualizador.ps1 - Windows Update + log local + envio para Google Sheets + coleta apenas das atualizações da execução atual

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Caminhos
$logPath = "C:\Appmax"
$logFile = Join-Path $logPath "update-log.txt"

# Criar diretório, se necessário
if (-Not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
}

# Início da execução
Add-Content -Path $logFile -Value "`n===== Início da execução: $(Get-Date) =====" -Encoding utf8

# Verifica e instala PSWindowsUpdate
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
    Install-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
    Add-Content -Path $logFile -Value "Módulo PSWindowsUpdate instalado com sucesso." -Encoding utf8
} else {
    Add-Content -Path $logFile -Value "Módulo PSWindowsUpdate já instalado." -Encoding utf8
}

# Importa e ativa Microsoft Update
Import-Module PSWindowsUpdate -Force
Add-WUServiceManager -MicrosoftUpdate -Confirm:$false

# Informações da máquina
$hostname = $env:COMPUTERNAME
$so = (Get-CimInstance Win32_OperatingSystem).Caption
$dataHoraInicio = Get-Date
$dataHoraTexto = $dataHoraInicio.ToString("yyyy-MM-dd HH:mm:ss")

# Lista atualizações disponíveis
$updates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll

if ($updates.Count -eq 0) {
    Add-Content -Path $logFile -Value "Nenhuma atualização pendente encontrada." -Encoding utf8
} else {
    Add-Content -Path $logFile -Value "$($updates.Count) atualização(ões) encontrada(s). Iniciando instalação..." -Encoding utf8

    try {
        Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -IgnoreReboot -Verbose |
            Tee-Object -FilePath $logFile -Append

        if ($?) {
            Add-Content -Path $logFile -Value "Comando Install-WindowsUpdate executado com sucesso." -Encoding utf8
        } else {
            Add-Content -Path $logFile -Value "Comando Install-WindowsUpdate não foi executado corretamente." -Encoding utf8
        }
    } catch {
        Add-Content -Path $logFile -Value "Erro ao instalar atualizações: $($_.Exception.Message)" -Encoding utf8
    }
}

# Aguardar alguns segundos para garantir que Get-WUHistory atualizou
Start-Sleep -Seconds 10

# Coleta histórico das atualizações instaladas após a execução
$dataHoraLimite = $dataHoraInicio.AddMinutes(-5) # Coletar updates nos últimos 5 minutos para segurança
$updatesHist = Get-WUHistory | Where-Object {
    $_.Result -eq "Succeeded" -and $_.Date -ge $dataHoraLimite
} | Sort-Object Date -Descending

$titulos = $updatesHist | ForEach-Object { $_.Title }
$todosUpdates = $titulos -join "; "

# Monta dados JSON com codificação segura
$payload = @{ 
    data         = $dataHoraTexto
    hostname     = $hostname
    so           = $so
    atualizacoes = $todosUpdates
}
$json = $payload | ConvertTo-Json -Depth 3
$body = [System.Text.Encoding]::UTF8.GetBytes($json)

# Envio para Google Sheets
try {
    $response = Invoke-RestMethod -Uri "https://script.google.com/macros/s/AKfycbwHp-e0DTsSk4u4GK3_m4Lryt7GMXIjxb68qFUsxuqjO5OkgBGQv48UGqitN5AT4WmM/exec" `
        -Method Post `
        -Body $body `
        -ContentType "application/json"
    Add-Content -Path $logFile -Value "Envio para Google Sheets concluído com sucesso. Resposta: $response" -Encoding utf8
} catch {
    Add-Content -Path $logFile -Value "Erro ao enviar para Google Sheets: $($_.Exception.Message)" -Encoding utf8
}

# Fim do log
Add-Content -Path $logFile -Value "===== Fim da execução: $(Get-Date) =====`n" -Encoding utf8
