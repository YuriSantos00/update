# Atualizador.ps1 - Windows Update + log local + envio para Google Sheets

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Caminhos
$logPath = "C:\Appmax"
$logFile = Join-Path $logPath "update-log.txt"

# Criar diretório, se necessário
if (-Not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
}

# Início do log
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
$dataHora = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Lista atualizações disponíveis
$updates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll

if ($updates.Count -eq 0) {
    Add-Content -Path $logFile -Value "Nenhuma atualização pendente encontrada." -Encoding utf8
} else {
    Add-Content -Path $logFile -Value "$($updates.Count) atualização(ões) encontrada(s). Iniciando instalação..." -Encoding utf8

    # Instala atualizações e registra no log
    Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -IgnoreReboot -Verbose |
        Tee-Object -FilePath $logFile -Append -Encoding utf8
}

# Coleta histórico das atualizações instaladas
$updatesHist = Get-WUHistory | Where-Object {$_.Result -eq "Succeeded"} | Sort-Object Date -Descending | Select-Object -First 10
$titulos = $updatesHist | ForEach-Object { $_.Title }
$todosUpdates = $titulos -join "; "

# JSON para envio
$dados = @{
    data         = $dataHora
    hostname     = $hostname
    so           = $so
    atualizacoes = $todosUpdates
} | ConvertTo-Json -Depth 3

# Envio para Google Sheets
try {
    Invoke-RestMethod -Uri "https://script.google.com/macros/s/AKfycbwHp-e0DTsSk4u4GK3_m4Lryt7GMXIjxb68qFUsxuqjO5OkgBGQv48UGqitN5AT4WmM/exec" `
        -Method Post `
        -Body $dados `
        -ContentType "application/json"
    Add-Content -Path $logFile -Value "Envio para Google Sheets concluído com sucesso." -Encoding utf8
} catch {
    Add-Content -Path $logFile -Value "Erro ao enviar para Google Sheets: $($_.Exception.Message)" -Encoding utf8
}

# Fim do log
Add-Content -Path $logFile -Value "===== Fim da execução: $(Get-Date) =====`n" -Encoding utf8
