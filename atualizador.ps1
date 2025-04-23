
# Atualizador.ps1 - Atualizações do Windows + log local + envio remoto

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Caminhos
$logPath = "C:\Appmax"
$logFile = Join-Path $logPath "update-log.txt"
$csvFile = Join-Path $logPath "report.csv"

# Criar diretório, se necessário
if (-Not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
}

# Início do log com Transcript
Start-Transcript -Path $logFile -Append
Write-Output "===== Início da execução: $(Get-Date) ====="

# Verifica e instala PSWindowsUpdate, se necessário
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Install-Module -Name PSWindowsUpdate -Force
    Write-Output "Módulo PSWindowsUpdate instalado com sucesso."
} else {
    Write-Output "Módulo PSWindowsUpdate já instalado."
}

# Importa o módulo e ativa Microsoft Update
Import-Module PSWindowsUpdate -Force
Add-WUServiceManager -MicrosoftUpdate -Confirm:$false

# Informações da máquina
$hostname = $env:COMPUTERNAME
$user = (Get-WmiObject Win32_ComputerSystem).UserName
$date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Criar CSV com cabeçalho, se necessário
if (-not (Test-Path $csvFile)) {
    Add-Content -Path $csvFile -Value "Data,Hostname,Usuário,Título da Atualização" -Encoding utf8
}

# Executa updates e captura resultado da execução atual
$updatesThisRun = Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -ForceDownload -ForceInstall -IgnoreReboot -Verbose

# Enviar cada update registrado
foreach ($update in $updatesThisRun) {
    $titulo = $update.Title
    $linha = $date + "," + $hostname + "," + $user + "," + '"' + $titulo + '"'
    Add-Content -Path $csvFile -Value $linha -Encoding utf8

    $json = @{
        data     = $date
        hostname = $hostname
        usuario  = $user
        titulo   = $titulo
    } | ConvertTo-Json -Depth 3

    try {
        Invoke-RestMethod -Uri "https://script.google.com/macros/s/AKfycby7UBZ4jFH10wmHC7KxYB6ZTFbeUfZcdFAoz5X3L9ln0CfomJ1Xtfqhpu14P6vlLVQ/exec" `
            -Method Post `
            -Body $json `
            -ContentType "application/json"
    } catch {
        Write-Output "Erro ao enviar update '$titulo': $($_.Exception.Message)"
    }
}

Write-Output "===== Fim da execução: $(Get-Date) ====="
Stop-Transcript
