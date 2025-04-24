# Atualizador.ps1 - Atualizações do Windows + log local (clássico) + envio remoto

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Caminhos
$logPath = "C:\Appmax"
$logFile = Join-Path $logPath "update-log.txt"
$csvFile = Join-Path $logPath "report.csv"

# Criar diretório, se necessário
if (-Not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
}

# Cabeçalho do log
Add-Content -Path $logFile -Value "`n===== Início da execução: $(Get-Date) =====" -Encoding utf8

# Verifica e instala PSWindowsUpdate, se necessário
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Install-Module -Name PSWindowsUpdate -Force
    Add-Content -Path $logFile -Value "Módulo PSWindowsUpdate instalado com sucesso." -Encoding utf8
} else {
    Add-Content -Path $logFile -Value "Módulo PSWindowsUpdate já instalado." -Encoding utf8
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

# Executar atualizações e registrar a saída detalhada diretamente no log
Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -ForceDownload -ForceInstall -IgnoreReboot -Verbose |
    Tee-Object -FilePath $logFile -Append

# Coleta os updates instalados com sucesso via histórico
$updatesInstalled = Get-WUHistory | Where-Object {$_.Result -eq "Succeeded"} | Sort-Object Date -Descending | Select-Object -First 20

# Envia para CSV e planilha
foreach ($update in $updatesInstalled) {
    $titulo = $update.Title
    if (-not $titulo) { continue }

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
        Add-Content -Path $logFile -Value "Erro ao enviar update '$titulo': $($_.Exception.Message)" -Encoding utf8
    }
}

# Fim do log
Add-Content -Path $logFile -Value "===== Fim da execução: $(Get-Date) =====`n" -Encoding utf8
