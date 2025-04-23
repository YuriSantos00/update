# Atualizador.ps1 - Atualizações + Logs locais + Envio para Google Sheets

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Caminhos
$logPath = "C:\Appmax"
$logFile = Join-Path $logPath "update-log.txt"
$csvFile = Join-Path $logPath "report.csv"

# Garantir diretório
if (-Not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
}

# Início do log
Add-Content -Path $logFile -Value "`n===== Início da execução: $(Get-Date) =====" -Encoding utf8

# Verifica e instala PSWindowsUpdate se necessário
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

# Busca atualizações
$updates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll

if ($updates.Count -eq 0) {
    Add-Content -Path $logFile -Value "Nenhuma atualização pendente encontrada." -Encoding utf8
} else {
    Add-Content -Path $logFile -Value "$($updates.Count) atualização(ões) encontrada(s). Iniciando instalação..." -Encoding utf8

    # Instala e registra no log
    Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -IgnoreReboot -Verbose |
        Out-File -FilePath $logFile -Encoding utf8 -Append
}

# REGISTRO LOCAL E ENVIO PARA GOOGLE SHEETS
$hostname = $env:COMPUTERNAME
$user = (Get-WmiObject Win32_ComputerSystem).UserName
$date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Coletar histórico de atualizações aplicadas com sucesso
$updatesInstalled = Get-WUHistory | Where-Object {$_.Result -eq "Succeeded"} | Sort-Object Date -Descending | Select-Object -First 10

# Criar CSV local com cabeçalho se necessário
if (-not (Test-Path $csvFile)) {
    Add-Content -Path $csvFile -Value "Data,Hostname,Usuário,Título da Atualização" -Encoding utf8
}

foreach ($update in $updatesInstalled) {
    $linha = "$date,$hostname,$user,""$($update.Title)"""
    Add-Content -Path $csvFile -Value $linha -Encoding utf8

    # Montar JSON para envio remoto
    $json = @{
        data     = $date
        hostname = $hostname
        usuario  = $user
        titulo   = $update.Title
    } | ConvertTo-Json -Depth 3

    # Enviar para o Web App
    Invoke-RestMethod -Uri "https://script.google.com/macros/s/AKfycby7UBZ4jFH10wmHC7KxYB6ZTFbeUfZcdFAoz5X3L9ln0CfomJ1Xtfqhpu14P6vlLVQ/exec" `
        -Method Post `
        -Body $json `
        -ContentType "application/json"
}

# Fim do log
Add-Content -Path $logFile -Value "===== Fim da execução: $(Get-Date) =====`n" -Encoding utf8
