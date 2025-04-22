# Atualizador.ps1 - Script de atualização do Windows com PSWindowsUpdate
# Criado por Appmax TI - Versão sem reinício automático

# Força o uso de TLS 1.2 para conexões seguras com repositórios do PowerShell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Define diretório para logs
$logPath = "C:\Appmax"
$logFile = Join-Path $logPath "update-log.txt"

# Cria diretório de log se não existir
if (-Not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
}

# Início do log
Add-Content -Path $logFile -Value "`n===== Início da execução: $(Get-Date) =====" -Encoding utf8

# Verifica se o módulo PSWindowsUpdate está instalado
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Output "Módulo PSWindowsUpdate não encontrado. Instalando..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
    Install-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
    Add-Content -Path $logFile -Value "Módulo PSWindowsUpdate instalado com sucesso."
} else {
    Add-Content -Path $logFile -Value "Módulo PSWindowsUpdate já instalado."
}

# Importa o módulo
Import-Module PSWindowsUpdate -Force

# Habilita o Microsoft Update (Office, Defender, etc)
Add-WUServiceManager -MicrosoftUpdate -Confirm:$false

# Lista atualizações disponíveis
$updates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll

if ($updates.Count -eq 0) {
    Add-Content -Path $logFile -Value "Nenhuma atualização pendente encontrada."
} else {
    Add-Content -Path $logFile -Value "$($updates.Count) atualização(ões) encontrada(s). Iniciando instalação..."

    # Instala todas as atualizações encontradas
    Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -IgnoreReboot -Verbose |
        Out-File -FilePath $logFile -Encoding utf8 -Append
}

# Fim do log
Add-Content -Path $logFile -Value "===== Fim da execução: $(Get-Date) =====`n" -Encoding utf8
