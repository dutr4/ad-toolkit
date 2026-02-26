<#
.SYNOPSIS
    AD Toolkit - Ferramenta de gestão do Active Directory
.DESCRIPTION
    Interface de linha de comando com menus para gerenciar usuários, grupos e computadores do AD.
.AUTHOR
    Guilherme Dutra Campos
.VERSION
    1.0.0
#>

# Configurações
$Script:Version = "1.0.0"
$Script:Title = @"
█████╗ ██████╗     ████████╗ ██████╗  ██████╗ ██╗     ██╗  ██╗██╗████████╗
██╔══██╗██╔══██╗    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██║ ██╔╝██║╚══██╔══╝
███████║██║  ██║       ██║   ██║   ██║██║   ██║██║     █████╔╝ ██║   ██║   
██╔══██║██║  ██║       ██║   ██║   ██║██║   ██║██║     ██╔═██╗ ██║   ██║   
██║  ██║██████╔╝       ██║   ╚██████╔╝╚██████╔╝███████╗██║  ██╗██║   ██║   
╚═╝  ╚═╝╚═════╝        ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝   ╚═╝   
"@

# Importar módulos
$ModulesPath = Join-Path $PSScriptRoot "modules"
Get-ChildItem -Path $ModulesPath -Filter "*.ps1" | ForEach-Object {
    . $_.FullName
}

# Funções de menu
function Show-Header {
    Clear-Host
    Write-Host $Script:Title -ForegroundColor Cyan
    Write-Host "  Versão $Script:Version | Guilherme Dutra Campos" -ForegroundColor DarkGray
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

function Show-MainMenu {
    Show-Header
    Write-Host "  MENU PRINCIPAL" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Usuários" -ForegroundColor White
    Write-Host "  [2] Grupos" -ForegroundColor White
    Write-Host "  [3] Computadores" -ForegroundColor White
    Write-Host "  [4] Relatórios" -ForegroundColor White
    Write-Host "  [5] Configurações" -ForegroundColor White
    Write-Host ""
    Write-Host "  [Q] Sair" -ForegroundColor Gray
    Write-Host ""
}

function Show-UsersMenu {
    Show-Header
    Write-Host "  USUÁRIOS" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Listar usuários bloqueados (detalhado)" -ForegroundColor White
    Write-Host "  [2] Desbloquear usuário" -ForegroundColor White
    Write-Host "  [3] Listar usuários por grupo" -ForegroundColor White
    Write-Host "  [4] Buscar usuários por filtro" -ForegroundColor White
    Write-Host "  [5] Usuários com senha prestes a expirar" -ForegroundColor White
    Write-Host "  [6] Usuários inativos (sem login há X dias)" -ForegroundColor White
    Write-Host ""
    Write-Host "  [0] Voltar" -ForegroundColor Gray
    Write-Host ""
}

function Show-GroupsMenu {
    Show-Header
    Write-Host "  GRUPOS" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Listar membros de um grupo" -ForegroundColor White
    Write-Host "  [2] Adicionar usuário a grupo" -ForegroundColor White
    Write-Host "  [3] Remover usuário de grupo" -ForegroundColor White
    Write-Host "  [4] Grupos vazios" -ForegroundColor White
    Write-Host ""
    Write-Host "  [0] Voltar" -ForegroundColor Gray
    Write-Host ""
}

function Show-ComputersMenu {
    Show-Header
    Write-Host "  COMPUTADORES" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Computadores inativos (90+ dias)" -ForegroundColor White
    Write-Host "  [2] Forçar GPUpdate remoto" -ForegroundColor White
    Write-Host "  [3] Listar computadores por OU" -ForegroundColor White
    Write-Host "  [4] Status de computador específico" -ForegroundColor White
    Write-Host ""
    Write-Host "  [0] Voltar" -ForegroundColor Gray
    Write-Host ""
}

function Show-ReportsMenu {
    Show-Header
    Write-Host "  RELATÓRIOS" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Exportar usuários bloqueados (CSV)" -ForegroundColor White
    Write-Host "  [2] Exportar usuários por grupo (CSV)" -ForegroundColor White
    Write-Host "  [3] Exportar computadores inativos (CSV)" -ForegroundColor White
    Write-Host "  [4] Relatório geral do domínio (HTML)" -ForegroundColor White
    Write-Host ""
    Write-Host "  [0] Voltar" -ForegroundColor Gray
    Write-Host ""
}

function Get-MenuChoice {
    param(
        [string]$Prompt = "  Opção"
    )
    Write-Host ""
    Write-Host "  $Prompt" -NoNewline -ForegroundColor Cyan
    Write-Host ": " -NoNewline
    return Read-Host
}

function Wait-KeyPress {
    Write-Host ""
    Write-Host "  Pressione qualquer tecla para continuar..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Loop principal
function Start-ADToolkit {
    # Verificar se o módulo AD está disponível
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Show-Header
        Write-Host "  [ERRO] Módulo Active Directory não encontrado!" -ForegroundColor Red
        Write-Host "  Instale com: Install-WindowsFeature RSAT-AD-PowerShell" -ForegroundColor Yellow
        Wait-KeyPress
        return
    }

    Import-Module ActiveDirectory -ErrorAction SilentlyContinue

    $running = $true
    while ($running) {
        Show-MainMenu
        $choice = Get-MenuChoice

        switch ($choice.ToUpper()) {
            "1" {
                # Menu Usuários
                $submenu = $true
                while ($submenu) {
                    Show-UsersMenu
                    $subchoice = Get-MenuChoice

                    switch ($subchoice) {
                        "1" { Get-LockedUsersDetailed; Wait-KeyPress }
                        "2" { Unlock-ADUserInteractive; Wait-KeyPress }
                        "3" { Get-UsersByGroup; Wait-KeyPress }
                        "4" { Search-ADUsers; Wait-KeyPress }
                        "5" { Get-UsersPasswordExpiring; Wait-KeyPress }
                        "6" { Get-InactiveUsers; Wait-KeyPress }
                        "0" { $submenu = $false }
                        default { Write-Host "  Opção inválida!" -ForegroundColor Red; Start-Sleep -Seconds 1 }
                    }
                }
            }
            "2" {
                # Menu Grupos
                $submenu = $true
                while ($submenu) {
                    Show-GroupsMenu
                    $subchoice = Get-MenuChoice

                    switch ($subchoice) {
                        "1" { Get-GroupMembers; Wait-KeyPress }
                        "2" { Add-UserToGroupInteractive; Wait-KeyPress }
                        "3" { Remove-UserFromGroupInteractive; Wait-KeyPress }
                        "4" { Get-EmptyGroups; Wait-KeyPress }
                        "0" { $submenu = $false }
                        default { Write-Host "  Opção inválida!" -ForegroundColor Red; Start-Sleep -Seconds 1 }
                    }
                }
            }
            "3" {
                # Menu Computadores
                $submenu = $true
                while ($submenu) {
                    Show-ComputersMenu
                    $subchoice = Get-MenuChoice

                    switch ($subchoice) {
                        "1" { Get-InactiveComputers; Wait-KeyPress }
                        "2" { Invoke-GPUpdateRemote; Wait-KeyPress }
                        "3" { Get-ComputersByOU; Wait-KeyPress }
                        "4" { Get-ComputerStatus; Wait-KeyPress }
                        "0" { $submenu = $false }
                        default { Write-Host "  Opção inválida!" -ForegroundColor Red; Start-Sleep -Seconds 1 }
                    }
                }
            }
            "4" {
                # Menu Relatórios
                $submenu = $true
                while ($submenu) {
                    Show-ReportsMenu
                    $subchoice = Get-MenuChoice

                    switch ($subchoice) {
                        "1" { Export-LockedUsersCSV; Wait-KeyPress }
                        "2" { Export-UsersByGroupCSV; Wait-KeyPress }
                        "3" { Export-InactiveComputersCSV; Wait-KeyPress }
                        "4" { Export-DomainReportHTML; Wait-KeyPress }
                        "0" { $submenu = $false }
                        default { Write-Host "  Opção inválida!" -ForegroundColor Red; Start-Sleep -Seconds 1 }
                    }
                }
            }
            "5" {
                Show-Header
                Write-Host "  CONFIGURAÇÕES" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Domínio: $env:USERDNSDOMAIN" -ForegroundColor White
                Write-Host "  Controlador: $((Get-ADDomainController).Name)" -ForegroundColor White
                Wait-KeyPress
            }
            "Q" {
                $running = $false
                Show-Header
                Write-Host "  Até mais!" -ForegroundColor Green
            }
            default {
                Write-Host "  Opção inválida!" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    }
}

# Iniciar
Start-ADToolkit
