<#
.SYNOPSIS
    Módulo de funções de relatórios do AD
.DESCRIPTION
    Funções para exportar dados do Active Directory em formatos CSV e HTML.
#>

# Configurações de exportação
$Script:ExportPath = Join-Path $PSScriptRoot "..\reports"
$Script:DateFormat = "dd/MM/yyyy HH:mm"

function Initialize-ExportPath {
    if (-not (Test-Path $Script:ExportPath)) {
        New-Item -Path $Script:ExportPath -ItemType Directory -Force | Out-Null
    }
}

function Export-LockedUsersCSV {
    <#
    .SYNOPSIS
        Exporta lista de usuários bloqueados para CSV
    #>
    
    Show-Header
    Write-Host "  EXPORTAR USUÁRIOS BLOQUEADOS (CSV)" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    try {
        Initialize-ExportPath
        
        $lockedUsers = Search-ADAccount -LockedOut -ErrorAction Stop
        
        if (-not $lockedUsers) {
            Write-Host "  Nenhum usuário bloqueado encontrado." -ForegroundColor Green
            return
        }

        # Buscar detalhes adicionais
        $reportData = foreach ($user in $lockedUsers) {
            $userDetails = Get-ADUser -Identity $user.SamAccountName -Properties LockedOut, lockoutTime, badPwdCount, badPasswordTime, AccountExpirationDate, PasswordLastSet, EmailAddress
            
            [PSCustomObject]@{
                Nome = $user.Name
                Login = $user.SamAccountName
                Email = $userDetails.EmailAddress
                OU = ($user.DistinguishedName -split ',', 2)[1]
                Bloqueado = "Sim"
                TentativasErradas = $userDetails.badPwdCount
                SenhaAlterada = if ($userDetails.PasswordLastSet) { $userDetails.PasswordLastSet.ToString($Script:DateFormat) } else { "-" }
                ContaExpira = if ($userDetails.AccountExpirationDate) { $userDetails.AccountExpirationDate.ToString($Script:DateFormat) } else { "Nunca" }
            }
        }

        $fileName = "locked_users_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $filePath = Join-Path $Script:ExportPath $fileName
        
        $reportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        
        Write-Host "  ✓ $($lockedUsers.Count) usuário(s) exportado(s)" -ForegroundColor Green
        Write-Host "  Arquivo: $filePath" -ForegroundColor Cyan
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Export-UsersByGroupCSV {
    <#
    .SYNOPSIS
        Exporta membros de um grupo para CSV
    #>
    
    Show-Header
    Write-Host "  EXPORTAR USUÁRIOS POR GRUPO (CSV)" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $groupName = Read-Host "  Nome do grupo"
    
    if ([string]::IsNullOrWhiteSpace($groupName)) {
        Write-Host "  Operação cancelada." -ForegroundColor Yellow
        return
    }

    try {
        Initialize-ExportPath
        
        $group = Get-ADGroup -Filter "Name -like '*$groupName*'" | Select-Object -First 1
        
        if (-not $group) {
            Write-Host "  Grupo não encontrado." -ForegroundColor Red
            return
        }

        $members = Get-ADGroupMember -Identity $group -Recursive -ErrorAction Stop
        
        if (-not $members) {
            Write-Host "  Grupo vazio." -ForegroundColor Yellow
            return
        }

        $reportData = foreach ($member in $members) {
            if ($member.objectClass -eq "user") {
                $userDetails = Get-ADUser -Identity $member.SamAccountName -Properties EmailAddress, Enabled, Title, Department -ErrorAction SilentlyContinue
                
                [PSCustomObject]@{
                    Nome = $member.Name
                    Login = $member.SamAccountName
                    Email = if ($userDetails) { $userDetails.EmailAddress } else { "-" }
                    Cargo = if ($userDetails) { $userDetails.Title } else { "-" }
                    Departamento = if ($userDetails) { $userDetails.Department } else { "-" }
                    Ativo = if ($userDetails) { if ($userDetails.Enabled) { "Sim" } else { "Não" } } else { "-" }
                    Tipo = "Usuário"
                    GrupoOrigem = $group.Name
                }
            }
            elseif ($member.objectClass -eq "group") {
                [PSCustomObject]@{
                    Nome = $member.Name
                    Login = $member.SamAccountName
                    Email = "-"
                    Cargo = "-"
                    Departamento = "-"
                    Ativo = "-"
                    Tipo = "Grupo Aninhado"
                    GrupoOrigem = $group.Name
                }
            }
            elseif ($member.objectClass -eq "computer") {
                [PSCustomObject]@{
                    Nome = $member.Name
                    Login = $member.SamAccountName
                    Email = "-"
                    Cargo = "-"
                    Departamento = "-"
                    Ativo = "-"
                    Tipo = "Computador"
                    GrupoOrigem = $group.Name
                }
            }
        }

        $fileName = "group_$($group.SamAccountName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $filePath = Join-Path $Script:ExportPath $fileName
        
        $reportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        
        Write-Host "  ✓ $($members.Count) membro(s) exportado(s)" -ForegroundColor Green
        Write-Host "  Arquivo: $filePath" -ForegroundColor Cyan
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Export-InactiveComputersCSV {
    <#
    .SYNOPSIS
        Exporta lista de computadores inativos para CSV
    #>
    
    Show-Header
    Write-Host "  EXPORTAR COMPUTADORES INATIVOS (CSV)" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $days = Read-Host "  Dias sem login (padrão: 90)"
    if ([string]::IsNullOrWhiteSpace($days)) { $days = 90 }

    try {
        Initialize-ExportPath
        
        $date = (Get-Date).AddDays(-$days)
        $computers = Get-ADComputer -Filter { LastLogonTimeStamp -lt $date -and Enabled -eq $true } -Properties LastLogonTimeStamp, OperatingSystem, DistinguishedName, WhenCreated | Select-Object -First 500
        
        if (-not $computers) {
            Write-Host "  Nenhum computador inativo encontrado." -ForegroundColor Green
            return
        }

        $reportData = foreach ($computer in $computers) {
            $lastLogon = if ($computer.LastLogonTimeStamp) { 
                [DateTime]::FromFileTime($computer.LastLogonTimeStamp).ToString($Script:DateFormat)
                $daysSince = [Math]::Floor(((Get-Date) - [DateTime]::FromFileTime($computer.LastLogonTimeStamp)).TotalDays)
            } else { 
                "Nunca"
                $daysSince = 9999
            }
            
            [PSCustomObject]@{
                Nome = $computer.Name
                DNSHostName = $computer.DNSHostName
                SistemaOperacional = $computer.OperatingSystem
                UltimoLogon = $lastLogon
                DiasInativo = $daysSince
                OU = ($computer.DistinguishedName -split ',', 2)[1]
                CriadoEm = $computer.WhenCreated.ToString($Script:DateFormat)
                Status = "Inativo"
            }
        }

        $fileName = "inactive_computers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $filePath = Join-Path $Script:ExportPath $fileName
        
        $reportData | Sort-Object DiasInativo -Descending | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        
        Write-Host "  ✓ $($computers.Count) computador(es) exportado(s)" -ForegroundColor Green
        Write-Host "  Arquivo: $filePath" -ForegroundColor Cyan
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Export-DomainReportHTML {
    <#
    .SYNOPSIS
        Gera relatório geral do domínio em HTML
    .DESCRIPTION
        Cria um relatório completo do domínio com estatísticas, usuários,
        computadores, grupos e políticas. Formato HTML com estilos CSS.
    #>
    
    Show-Header
    Write-Host "  RELATÓRIO GERAL DO DOMÍNIO (HTML)" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    try {
        Initialize-ExportPath
        
        Write-Host "  Coletando dados do domínio..." -ForegroundColor DarkGray

        # Informações do domínio
        $domain = Get-ADDomain
        $dc = Get-ADDomainController
        $pwdPolicy = Get-ADDefaultDomainPasswordPolicy
        
        # Estatísticas de usuários
        $allUsers = Get-ADUser -Filter * -Properties Enabled, LockedOut, PasswordExpired
        $enabledUsers = ($allUsers | Where-Object { $_.Enabled -eq $true }).Count
        $disabledUsers = ($allUsers | Where-Object { $_.Enabled -eq $false }).Count
        $lockedUsers = ($allUsers | Where-Object { $_.Enabled -eq $true -and $_.LockedOut -eq $true }).Count
        $expiredPasswords = ($allUsers | Where-Object { $_.Enabled -eq $true -and $_.PasswordExpired -eq $true }).Count
        
        # Estatísticas de computadores
        $allComputers = Get-ADComputer -Filter * -Properties Enabled, OperatingSystem, LastLogonTimeStamp
        $enabledComputers = ($allComputers | Where-Object { $_.Enabled -eq $true }).Count
        $disabledComputers = ($allComputers | Where-Object { $_.Enabled -eq $false }).Count
        
        # Contagem por SO
        $osCount = $allComputers | Where-Object { $_.OperatingSystem } | Group-Object OperatingSystem | Sort-Object Count -Descending | Select-Object -First 10
        
        # Computadores inativos (90+ dias)
        $inactiveDate = (Get-Date).AddDays(-90)
        $inactiveComputers = ($allComputers | Where-Object { 
            $_.LastLogonTimeStamp -and [DateTime]::FromFileTime($_.LastLogonTimeStamp) -lt $inactiveDate 
        }).Count
        
        # Grupos
        $allGroups = Get-ADGroup -Filter * -Properties Members
        $emptyGroups = ($allGroups | Where-Object { $_.Members.Count -eq 0 }).Count
        
        # OUs
        $allOUs = Get-ADOrganizationalUnit -Filter *
        
        Write-Host "  Gerando relatório HTML..." -ForegroundColor DarkGray
        
        # HTML do relatório
        $html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório AD - $($domain.Name.ToUpper())</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: #1a1a2e; 
            color: #eee; 
            padding: 20px;
        }
        .container { max-width: 1200px; margin: 0 auto; }
        .header { 
            background: linear-gradient(135deg, #16213e, #0f3460); 
            padding: 30px; 
            border-radius: 10px; 
            margin-bottom: 20px;
            border: 1px solid #0f3460;
        }
        .header h1 { font-size: 2em; margin-bottom: 10px; }
        .header .domain { color: #4fc3f7; font-size: 1.2em; }
        .header .date { color: #888; font-size: 0.9em; margin-top: 10px; }
        
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; margin-bottom: 20px; }
        
        .card { 
            background: #16213e; 
            border-radius: 10px; 
            padding: 20px; 
            border: 1px solid #0f3460;
        }
        .card h2 { 
            color: #4fc3f7; 
            font-size: 1.1em; 
            margin-bottom: 15px; 
            padding-bottom: 10px; 
            border-bottom: 1px solid #0f3460;
        }
        
        .stat { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #1a1a2e; }
        .stat:last-child { border-bottom: none; }
        .stat-label { color: #aaa; }
        .stat-value { font-weight: bold; }
        .stat-value.good { color: #4caf50; }
        .stat-value.warning { color: #ff9800; }
        .stat-value.danger { color: #f44336; }
        .stat-value.info { color: #4fc3f7; }
        
        .section { 
            background: #16213e; 
            border-radius: 10px; 
            padding: 20px; 
            margin-bottom: 20px;
            border: 1px solid #0f3460;
        }
        .section h2 { 
            color: #4fc3f7; 
            font-size: 1.2em; 
            margin-bottom: 15px; 
        }
        
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #1a1a2e; }
        th { color: #4fc3f7; font-weight: normal; }
        tr:hover { background: #1a1a2e; }
        
        .policy-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; }
        .policy-item { 
            background: #1a1a2e; 
            padding: 15px; 
            border-radius: 8px; 
            text-align: center;
        }
        .policy-item .label { color: #888; font-size: 0.85em; }
        .policy-item .value { color: #fff; font-size: 1.5em; font-weight: bold; margin-top: 5px; }
        
        .footer { 
            text-align: center; 
            color: #666; 
            padding: 20px; 
            font-size: 0.85em;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Relatório do Active Directory</h1>
            <div class="domain">Domínio: $($domain.DNSRoot.ToUpper())</div>
            <div class="date">Gerado em: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') | Controlador: $($dc.Name)</div>
        </div>
        
        <div class="grid">
            <div class="card">
                <h2>👥 Usuários</h2>
                <div class="stat">
                    <span class="stat-label">Total</span>
                    <span class="stat-value info">$($allUsers.Count)</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Ativos</span>
                    <span class="stat-value good">$enabledUsers</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Desativados</span>
                    <span class="stat-value warning">$disabledUsers</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Bloqueados</span>
                    <span class="stat-value danger">$lockedUsers</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Senha expirada</span>
                    <span class="stat-value warning">$expiredPasswords</span>
                </div>
            </div>
            
            <div class="card">
                <h2>💻 Computadores</h2>
                <div class="stat">
                    <span class="stat-label">Total</span>
                    <span class="stat-value info">$($allComputers.Count)</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Ativos</span>
                    <span class="stat-value good">$enabledComputers</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Desativados</span>
                    <span class="stat-value warning">$disabledComputers</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Inativos (90+ dias)</span>
                    <span class="stat-value danger">$inactiveComputers</span>
                </div>
            </div>
            
            <div class="card">
                <h2>📁 Grupos</h2>
                <div class="stat">
                    <span class="stat-label">Total</span>
                    <span class="stat-value info">$($allGroups.Count)</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Vazios</span>
                    <span class="stat-value warning">$emptyGroups</span>
                </div>
            </div>
            
            <div class="card">
                <h2>📂 Estrutura</h2>
                <div class="stat">
                    <span class="stat-label">Unidades Organizacionais</span>
                    <span class="stat-value info">$($allOUs.Count)</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Forest</span>
                    <span class="stat-value">$($domain.Forest)</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Nível funcional</span>
                    <span class="stat-value">$($domain.DomainMode)</span>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h2>🔐 Política de Senhas</h2>
            <div class="policy-grid">
                <div class="policy-item">
                    <div class="label">Tamanho mínimo</div>
                    <div class="value">$($pwdPolicy.MinPasswordLength)</div>
                </div>
                <div class="policy-item">
                    <div class="label">Duração (dias)</div>
                    <div class="value">$($pwdPolicy.MaxPasswordAge.Days)</div>
                </div>
                <div class="policy-item">
                    <div class="label">Histórico</div>
                    <div class="value">$($pwdPolicy.PasswordHistoryCount)</div>
                </div>
                <div class="policy-item">
                    <div class="label">Complexidade</div>
                    <div class="value">$(if($pwdPolicy.ComplexityEnabled){'✓'}else{'✗'})</div>
                </div>
                <div class="policy-item">
                    <div class="label">Tentativas bloqueio</div>
                    <div class="value">$($pwdPolicy.LockoutThreshold)</div>
                </div>
                <div class="policy-item">
                    <div class="label">Duração bloqueio (min)</div>
                    <div class="value">$($pwdPolicy.LockoutDuration.Minutes)</div>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h2>🖥️ Sistemas Operacionais</h2>
            <table>
                <thead>
                    <tr>
                        <th>Sistema Operacional</th>
                        <th>Quantidade</th>
                    </tr>
                </thead>
                <tbody>
"@

        foreach ($os in $osCount) {
            $html += @"
                    <tr>
                        <td>$($os.Name)</td>
                        <td>$($os.Count)</td>
                    </tr>
"@
        }

        $html += @"
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            Relatório gerado automaticamente pelo AD Toolkit v1.0.0
        </div>
    </div>
</body>
</html>
"@

        $fileName = "domain_report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        $filePath = Join-Path $Script:ExportPath $fileName
        
        $html | Out-File -FilePath $filePath -Encoding UTF8
        
        Write-Host ""
        Write-Host "  ✓ Relatório gerado com sucesso!" -ForegroundColor Green
        Write-Host "  Arquivo: $filePath" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Resumo:" -ForegroundColor White
        Write-Host "    • $($allUsers.Count) usuários ($lockedUsers bloqueados)" -ForegroundColor DarkGray
        Write-Host "    • $($allComputers.Count) computadores ($inactiveComputers inativos)" -ForegroundColor DarkGray
        Write-Host "    • $($allGroups.Count) grupos ($emptyGroups vazios)" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}
