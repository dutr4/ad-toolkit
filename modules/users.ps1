<#
.SYNOPSIS
    Módulo de funções de usuários do AD
#>

function Get-LockedUsersDetailed {
    <#
    .SYNOPSIS
        Lista usuários bloqueados com causa e origem do bloqueio
    .DESCRIPTION
        Consulta os eventos de segurança para determinar:
        - Causa ESPECÍFICA do bloqueio (senha errada, expirada, conta expirada, etc.)
        - Origem (computador/servidor de onde veio)
        - Tipo de logon (interativo, rede, RDP, e-mail, etc.)
        - Timestamp
        - Número de tentativas
        - Se foi auto-bloqueio ou ação de admin
    #>
    
    Show-Header
    Write-Host "  USUÁRIOS BLOQUEADOS" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    # Função auxiliar para traduzir substatus de erro
    function Get-LockoutReason {
        param([string]$SubStatus)
        
        switch ($SubStatus) {
            "0xC000006A" { return "Senha incorreta"; break }
            "0xC000006D" { return "Logon falhou (motivo não especificado)"; break }
            "0xC000006E" { return "Restrição de conta (horário/workstation)"; break }
            "0xC000006F" { return "Logon fora do horário permitido"; break }
            "0xC0000070" { return "Logon de workstation não autorizada"; break }
            "0xC0000071" { return "Senha expirada"; break }
            "0xC0000072" { return "Conta expirada"; break }
            "0xC00000DC" { return "Status de conta inválido"; break }
            "0xC0000133" { return "Relógio do DC fora de sync"; break }
            "0xC000015B" { return "Tipo de logon não permitido"; break }
            "0xC000018C" { return "Erro de trust account"; break }
            "0xC0000193" { return "Conta expirada"; break }
            "0xC0000224" { return "Senha deve ser alterada no próximo logon"; break }
            "0xC0000225" { return "Erro do Windows (bug)"; break }
            "0xC0000234" { return "Conta bloqueada por tentativas incorretas"; break }
            default { return "Motivo não identificado ($SubStatus)"; break }
        }
    }
    
    # Função auxiliar para traduzir tipo de logon
    function Get-LogonType {
        param([string]$Type)
        
        switch ($Type) {
            "2"  { return "Interativo (local/RDP)"; break }
            "3"  { return "Rede (compartilhamento, SQL)"; break }
            "4"  { return "Batch (tarefa agendada)"; break }
            "5"  { return "Serviço Windows"; break }
            "7"  { return "Desbloqueio de tela"; break }
            "8"  { return "Rede cleartext (IIS/Exchange)"; break }
            "9"  { return "NewCredentials (runas /netonly)"; break }
            "10" { return "RemoteInteractive (RDP)"; break }
            "11" { return "CachedInteractive (offline)"; break }
            default { return "Tipo $Type"; break }
        }
    }

    try {
        # Buscar usuários bloqueados
        $lockedUsers = Search-ADAccount -LockedOut -ErrorAction Stop
        
        if (-not $lockedUsers) {
            Write-Host "  Nenhum usuário bloqueado no momento." -ForegroundColor Green
            return
        }

        Write-Host "  Encontrados $($lockedUsers.Count) usuário(s) bloqueado(s):" -ForegroundColor Cyan
        Write-Host ""

        $PDC = (Get-ADDomainController -Discover -Service PrimaryDC).Name

        foreach ($user in $lockedUsers) {
            $userDetails = Get-ADUser -Identity $user.SamAccountName -Properties LockedOut, lockoutTime, badPwdCount, badPasswordTime, AccountExpirationDate, AccountLockoutTime
            
            Write-Host "  ┌─────────────────────────────────────────────────" -ForegroundColor DarkGray
            Write-Host "  │ Nome:       $($user.Name)" -ForegroundColor White
            Write-Host "  │ Login:      $($user.SamAccountName)" -ForegroundColor White
            Write-Host "  │ OU:         $(($user.DistinguishedName -split ',', 2)[1])" -ForegroundColor DarkGray
            
            # Buscar evento de bloqueio (4740) no PDC
            $lockoutEvent = Get-WinEvent -ComputerName $PDC -FilterHashtable @{
                LogName = 'Security'
                ID = 4740
            } -MaxEvents 100 -ErrorAction SilentlyContinue | 
            Where-Object { $_.Message -match $user.SamAccountName } | 
            Select-Object -First 1

            # Buscar eventos de falha de logon (4625) para este usuário
            $failedLogonEvents = Get-WinEvent -ComputerName $PDC -FilterHashtable @{
                LogName = 'Security'
                ID = 4625
            } -MaxEvents 200 -ErrorAction SilentlyContinue | 
            Where-Object { $_.Message -match $user.SamAccountName } | 
            Select-Object -First 5

            if ($lockoutEvent) {
                $eventXML = [xml]$lockoutEvent.ToXml()
                $callerComputer = $eventXML.Event.EventData.Data | Where-Object { $_.Name -eq 'CallerComputerName' }
                
                Write-Host "  │ Bloqueado em: $($lockoutEvent.TimeCreated.ToString('dd/MM/yyyy HH:mm:ss'))" -ForegroundColor Yellow
                
                if ($callerComputer -and $callerComputer.'#text') {
                    Write-Host "  │ Origem:      $($callerComputer.'#text')" -ForegroundColor Red
                }
            }

            # Analisar eventos de falha para determinar causa
            if ($failedLogonEvents) {
                $latestFailure = $failedLogonEvents[0]
                $failureXML = [xml]$latestFailure.ToXml()
                
                $subStatus = ($failureXML.Event.EventData.Data | Where-Object { $_.Name -eq 'SubStatus' }).'#text'
                $logonType = ($failureXML.Event.EventData.Data | Where-Object { $_.Name -eq 'LogonType' }).'#text'
                $targetMachine = ($failureXML.Event.EventData.Data | Where-Object { $_.Name -eq 'WorkstationName' }).'#text'
                $sourceIP = ($failureXML.Event.EventData.Data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'
                
                $reason = Get-LockoutReason -SubStatus $subStatus
                $logonTypeStr = Get-LogonType -Type $logonType
                
                Write-Host "  │ Causa:      $reason" -ForegroundColor $(if ($subStatus -eq "0xC0000071" -or $subStatus -eq "0xC0000193") { "Red" } else { "Yellow" })
                Write-Host "  │ Tipo logon: $logonTypeStr" -ForegroundColor DarkGray
                
                if ($targetMachine) {
                    Write-Host "  │ Máquina:    $targetMachine" -ForegroundColor DarkGray
                }
                if ($sourceIP -and $sourceIP -ne "-" -and $sourceIP -ne "::1" -and $sourceIP -ne "127.0.0.1") {
                    Write-Host "  │ IP origem:  $sourceIP" -ForegroundColor DarkGray
                }
                
                # Contar tentativas recentes
                $recentFailures = ($failedLogonEvents | Measure-Object).Count
                Write-Host "  │ Tentativas recentes visíveis: $recentFailures" -ForegroundColor DarkGray
            }
            elseif (-not $lockoutEvent) {
                Write-Host "  │ Detalhes não encontrados nos logs" -ForegroundColor DarkGray
                Write-Host "  │ (Eventos podem ter sido sobrescritos)" -ForegroundColor DarkGray
            }

            # Verificar status da senha
            $pwdLastSet = $userDetails.PasswordLastSet
            $maxPwdAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
            
            if ($pwdLastSet) {
                $pwdAge = (Get-Date) - $pwdLastSet
                if ($pwdAge.Days -ge $maxPwdAge) {
                    Write-Host "  │ ⚠ Senha EXPIRADA há $($pwdAge.Days - $maxPwdAge) dia(s)" -ForegroundColor Red
                }
                elseif ($pwdAge.Days -ge ($maxPwdAge - 7)) {
                    Write-Host "  │ ⚠ Senha expira em $($maxPwdAge - $pwdAge.Days) dia(s)" -ForegroundColor Yellow
                }
            }

            # Verificar se a conta expirou
            if ($userDetails.AccountExpirationDate -and $userDetails.AccountExpirationDate -lt (Get-Date)) {
                Write-Host "  │ ⚠ Conta EXPIRADA em $($userDetails.AccountExpirationDate.ToString('dd/MM/yyyy'))" -ForegroundColor Red
            }

            Write-Host "  └─────────────────────────────────────────────────" -ForegroundColor DarkGray
            Write-Host ""
        }

        Write-Host "  Dica: Use a opção [2] para desbloquear um usuário." -ForegroundColor DarkGray

    }
    catch {
        Write-Host "  Erro ao buscar usuários: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Unlock-ADUserInteractive {
    <#
    .SYNOPSIS
        Desbloqueia um usuário interativamente
    #>
    
    Show-Header
    Write-Host "  DESBLOQUEAR USUÁRIO" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $username = Read-Host "  Digite o login do usuário (ou parte do nome)"
    
    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host "  Operação cancelada." -ForegroundColor Yellow
        return
    }

    try {
        # Buscar usuário
        $users = Get-ADUser -Filter "SamAccountName -like '*$username*' -or Name -like '*$username*'" -Properties LockedOut |
                 Where-Object { $_.LockedOut -eq $true }
        
        if (-not $users) {
            Write-Host "  Nenhum usuário bloqueado encontrado com esse nome." -ForegroundColor Yellow
            return
        }

        if ($users.Count -gt 1) {
            Write-Host "  Múltiplos usuários encontrados:" -ForegroundColor Cyan
            Write-Host ""
            $i = 1
            foreach ($u in $users) {
                Write-Host "  [$i] $($u.Name) ($($u.SamAccountName))" -ForegroundColor White
                $i++
            }
            Write-Host ""
            $selection = Read-Host "  Selecione o número"
            $user = $users[$selection - 1]
        }
        else {
            $user = $users
        }

        Write-Host ""
        Write-Host "  Desbloqueando $($user.Name)..." -ForegroundColor Cyan
        Unlock-ADAccount -Identity $user.SamAccountName
        Write-Host "  ✓ Usuário desbloqueado com sucesso!" -ForegroundColor Green

        # Log da ação
        Write-EventLog -LogName "Application" -Source "AD-Toolkit" -EntryType Information -EventId 1001 -Message "Usuário $($user.SamAccountName) desbloqueado por $env:USERNAME" -ErrorAction SilentlyContinue

    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-UsersByGroup {
    <#
    .SYNOPSIS
        Lista usuários de um grupo específico
    #>
    
    Show-Header
    Write-Host "  USUÁRIOS POR GRUPO" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $groupName = Read-Host "  Digite o nome do grupo"
    
    if ([string]::IsNullOrWhiteSpace($groupName)) {
        Write-Host "  Operação cancelada." -ForegroundColor Yellow
        return
    }

    try {
        $group = Get-ADGroup -Filter "Name -like '*$groupName*'" | Select-Object -First 1
        
        if (-not $group) {
            Write-Host "  Grupo não encontrado." -ForegroundColor Red
            return
        }

        Write-Host ""
        Write-Host "  Grupo: $($group.Name)" -ForegroundColor Cyan
        Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
        
        $members = Get-ADGroupMember -Identity $group -ErrorAction Stop
        
        if (-not $members) {
            Write-Host "  Grupo vazio." -ForegroundColor Yellow
            return
        }

        Write-Host ""
        Write-Host "  $($members.Count) membro(s):" -ForegroundColor White
        Write-Host ""

        foreach ($member in $members | Sort-Object Name) {
            $type = switch ($member.objectClass) {
                "user" { "👤" }
                "group" { "📁" }
                "computer" { "💻" }
                default { "❓" }
            }
            Write-Host "  $type $($member.Name)" -ForegroundColor White
        }
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Search-ADUsers {
    <#
    .SYNOPSIS
        Busca usuários por filtro
    #>
    
    Show-Header
    Write-Host "  BUSCAR USUÁRIOS" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Filtros disponíveis:" -ForegroundColor DarkGray
    Write-Host "  [1] Por nome" -ForegroundColor White
    Write-Host "  [2] Por login (SamAccountName)" -ForegroundColor White
    Write-Host "  [3] Por OU" -ForegroundColor White
    Write-Host "  [4] Por e-mail" -ForegroundColor White
    Write-Host ""
    
    $filterType = Read-Host "  Tipo de filtro"
    $filterValue = Read-Host "  Valor"
    
    if ([string]::IsNullOrWhiteSpace($filterValue)) {
        Write-Host "  Operação cancelada." -ForegroundColor Yellow
        return
    }

    try {
        $filter = switch ($filterType) {
            "1" { "Name -like '*$filterValue*'" }
            "2" { "SamAccountName -like '*$filterValue*'" }
            "3" { "DistinguishedName -like '*$filterValue*'" }
            "4" { "EmailAddress -like '*$filterValue*'" }
            default { "Name -like '*$filterValue*'" }
        }

        $users = Get-ADUser -Filter $filter -Properties EmailAddress, Enabled, LockedOut | Select-Object -First 50
        
        if (-not $users) {
            Write-Host "  Nenhum usuário encontrado." -ForegroundColor Yellow
            return
        }

        Write-Host ""
        Write-Host "  $($users.Count) usuário(s) encontrado(s):" -ForegroundColor Cyan
        Write-Host ""

        foreach ($user in $users | Sort-Object Name) {
            $status = if ($user.Enabled) { "✓" } else { "✗" }
            $lock = if ($user.LockedOut) { "🔒" } else { "" }
            Write-Host "  $status $($user.Name) ($($user.SamAccountName)) $lock" -ForegroundColor $(if ($user.Enabled) { "White" } else { "DarkGray" })
        }
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-UsersPasswordExpiring {
    <#
    .SYNOPSIS
        Lista usuários com senha prestes a expirar
    #>
    
    Show-Header
    Write-Host "  SENHAS PRESTES A EXPIRAR" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $days = Read-Host "  Dias até expirar (padrão: 7)"
    if ([string]::IsNullOrWhiteSpace($days)) { $days = 7 }

    try {
        $maxPwdAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
        
        $users = Get-ADUser -Filter { Enabled -eq $true -and PasswordNeverExpires -eq $false } -Properties PasswordLastSet, PasswordNeverExpires, PasswordExpired
        
        $expiringUsers = foreach ($user in $users) {
            if ($user.PasswordLastSet) {
                $daysUntilExpire = $maxPwdAge - ((Get-Date) - $user.PasswordLastSet).Days
                if ($daysUntilExpire -le $days -and $daysUntilExpire -gt 0) {
                    [PSCustomObject]@{
                        Name = $user.Name
                        SamAccountName = $user.SamAccountName
                        DaysUntilExpire = $daysUntilExpire
                        PasswordLastSet = $user.PasswordLastSet.ToString('dd/MM/yyyy')
                    }
                }
            }
        }

        if (-not $expiringUsers) {
            Write-Host "  Nenhum usuário com senha expirando em $days dias." -ForegroundColor Green
            return
        }

        Write-Host "  $($expiringUsers.Count) usuário(s) com senha expirando:" -ForegroundColor Yellow
        Write-Host ""

        $expiringUsers | Sort-Object DaysUntilExpire | ForEach-Object {
            $color = if ($_.DaysUntilExpire -le 3) { "Red" } elseif ($_.DaysUntilExpire -le 5) { "Yellow" } else { "White" }
            Write-Host "  $($_.DaysUntilExpire) dia(s) - $($_.Name) ($($_.SamAccountName))" -ForegroundColor $color
        }
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-InactiveUsers {
    <#
    .SYNOPSIS
        Lista usuários inativos (sem login há X dias)
    #>
    
    Show-Header
    Write-Host "  USUÁRIOS INATIVOS" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $days = Read-Host "  Dias sem login (padrão: 90)"
    if ([string]::IsNullOrWhiteSpace($days)) { $days = 90 }

    try {
        $date = (Get-Date).AddDays(-$days)
        $users = Get-ADUser -Filter { LastLogonTimeStamp -lt $date -and Enabled -eq $true } -Properties LastLogonTimeStamp, WhenCreated | Select-Object -First 100
        
        if (-not $users) {
            Write-Host "  Nenhum usuário inativo encontrado." -ForegroundColor Green
            return
        }

        Write-Host "  $($users.Count) usuário(s) sem login há mais de $days dias:" -ForegroundColor Yellow
        Write-Host ""

        foreach ($user in $users | Sort-Object Name) {
            $lastLogon = if ($user.LastLogonTimeStamp) { 
                [DateTime]::FromFileTime($user.LastLogonTimeStamp).ToString('dd/MM/yyyy') 
            } else { 
                "Nunca" 
            }
            Write-Host "  $($user.Name) - Último login: $lastLogon" -ForegroundColor White
        }
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}
