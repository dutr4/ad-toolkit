<#
.SYNOPSIS
    Módulo de funções de computadores do AD
#>

function Get-InactiveComputers {
    <#
    .SYNOPSIS
        Lista computadores inativos (sem login há X dias)
    .DESCRIPTION
        Identifica máquinas que não se conectaram ao domínio há mais do que o período especificado.
        Útil para identificar equipamentos desativados, offboarding ou problemas de replicação.
    #>
    
    Show-Header
    Write-Host "  COMPUTADORES INATIVOS" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $days = Read-Host "  Dias sem login (padrão: 90)"
    if ([string]::IsNullOrWhiteSpace($days)) { $days = 90 }

    try {
        $date = (Get-Date).AddDays(-$days)
        
        # Buscar computadores inativos
        $computers = Get-ADComputer -Filter { LastLogonTimeStamp -lt $date -and Enabled -eq $true } -Properties LastLogonTimeStamp, OperatingSystem, DistinguishedName, WhenCreated | Select-Object -First 200
        
        if (-not $computers) {
            Write-Host "  Nenhum computador inativo encontrado." -ForegroundColor Green
            return
        }

        Write-Host "  $($computers.Count) computador(es) inativo(s) há mais de $days dias:" -ForegroundColor Yellow
        Write-Host ""

        # Agrupar por OU para melhor visualização
        $byOU = $computers | Group-Object { ($_ | Get-ADOrganizationalUnit -ErrorAction SilentlyContinue).Name } | Sort-Object Count -Descending

        foreach ($computer in $computers | Sort-Object Name) {
            $lastLogon = if ($computer.LastLogonTimeStamp) { 
                [DateTime]::FromFileTime($computer.LastLogonTimeStamp).ToString('dd/MM/yyyy') 
            } else { 
                "Nunca" 
            }
            
            $os = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Desconhecido" }
            $ouName = ($computer.DistinguishedName -split ',', 2)[1]
            
            Write-Host "  💻 $($computer.Name)" -ForegroundColor White
            Write-Host "     Último login: $lastLogon | OS: $os" -ForegroundColor DarkGray
            Write-Host "     OU: $ouName" -ForegroundColor DarkGray
            Write-Host ""
        }

        Write-Host "  Dica: Computadores inativos podem ser desativados ou removidos." -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-GPUpdateRemote {
    <#
    .SYNOPSIS
        Força atualização de GPO em computadores remotos
    .DESCRIPTION
        Executa gpupdate /force remotamente em um ou mais computadores.
        Requer privilégios de administrador e WinRM habilitado nos destinos.
    #>
    
    Show-Header
    Write-Host "  GPUPDATE REMOTO" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Opções:" -ForegroundColor DarkGray
    Write-Host "  [1] Um computador específico" -ForegroundColor White
    Write-Host "  [2] Todos os computadores de uma OU" -ForegroundColor White
    Write-Host "  [3] Múltiplos computadores (lista)" -ForegroundColor White
    Write-Host ""
    
    $option = Read-Host "  Opção"
    
    $computerNames = @()
    
    switch ($option) {
        "1" {
            $computerName = Read-Host "  Nome do computador"
            if (-not [string]::IsNullOrWhiteSpace($computerName)) {
                $computerNames += $computerName
            }
        }
        "2" {
            $ouPath = Read-Host "  OU (ex: OU=Computadores,DC=empresa,DC=local)"
            if (-not [string]::IsNullOrWhiteSpace($ouPath)) {
                try {
                    $computers = Get-ADComputer -Filter { Enabled -eq $true } -SearchBase $ouPath | Select-Object -First 50
                    $computerNames = $computers.Name
                }
                catch {
                    Write-Host "  Erro ao buscar computadores: $($_.Exception.Message)" -ForegroundColor Red
                    return
                }
            }
        }
        "3" {
            Write-Host "  Digite os nomes (um por linha, linha vazia para terminar):" -ForegroundColor DarkGray
            do {
                $name = Read-Host "  "
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $computerNames += $name
                }
            } until ([string]::IsNullOrWhiteSpace($name))
        }
        default {
            Write-Host "  Opção inválida." -ForegroundColor Red
            return
        }
    }
    
    if ($computerNames.Count -eq 0) {
        Write-Host "  Nenhum computador especificado." -ForegroundColor Yellow
        return
    }

    Write-Host ""
    Write-Host "  Iniciando GPUpdate em $($computerNames.Count) computador(es)..." -ForegroundColor Cyan
    Write-Host ""

    $success = 0
    $failed = 0

    foreach ($computer in $computerNames) {
        Write-Host "  $($computer)... " -NoNewline -ForegroundColor White
        
        try {
            # Verificar se o computador está online
            if (Test-Connection -ComputerName $computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                # Executar GPUpdate remotamente via Invoke-Command
                $result = Invoke-Command -ComputerName $computer -ScriptBlock {
                    gpupdate /force
                } -ErrorAction Stop -TimeoutSeconds 60
                
                Write-Host "✓ OK" -ForegroundColor Green
                $success++
            }
            else {
                Write-Host "✗ Offline" -ForegroundColor Red
                $failed++
            }
        }
        catch {
            Write-Host "✗ Erro: $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host ""
    Write-Host "  Resumo: $success sucesso(s), $failed falha(s)" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Yellow" })
}
catch {
    Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
}
}

function Get-ComputersByOU {
    <#
    .SYNOPSIS
        Lista computadores por Unidade Organizacional
    .DESCRIPTION
        Exibe todos os computadores de uma OU específica com informações básicas.
    #>
    
    Show-Header
    Write-Host "  COMPUTADORES POR OU" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    try {
        # Listar OUs disponíveis
        Write-Host "  Buscando OUs..." -ForegroundColor DarkGray
        $ous = Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Sort-Object Name
        
        Write-Host ""
        Write-Host "  OUs disponíveis:" -ForegroundColor Cyan
        $i = 1
        foreach ($ou in $ous | Select-Object -First 20) {
            Write-Host "  [$i] $($ou.Name)" -ForegroundColor White
            $i++
        }
        
        if ($ous.Count -gt 20) {
            Write-Host "  ... e mais $($ous.Count - 20) OUs" -ForegroundColor DarkGray
        }
        
        Write-Host ""
        Write-Host "  [T] Digitar OU manualmente" -ForegroundColor White
        Write-Host ""
        
        $selection = Read-Host "  Selecione"
        
        $ouPath = $null
        
        if ($selection -eq "T" -or $selection -eq "t") {
            $ouPath = Read-Host "  Caminho da OU (DistinguishedName)"
        }
        elseif ($selection -match '^\d+$' -and [int]$selection -le $ous.Count) {
            $ouPath = $ous[[int]$selection - 1].DistinguishedName
        }
        else {
            Write-Host "  Seleção inválida." -ForegroundColor Red
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($ouPath)) {
            Write-Host "  Operação cancelada." -ForegroundColor Yellow
            return
        }

        Write-Host ""
        Write-Host "  Buscando computadores em: $ouPath" -ForegroundColor Cyan
        Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray

        $computers = Get-ADComputer -Filter * -SearchBase $ouPath -Properties OperatingSystem, LastLogonTimeStamp, Enabled | Sort-Object Name
        
        if (-not $computers) {
            Write-Host "  Nenhum computador encontrado nesta OU." -ForegroundColor Yellow
            return
        }

        Write-Host ""
        Write-Host "  $($computers.Count) computador(es) encontrado(s):" -ForegroundColor White
        Write-Host ""

        $enabled = ($computers | Where-Object { $_.Enabled -eq $true }).Count
        $disabled = ($computers | Where-Object { $_.Enabled -eq $false }).Count

        Write-Host "  Ativos: $enabled | Desativados: $disabled" -ForegroundColor DarkGray
        Write-Host ""

        foreach ($computer in $computers) {
            $status = if ($computer.Enabled) { "✓" } else { "✗" }
            $color = if ($computer.Enabled) { "White" } else { "DarkGray" }
            $lastLogon = if ($computer.LastLogonTimeStamp) { 
                [DateTime]::FromFileTime($computer.LastLogonTimeStamp).ToString('dd/MM/yyyy') 
            } else { 
                "-" 
            }
            $os = if ($computer.OperatingSystem) { $computer.OperatingSystem.Split(' ')[0] } else { "-" }
            
            Write-Host "  $status $($computer.Name.PadRight(20)) $os".PadRight(40) -ForegroundColor $color -NoNewline
            Write-Host "Último: $lastLogon" -ForegroundColor DarkGray
        }
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-ComputerStatus {
    <#
    .SYNOPSIS
        Exibe status detalhado de um computador específico
    .DESCRIPTION
        Mostra informações completas: status de conta, sistema operacional,
        último logon, grupos, GPOs aplicadas, e status de conexão.
    #>
    
    Show-Header
    Write-Host "  STATUS DO COMPUTADOR" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $computerName = Read-Host "  Nome do computador"
    
    if ([string]::IsNullOrWhiteSpace($computerName)) {
        Write-Host "  Operação cancelada." -ForegroundColor Yellow
        return
    }

    try {
        # Buscar computador no AD
        $computer = Get-ADComputer -Identity $computerName -Properties * -ErrorAction Stop
        
        Write-Host ""
        Write-Host "  ┌─────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │ INFORMAÇÕES BÁSICAS" -ForegroundColor Cyan
        Write-Host "  ├─────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │ Nome:         $($computer.Name)" -ForegroundColor White
        Write-Host "  │ DNS Hostname: $($computer.DNSHostName)" -ForegroundColor White
        Write-Host "  │ Status:       $(if ($computer.Enabled) { '✓ Ativo' } else { '✗ Desativado' })" -ForegroundColor $(if ($computer.Enabled) { 'Green' } else { 'Red' })
        Write-Host "  │ Criado em:    $($computer.WhenCreated.ToString('dd/MM/yyyy HH:mm'))" -ForegroundColor White
        Write-Host "  └─────────────────────────────────────────────────" -ForegroundColor DarkGray
        
        # Sistema operacional
        Write-Host ""
        Write-Host "  ┌─────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │ SISTEMA OPERACIONAL" -ForegroundColor Cyan
        Write-Host "  ├─────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │ OS:           $($computer.OperatingSystem)" -ForegroundColor White
        Write-Host "  │ Versão:       $($computer.OperatingSystemVersion)" -ForegroundColor White
        Write-Host "  │ Service Pack: $($computer.OperatingSystemServicePack)" -ForegroundColor White
        Write-Host "  └─────────────────────────────────────────────────" -ForegroundColor DarkGray
        
        # Atividade
        Write-Host ""
        Write-Host "  ┌─────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │ ATIVIDADE" -ForegroundColor Cyan
        Write-Host "  ├─────────────────────────────────────────────────" -ForegroundColor DarkGray
        
        $lastLogon = if ($computer.LastLogonTimeStamp) { 
            [DateTime]::FromFileTime($computer.LastLogonTimeStamp).ToString('dd/MM/yyyy HH:mm')
            $daysSinceLogon = [Math]::Floor((Get-Date) - [DateTime]::FromFileTime($computer.LastLogonTimeStamp)).TotalDays
        } else { 
            "Nunca"
            $daysSinceLogon = "-"
        }
        
        $lastBadPwd = if ($computer.LastBadPasswordAttempt) { 
            [DateTime]::FromFileTime($computer.LastBadPasswordAttempt).ToString('dd/MM/yyyy HH:mm')
        } else { 
            "-" 
        }
        
        Write-Host "  │ Último logon: $lastLogon" -ForegroundColor White
        if ($daysSinceLogon -ne "-") {
            $logonColor = if ($daysSinceLogon -gt 90) { "Red" } elseif ($daysSinceLogon -gt 30) { "Yellow" } else { "Green" }
            Write-Host "  │ Dias desde logon: $daysSinceLogon" -ForegroundColor $logonColor
        }
        Write-Host "  │ Última senha errada: $lastBadPwd" -ForegroundColor White
        Write-Host "  │ Tentativas erradas: $($computer.BadPwdCount)" -ForegroundColor White
        Write-Host "  └─────────────────────────────────────────────────" -ForegroundColor DarkGray
        
        # Localização
        Write-Host ""
        Write-Host "  ┌─────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │ LOCALIZAÇÃO" -ForegroundColor Cyan
        Write-Host "  ├─────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │ OU: $(($computer.DistinguishedName -split ',', 2)[1])" -ForegroundColor White
        Write-Host "  └─────────────────────────────────────────────────" -ForegroundColor DarkGray
        
        # Verificar se está online
        Write-Host ""
        Write-Host "  Testando conexão..." -ForegroundColor DarkGray
        
        if (Test-Connection -ComputerName $computer.DNSHostName -Count 1 -Quiet -ErrorAction SilentlyContinue) {
            Write-Host "  ✓ Computador ONLINE" -ForegroundColor Green
            
            # Tentar buscar mais informações via WinRM
            try {
                $osInfo = Invoke-Command -ComputerName $computer.DNSHostName -ScriptBlock {
                    Get-CimInstance Win32_OperatingSystem
                } -ErrorAction Stop -TimeoutSeconds 10
                
                $uptime = (Get-Date) - $osInfo.LastBootUpTime
                Write-Host "  Uptime: $($uptime.Days) dias, $($uptime.Hours) horas" -ForegroundColor White
                Write-Host "  RAM livre: $([Math]::Round($osInfo.FreePhysicalMemory / 1MB, 1)) GB" -ForegroundColor White
            }
            catch {
                Write-Host "  (Não foi possível obter detalhes via WinRM)" -ForegroundColor DarkGray
            }
        }
        else {
            Write-Host "  ✗ Computador OFFLINE ou inacessível" -ForegroundColor Red
        }
        
        # Grupos
        Write-Host ""
        Write-Host "  ┌─────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │ GRUPOS" -ForegroundColor Cyan
        Write-Host "  ├─────────────────────────────────────────────────" -ForegroundColor DarkGray
        
        $groups = Get-ADPrincipalGroupMembership -Identity $computer -ErrorAction SilentlyContinue
        if ($groups) {
            foreach ($group in $groups | Sort-Object Name) {
                Write-Host "  │ • $($group.Name)" -ForegroundColor White
            }
        }
        else {
            Write-Host "  │ (Nenhum grupo além do padrão)" -ForegroundColor DarkGray
        }
        Write-Host "  └─────────────────────────────────────────────────" -ForegroundColor DarkGray
        
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "  Computador não encontrado no AD." -ForegroundColor Red
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}
