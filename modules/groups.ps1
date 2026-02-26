<#
.SYNOPSIS
    Módulo de funções de grupos do AD
#>

function Get-GroupMembers {
    <#
    .SYNOPSIS
        Lista membros de um grupo
    #>
    
    Show-Header
    Write-Host "  MEMBROS DO GRUPO" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $groupName = Read-Host "  Digite o nome do grupo"
    
    if ([string]::IsNullOrWhiteSpace($groupName)) {
        Write-Host "  Operação cancelada." -ForegroundColor Yellow
        return
    }

    try {
        $groups = Get-ADGroup -Filter "Name -like '*$groupName*'"
        
        if (-not $groups) {
            Write-Host "  Grupo não encontrado." -ForegroundColor Red
            return
        }

        if ($groups.Count -gt 1) {
            Write-Host "  Múltiplos grupos encontrados:" -ForegroundColor Cyan
            $i = 1
            foreach ($g in $groups) {
                Write-Host "  [$i] $($g.Name)" -ForegroundColor White
                $i++
            }
            Write-Host ""
            $selection = Read-Host "  Selecione"
            $group = $groups[$selection - 1]
        }
        else {
            $group = $groups
        }

        Write-Host ""
        Write-Host "  Grupo: $($group.Name)" -ForegroundColor Cyan
        Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray

        $members = Get-ADGroupMember -Identity $group -Recursive -ErrorAction Stop
        
        if (-not $members) {
            Write-Host "  Grupo vazio." -ForegroundColor Yellow
            return
        }

        # Agrupar por tipo
        $users = $members | Where-Object { $_.objectClass -eq "user" }
        $nestedGroups = $members | Where-Object { $_.objectClass -eq "group" }
        $computers = $members | Where-Object { $_.objectClass -eq "computer" }

        Write-Host ""
        Write-Host "  👤 Usuários: $($users.Count)" -ForegroundColor White
        Write-Host "  📁 Grupos aninhados: $($nestedGroups.Count)" -ForegroundColor White
        Write-Host "  💻 Computadores: $($computers.Count)" -ForegroundColor White
        Write-Host ""

        if ($users) {
            Write-Host "  Lista de usuários:" -ForegroundColor DarkGray
            $users | Sort-Object Name | ForEach-Object {
                Write-Host "    • $($_.Name)" -ForegroundColor White
            }
        }
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Add-UserToGroupInteractive {
    <#
    .SYNOPSIS
        Adiciona usuário a um grupo
    #>
    
    Show-Header
    Write-Host "  ADICIONAR USUÁRIO AO GRUPO" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $username = Read-Host "  Login do usuário"
    if ([string]::IsNullOrWhiteSpace($username)) { return }

    $groupName = Read-Host "  Nome do grupo"
    if ([string]::IsNullOrWhiteSpace($groupName)) { return }

    try {
        $user = Get-ADUser -Identity $username -ErrorAction Stop
        $group = Get-ADGroup -Identity $groupName -ErrorAction Stop

        # Verificar se já é membro
        $memberOf = Get-ADUser -Identity $user | Get-ADPrincipalGroupMembership -ErrorAction SilentlyContinue
        if ($memberOf | Where-Object { $_.DistinguishedName -eq $group.DistinguishedName }) {
            Write-Host "  Usuário já é membro deste grupo." -ForegroundColor Yellow
            return
        }

        Add-ADGroupMember -Identity $group -Members $user
        Write-Host "  ✓ $($user.Name) adicionado ao grupo $($group.Name)!" -ForegroundColor Green
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Remove-UserFromGroupInteractive {
    <#
    .SYNOPSIS
        Remove usuário de um grupo
    #>
    
    Show-Header
    Write-Host "  REMOVER USUÁRIO DO GRUPO" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    $username = Read-Host "  Login do usuário"
    if ([string]::IsNullOrWhiteSpace($username)) { return }

    $groupName = Read-Host "  Nome do grupo"
    if ([string]::IsNullOrWhiteSpace($groupName)) { return }

    try {
        $user = Get-ADUser -Identity $username -ErrorAction Stop
        $group = Get-ADGroup -Identity $groupName -ErrorAction Stop

        Remove-ADGroupMember -Identity $group -Members $user -Confirm:$true
        Write-Host "  ✓ $($user.Name) removido do grupo $($group.Name)!" -ForegroundColor Green
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-EmptyGroups {
    <#
    .SYNOPSIS
        Lista grupos vazios
    #>
    
    Show-Header
    Write-Host "  GRUPOS VAZIOS" -ForegroundColor Yellow
    Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    try {
        $allGroups = Get-ADGroup -Filter * -Properties Members | Select-Object -First 200
        $emptyGroups = $allGroups | Where-Object { $_.Members.Count -eq 0 }

        if (-not $emptyGroups) {
            Write-Host "  Nenhum grupo vazio encontrado." -ForegroundColor Green
            return
        }

        Write-Host "  $($emptyGroups.Count) grupo(s) vazio(s):" -ForegroundColor Yellow
        Write-Host ""

        foreach ($group in $emptyGroups | Sort-Object Name) {
            Write-Host "  • $($group.Name)" -ForegroundColor White
        }

        Write-Host ""
        Write-Host "  Dica: Grupos vazios podem ser candidatos a exclusão." -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}
