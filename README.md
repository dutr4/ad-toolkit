# AD Toolkit

Ferramenta em PowerShell para gestão do Active Directory com interface de linha de comando interativa.

```
█████╗ ██████╗     ████████╗ ██████╗  ██████╗ ██╗     ██╗  ██╗██╗████████╗
██╔══██╗██╔══██╗    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██║ ██╔╝██║╚══██╔══╝
███████║██║  ██║       ██║   ██║   ██║██║   ██║██║     █████╔╝ ██║   ██║   
██╔══██║██║  ██║       ██║   ██║   ██║██║   ██║██║     ██╔═██╗ ██║   ██║   
██║  ██║██████╔╝       ██║   ╚██████╔╝╚██████╔╝███████╗██║  ██╗██║   ██║   
╚═╝  ╚═╝╚═════╝        ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝   ╚═╝   
```

## ✨ Funcionalidades

### 👤 Usuários (6 funções)
| Função | Descrição |
|--------|-----------|
| `Get-LockedUsersDetailed` | Lista bloqueados com **causa específica**, origem (IP/máquina), tipo de logon |
| `Unlock-ADUserInteractive` | Desbloqueia usuários com seleção interativa |
| `Get-UsersByGroup` | Lista todos os membros de um grupo |
| `Search-ADUsers` | Busca por nome, login, OU ou e-mail |
| `Get-UsersPasswordExpiring` | Senhas prestes a expirar |
| `Get-InactiveUsers` | Usuários sem login há X dias |

### 📁 Grupos (4 funções)
| Função | Descrição |
|--------|-----------|
| `Get-GroupMembers` | Lista membros recursivamente (usuários, grupos aninhados, computadores) |
| `Add-UserToGroupInteractive` | Adiciona usuário a um grupo |
| `Remove-UserFromGroupInteractive` | Remove usuário de um grupo |
| `Get-EmptyGroups` | Lista grupos sem membros |

### 💻 Computadores (4 funções)
| Função | Descrição |
|--------|-----------|
| `Get-InactiveComputers` | Computadores sem login há X dias (padrão: 90) |
| `Invoke-GPUpdateRemote` | GPUpdate remoto em um, vários ou toda uma OU |
| `Get-ComputersByOU` | Lista computadores por Unidade Organizacional |
| `Get-ComputerStatus` | Status detalhado: OS, uptime, grupos, online/offline |

### 📊 Relatórios (4 funções)
| Função | Descrição |
|--------|-----------|
| `Export-LockedUsersCSV` | Exporta usuários bloqueados para CSV |
| `Export-UsersByGroupCSV` | Exporta membros de um grupo para CSV |
| `Export-InactiveComputersCSV` | Exporta computadores inativos para CSV |
| `Export-DomainReportHTML` | Relatório completo em HTML com dashboard e estatísticas |

## 🔍 Investigação de Bloqueios

A função `Get-LockedUsersDetailed` fornece diagnóstico completo:

**Causas identificadas:**
- `0xC000006A` — Senha incorreta
- `0xC0000071` — Senha expirada
- `0xC0000072` / `0xC0000193` — Conta expirada
- `0xC0000234` — Conta bloqueada por tentativas incorretas
- `0xC000006F` — Logon fora do horário permitido
- `0xC0000070` — Logon de workstation não autorizada
- E mais...

**Informações exibidas:**
- 🕐 Timestamp do bloqueio
- 🖥️ Computador/IP de origem
- 📋 Tipo de logon (interativo, RDP, rede, serviço, etc.)
- 📈 Número de tentativas recentes
- ⚠️ Status da senha e conta

## 📋 Requisitos

- **Windows Server 2008 R2+** ou **Windows 10/11**
- **PowerShell 5.1+**
- **Módulo Active Directory**
  ```powershell
  Install-WindowsFeature RSAT-AD-PowerShell
  ```
- Privilégios de **Domain Admin** ou delegação adequada

## 🚀 Uso

```powershell
# Clonar o repositório
git clone https://github.com/dutr4/ad-toolkit.git

# Entrar na pasta
cd ad-toolkit

# Executar como Administrador
.\ad-toolkit.ps1
```

## 📁 Estrutura

```
ad-toolkit/
├── ad-toolkit.ps1      # Menu principal + navegação
├── modules/
│   ├── users.ps1       # 6 funções de usuários
│   ├── groups.ps1      # 4 funções de grupos
│   ├── computers.ps1   # 4 funções de computadores
│   └── reports.ps1     # 4 funções de exportação
├── reports/            # Pasta de saída (criada automaticamente)
└── README.md
```

## 📸 Screenshots

### Menu Principal
```
  MENU PRINCIPAL

  [1] Usuários
  [2] Grupos
  [3] Computadores
  [4] Relatórios
  [5] Configurações

  [Q] Sair
```

### Usuários Bloqueados (Detalhado)
```
  ┌─────────────────────────────────────────────────
  │ Nome:       João Silva
  │ Login:      jsilva
  │ Bloqueado em: 26/02/2026 10:45:32
  │ Origem:     WORKSTATION-05
  │ Causa:      Senha incorreta
  │ Tipo logon: RemoteInteractive (RDP)
  │ Tentativas recentes visíveis: 5
  └─────────────────────────────────────────────────
```

### Relatório HTML
Dashboard com:
- 👥 Estatísticas de usuários (ativos, desativados, bloqueados)
- 💻 Estatísticas de computadores e sistemas operacionais
- 📁 Grupos e OUs
- 🔐 Política de senhas do domínio

## 📝 Licença

MIT

## 👤 Autor

Guilherme Dutra Campos
