# AD Toolkit

Ferramenta em PowerShell para gestão do Active Directory com interface de linha de comando estilo CMD.

## Funcionalidades

### Usuários
- 🔒 Listar usuários bloqueados (com causa e origem do bloqueio)
- 👥 Listar usuários por grupo
- 🔍 Buscar usuários por filtro (nome, OU, etc.)
- 📊 Exportar relatórios em CSV

### Computadores
- 💻 Listar computadores inativos (90+ dias sem login)
- 🔄 Forçar GPUpdate remoto

### Grupos
- 📋 Membros de um grupo específico
- ➕ Adicionar/remover usuários de grupos

## Requisitos

- Windows Server 2008 R2+ ou Windows 10/11
- PowerShell 5.1+
- Módulo Active Directory (`Install-WindowsFeature RSAT-AD-PowerShell`)
- Privilégios de Domain Admin ou delegação adequada

## Uso

```powershell
# Executar como Administrador
.\ad-toolkit.ps1
```

## Estrutura

```
ad-toolkit/
├── ad-toolkit.ps1      # Menu principal
├── modules/
│   ├── users.ps1       # Funções de usuários
│   ├── groups.ps1      # Funções de grupos
│   ├── computers.ps1   # Funções de computadores
│   └── reports.ps1     # Exportação de relatórios
└── README.md
```

## Investigação de Bloqueios

A função de usuários bloqueados mostra:
- **Causa**: Senha expirada, senha errada, conta expirada, etc.
- **Origem**: De qual computador/servidor veio o bloqueio
- **Timestamp**: Quando aconteceu
- **Tentativas**: Quantas tentativas foram feitas

## Licença

MIT
