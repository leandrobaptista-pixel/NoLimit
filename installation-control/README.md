# Cabinets Control

Aplicativo web (PWA) para controle de entrega e instalacao de cabinets, vanities, med cabinets e countertops, com uso em obra mesmo sem internet.

## Fluxo de abertura

1. Criar/login de usuario para identificar quem opera o sistema.
2. Cadastrar `Clientes`.
3. Cadastrar `Projetos` vinculados ao cliente (um cliente pode ter varios projetos).
4. Cadastrar `Pessoas` por projeto (Sub Contractor, Foreman, Project Manager, Cliente etc).
5. Cadastrar `Containers` do fabricante com ETA/status e quantidades.
6. Cadastrar unidades selecionando o projeto ja existente.
7. Controlar envio/retencao do warehouse usando saldo do container para evitar erro/excesso.

## O que esta pronto

- Controle por unidade desde `warehouse` ate `instalacao`
- Baixa por etapa com data/hora e trilha por perfil
- Login local com perfis:
  - `admin`
  - `warehouse`
  - `transport`
  - `qa`
  - `installer`
- Cadastro base de cliente/projeto/pessoas
- Modulo de containers:
  - lista de materiais (manifesto do fabricante)
  - schedule de chegada (ETA/status)
  - quantidade de kitchens/vanities/med cabinets/countertops por container
- controle de envio/retencao no warehouse com validacao de saldo
  - vinculacao do envio com unidade destino
- Controle de qualidade da entrega e status da instalacao
- Campo de pendencias (faltante, defeito, ajuste de shopdrawing, condicoes de parede)
- `Scope work` e referencia de shopdrawing por unidade
- Checklist detalhado por item
- Anexo de fotos por unidade
- Relatorio PDF por unidade
- Relatorio PDF consolidado por projeto (inclui schedule/saldo de containers)
- Sincronizacao em nuvem (Supabase) para multi-dispositivo
- Backup manual (export/import JSON)
- Modo offline com Service Worker + IndexedDB

## Como usar

1. Rode um servidor local no diretorio do projeto:

```bash
python3 -m http.server 8080
```

2. Acesse `http://localhost:8080`
3. No primeiro acesso, crie o usuario administrador.
4. Configure clientes, projetos, pessoas e containers.
5. Configure nuvem no painel `Sincronizacao na nuvem`.

## Sync em nuvem

Ver `SUPABASE_SETUP.md`.

## Observacoes

- Dados continuam operando offline no aparelho.
- Para compliance/seguranca forte, use JWT por usuario e RLS por tenant no Supabase.
