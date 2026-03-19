# Resumo de Alteracoes Recentes

Atualizado em: 2026-03-19
Objetivo: manter um registro rapido das areas alteradas nos ultimos dias, para facilitar acompanhamento sem precisar revisar todo o historico do Git.

## 2026-03-19
### Website principal (`index.html`, `assets/*`)
- A navegacao da galeria foi melhorada.
- A galeria passou a exibir ate 12 imagens por categoria.
- O viewer de fotos agora permite navegar entre imagens e voltar para a secao `Gallery` na pagina principal.
- Referencias de catalogo e galeria foram ajustadas para evitar links quebrados ou imagens faltando.
- O email principal em `Request a Free On-Site Visit` foi alterado para `leandro@nolimitcontractor.net`.
- O formulario publico foi redesenhado para ficar mais profissional, com melhor espacamento, hierarquia visual, validacao e alinhamento de botoes.
- As acoes e estados visuais do formulario publico foram padronizados para ficar mais consistentes.
- Foi preparada a integracao do formulario publico com Supabase, com fallback quando o armazenamento direto nao estiver disponivel.

### App Installation Control (`installation-control/*`)
- O fluxo de clientes foi melhorado para que o clique em um cliente possa abrir a area de detalhes e mostrar os projetos relacionados abaixo.
- As antigas abas fixas de quantidades do container foram substituidas por uma estrutura flexivel de manifesto de materiais.
- O manifesto do container passou a aceitar categorias abertas, como kitchens, vanities, med cabinets, countertops, wood floor, tile, curtain, extra materials e outros itens customizados.
- Foi criado um padrao de identificacao por QR para evitar mistura entre codigo de produto e identificacao da peca fisica.
- Os itens do manifesto do fabricante passaram a poder ser editados apos criacao.
- Foram adicionados campos extras no manifesto: `ADA/Normal`, `Color`, `Format`, `Size` e `Mfr/Dist code`.
- Os metadados dos QR codes passaram a seguir esses campos, e a alteracao de quantidade passou a reconciliar as pecas QR vinculadas.
- Foram adicionados filtros no manifesto por `ADA`, `color`, `format`, `size` e `Mfr/Dist code`.
- Foi adicionada impressao de QR por linha.
- Foi adicionada impressao em lote para linhas filtradas.
- Foi adicionada exportacao CSV e PDF das linhas filtradas do manifesto.
- Foi adicionada selecao em massa por checkbox para que print/export possam usar linhas selecionadas, e nao apenas o filtro.
- `QR Count` passou a ser clicavel e abrir a lista detalhada das pecas QR daquela linha do manifesto.

### Suporte / configuracao
- `assets/contact-config.js` foi criado para configuracao do formulario publico.
- `contact-form-supabase.sql` foi criado para a tabela do formulario publico no Supabase.

### Observacao de deploy
- Commit publicado no GitHub: `129a87a`.
- O deploy automatico no Cloudflare Pages foi disparado, mas falhou no GitHub Actions na etapa de deploy.
- Workflow para revisar: `.github/workflows/deploy-cloudflare-pages.yml`.

## 2026-03-13
### App Installation Control
- Foram adicionadas ferramentas locais de backup para developer.
- A cobertura de auditoria e sincronizacao foi reforcada.
- Foram adicionados arquivos de apoio para perfis/avatares e suporte da aplicacao.
- Principais arquivos afetados nesta atualizacao:
  - `installation-control/app.js`
  - `installation-control/index.html`
  - `installation-control/styles.css`
  - `installation-control/cloud-sync-config.js`
  - `installation-control/sw.js`
  - avatares e arquivos de apoio

## 2026-03-12
### Fluxo de clientes
- Foi criado um hub de transicao antes dos formularios de cliente.
- O fluxo de navegacao do workspace de clientes foi reorganizado para melhorar o acesso antes da entrada nos formularios detalhados.

### Seguranca / recuperacao
- Foram adicionadas confirmacoes de exclusao.
- Foi adicionada uma lixeira com janela de recuperacao de 48 horas para reduzir exclusoes permanentes acidentais.

## 2026-03-08
### Estrutura de Clientes / Projetos
- O fluxo entre clientes e projetos foi reestruturado.
- Foram adicionados templates de escopo e checklist de projeto.
- As secoes Projects e People passaram a ficar somente dentro do menu Projects para reduzir duplicidade.
- O painel de clientes foi ampliado e passou a ter links rapidos para projetos.
- A lista de clientes foi simplificada e o layout do painel foi ajustado para ficar mais enxuto.

## Principais areas alteradas recentemente
- Website publico e experiencia da galeria
- Formulario publico de contato / leads
- Navegacao de clientes no Installation Control
- Manufacturer material manifest
- Estrutura de containers e rastreio por QR no warehouse
- Controles de exportacao e impressao do manifesto
- Ferramentas de backup, auditoria e sincronizacao
- Confirmacoes de exclusao e lixeira com recuperacao

## Regra sugerida de manutencao
- A cada alteracao relevante, adicionar um item curto aqui com:
  - data
  - area alterada
  - o que foi feito
  - arquivos principais
  - observacao de deploy, se necessario
