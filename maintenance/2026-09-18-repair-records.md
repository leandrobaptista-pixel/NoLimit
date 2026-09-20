# Recuperação dos cadastros — 18/09/2026

Ambiente principal: https://nolimitcontractor.net/no-limit-admin-beta/

Correções aplicadas no serviço publicado:
- RPC save_admin_workspace corrigida: coluna updated_at qualificada, preservando RLS e controle de concorrência.
- Visit Requests agora autentica o usuário, confere vínculo ativo de atendimento e sincroniza os registros UUID reais do website.
- Permissões do serviço interno corrigidas sem criar leitura anônima da tabela privada.
- URL e credencial da conexão de entrada corrigidas no gerenciador de segredos. Nenhuma credencial consta neste documento.
- Inicialização aguarda autenticação; erros de conexão aparecem no painel.
- Status e notas da equipe não são substituídos pela sincronização.
- Categoria antiga Ceiling compatibilizada com o filtro atual.
- Cabeçalho e resumo de categorias compactados para priorizar a lista de pedidos.

Validação realizada no site publicado:
- Recuperados os 16 pedidos existentes.
- Novo pedido TESTE — Pedido do site 18-09 — não é cliente enviado pelo formulário público; confirmação exibida e 17 pedidos no painel.
- Revisão do pedido de teste gravada e mantida após atualizar.
- Fornecedor, subcontratado e cliente identificados como TESTE criados pelo painel, com confirmação de sincronização.
- Permissões da conta administrativa da Geanniny verificadas no banco para leitura e gravação do workspace; conflito de versões também testado em transação revertida. Não equivale a um teste de login no telefone dela.

WhatsApp/SMS permanecem adiados por decisão do proprietário. Nenhum serviço pago foi ativado. Envio automático de email de novo pedido ainda não foi comprovado: o endpoint atual salva o pedido e o painel oferece contato manual por email.
