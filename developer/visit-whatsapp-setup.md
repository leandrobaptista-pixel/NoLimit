# No Limit — avisos de novos pedidos pelo WhatsApp

Atualizado: 18/09/2026. Status: ESTUDO ADIADO por decisão do proprietário; não implantar nem ativar.
O site publicado é a versão principal. Cópias locais são apoio/backup.
Twilio não é utilizado por esta integração.

## Decisão do proprietário — 18/09/2026

A integração foi adiada para quando o negócio estiver movimentando melhor.
Não concluir a criação do aplicativo Meta, aceitar termos, configurar credenciais,
ativar testes reais, cadastrar pagamento ou habilitar mensagens sem nova solicitação.
A autorização anterior para implantação deixou de valer com esta decisão.
O formulário publicado foi desconectado do adaptador WhatsApp: salvar um pedido
não chama a Meta, mesmo que variáveis de ativação sejam configuradas por engano.
Código e documentação permanecem somente como estudo para uma futura retomada.
As etapas abaixo são referência futura, não trabalho autorizado agora.

## Comportamento estudado

Após salvar um cadastro com sucesso, avisar a equipe de atendimento por WhatsApp:
“No Limit: novo pedido recebido no site. Abra para atender: {{1}}”.
O parâmetro contém o link daquele pedido, sem dados pessoais do cliente.
O link exige a autenticação e as permissões já existentes do painel.
Consultar um pedido não muda seu status automaticamente.
Uma falha de WhatsApp não deve transformar um cadastro salvo em erro no formulário.

## Preparado nesta alteração

- Adaptador servidor → Meta Cloud API, sem intermediário Twilio.
- Destinatários definidos apenas no servidor; nunca pelo formulário público.
- Envio desativado por padrão e condicionado à validação do link do pedido.
- Modo de teste separado do modo de produção.
- Produção exige autorização explícita de custos na configuração.
- URL `https://nolimitcontractor.net/no-limit-admin-beta/?visit=ID#requests`.
- Painel abre detalhes de um pedido autorizado por `id` ou `source_record_id`,
  mesmo se concluído/arquivado; mostra aviso quando não consegue localizar o pedido.
- Testes automatizados simulam a Meta; não enviam mensagens reais.

## Bloqueios antes de ativar

1. Concluir a criação do app No Limit Avisos com o caso de uso WhatsApp.
   A tela final da Meta exige aceitar Termos da Plataforma e Políticas do Desenvolvedor.
   O aplicativo antigo encontrado estava desativado por inatividade.
2. Obter ambiente/número de teste da Meta e verificar os dois destinatários.
   Confirmar no painel que o remetente é de TESTE e não possui cobrança habilitada.
   Não cadastrar pagamento para esta fase.
3. Corrigir e publicar a sincronização da entrada de pedidos:
   `functions/api/visit-request.js` grava `app_records`, kind `visitRequest`, ID UUID;
   a cópia atual de `manage-visit-requests` consulta `public_visit_requests`, ID numérico.
   É necessário preservar o UUID como `source_record_id` no retorno autorizado ou
   usar o mesmo UUID como `id`, sem perder os registros/avaliações existentes.
   Também validar as permissões com contas separadas da equipe.
4. Comprovar que o pedido criado no formulário abre pelo link após login.
5. Somente então habilitar o teste WhatsApp, confirmar recebimento nos celulares
   e clicar no link em cada aparelho. Aceitação HTTP da Meta não comprova entrega.

## Configuração no Cloudflare Pages (servidor)

Projeto: `nolimitcontractor`. Nenhuma credencial vai ao JavaScript do navegador.
- `VISIT_WHATSAPP_ENABLED=false` (padrão; mudar somente na ativação autorizada).
- `VISIT_REQUEST_LINK_VERIFIED=false` (mudar após teste do pedido real).
- `META_WHATSAPP_MODE=test`.
- `META_TEST_SENDER_CONFIRMED=true` somente após confirmar o número de teste na Meta.
- `META_PHONE_NUMBER_ID`: identificador do remetente de teste, não o telefone da equipe.
- `META_GRAPH_VERSION`: versão suportada exibida pelo painel da Meta, por exemplo `vNN.0`.
- `META_WHATSAPP_TOKEN`: segredo do servidor, nunca publicado, nunca em logs.
- `META_WHATSAPP_RECIPIENTS`: um ou dois números autorizados, formato internacional
  somente dígitos, separados por vírgula. Configurar os números do proprietário
  e da administradora informados pelo usuário, sem publicá-los neste documento.

No modo de teste, o adaptador usa texto livre com link. Cada destinatário deve
abrir uma conversa com o número de teste; texto livre depende da janela permitida
pela Meta. Se o painel só permitir `hello_world`, usar esse modelo apenas para
validar a conexão: ele NÃO comprova o fluxo com link personalizado.
Tokens temporários podem expirar. Não considerar o ambiente de teste uma solução
permanente e gratuita de produção.

## Produção futura, somente após avaliar custos

- Modelo aprovado `novo_pedido_site`, idioma `pt_BR`, corpo acima com um parâmetro.
- `META_VISIT_TEMPLATE=novo_pedido_site`, `META_TEMPLATE_LANGUAGE=pt_BR`.
- `META_WHATSAPP_MODE=production` e `META_WHATSAPP_COST_APPROVED=true` apenas após
  aprovação do proprietário. Remetente e credenciais de produção devem ser revisados.
- Antes de ativar, adicionar proteção contra abuso do formulário, controle de volume,
  registro persistente do envio por pedido/destinatário e deduplicação.
- Esta preparação NÃO contém fila persistente, retentativa automática nem webhook
  de confirmação de entrega. Logs indicam somente aceitação/falha do provedor.
- Em caso de problema: `VISIT_WHATSAPP_ENABLED=false` e republicar.

## Verificação

`node --test tests/visit-whatsapp.test.mjs`

Critérios de aceite operacional ainda pendentes: cadastro salvo, visível nas duas
contas, recebimento nos dois celulares, link abre o pedido certo após login,
usuário sem permissão não acessa o pedido, custo do teste confirmado na Meta.

## Referências

- https://www.postman.com/meta/whatsapp-business-platform/documentation/wlk6lh4/whatsapp-cloud-api
- https://developers.facebook.com/docs/whatsapp/cloud-api/get-started
- https://business.whatsapp.com/products/platform-pricing
