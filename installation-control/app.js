const DB_NAME = "cabinets-control-db";
const DB_VERSION = 5;
const UNIT_STORE = "units";
const PHOTO_STORE = "photos";
const USER_STORE = "users";
const SETTINGS_STORE = "settings";
const CLIENT_STORE = "clients";
const PROJECT_STORE = "projects";
const CONTACT_STORE = "contacts";
const CONTAINER_STORE = "containers";
const MATERIAL_STORE = "materials";
const SESSION_KEY = "cc-session-user-id";
const APP_AUDIT_KEY = "cc-app-audit-log";
const APP_AUDIT_MAX = 3000;

const STAGES = [
  { key: "warehouse", label: "Warehouse" },
  { key: "transportation", label: "Transportation" },
  { key: "siteDelivery", label: "Site Job" },
  { key: "distribution", label: "Distribuicao" },
  { key: "quality", label: "Qualidade" },
  { key: "installation", label: "Instalacao" },
];

const ROLE_LABEL = {
  developer: "Developer",
  admin: "Admin",
  warehouse: "Warehouse",
  transport: "Transport",
  qa: "QA",
  installer: "Installer",
  visitor: "Visitor",
};

const CONTACT_ROLE_LABEL = {
  subcontractor: "Sub Contractor",
  foreman: "Foreman",
  "project-manager": "Project Manager",
  client: "Cliente",
  other: "Other",
};

const CONTAINER_STATUS_LABEL = {
  scheduled: "Scheduled",
  "in-transit": "In Transit",
  arrived: "Arrived",
  delayed: "Delayed",
};

const STAGE_ROLE_ACCESS = {
  warehouse: ["warehouse"],
  transport: ["transportation", "siteDelivery"],
  qa: ["quality"],
  installer: ["distribution", "installation"],
};

const DEFAULT_SYNC = {
  supabaseUrl: "",
  supabaseAnonKey: "",
  tenant: "",
  autoSync: false,
  lastSyncAt: null,
};

const DEFAULT_LANG = "en";
const SUPPORTED_LANGS = ["en"];
const PRIMARY_DEVELOPER_NAME = "leandro baptista";
const PROJECT_SECTORS = ["fabrica", "warehouse", "delivery", "distribuicao", "instalacao", "punchlist"];
const COI_REMINDER_DAYS = 30;
const COI_MAX_FILE_SIZE = 8 * 1024 * 1024;
const SECTOR_STAGE_MAP = {
  fabrica: [],
  warehouse: ["warehouse"],
  delivery: ["transportation", "siteDelivery"],
  distribuicao: ["distribution"],
  instalacao: ["installation"],
  punchlist: ["quality"],
};

const STAGE_LABELS = {
  en: {
    warehouse: "Warehouse",
    transportation: "Transportation",
    siteDelivery: "Site Job",
    distribution: "Distribution",
    quality: "Quality",
    installation: "Installation",
  },
  pt: {
    warehouse: "Warehouse",
    transportation: "Transportation",
    siteDelivery: "Site Job",
    distribution: "Distribuicao",
    quality: "Qualidade",
    installation: "Instalacao",
  },
  es: {
    warehouse: "Almacen",
    transportation: "Transporte",
    siteDelivery: "Obra",
    distribution: "Distribucion",
    quality: "Calidad",
    installation: "Instalacion",
  },
};

const ROLE_LABELS = {
  en: {
    developer: "Developer",
    admin: "Admin",
    warehouse: "Warehouse",
    transport: "Transport",
    qa: "QA",
    installer: "Installer",
    visitor: "Visitor",
  },
  pt: {
    developer: "Desenvolvedor",
    admin: "Admin",
    warehouse: "Warehouse",
    transport: "Transport",
    qa: "QA",
    installer: "Installer",
    visitor: "Visitante",
  },
  es: {
    developer: "Desarrollador",
    admin: "Admin",
    warehouse: "Almacen",
    transport: "Transporte",
    qa: "Calidad",
    installer: "Instalador",
    visitor: "Visitante",
  },
};

const CONTACT_ROLE_LABELS = {
  en: {
    subcontractor: "Sub Contractor",
    foreman: "Foreman",
    "project-manager": "Project Manager",
    client: "Client",
    other: "Other",
  },
  pt: CONTACT_ROLE_LABEL,
  es: {
    subcontractor: "Subcontratista",
    foreman: "Capataz",
    "project-manager": "Project Manager",
    client: "Cliente",
    other: "Otro",
  },
};

const CONTAINER_STATUS_LABELS = {
  en: {
    scheduled: "Scheduled",
    "in-transit": "In Transit",
    arrived: "Arrived",
    delayed: "Delayed",
  },
  pt: {
    scheduled: "Agendado",
    "in-transit": "Em transito",
    arrived: "Recebido",
    delayed: "Atrasado",
  },
  es: {
    scheduled: "Programado",
    "in-transit": "En transito",
    arrived: "Recibido",
    delayed: "Atrasado",
  },
};

const I18N = {
  en: {
    "Verificando conexao...": "Checking connection...",
    Sair: "Logout",
    "Instalar App": "Install App",
    "Falha ao iniciar o app": "App startup failed",
    "Primeiro acesso: criar administrador": "First access: create administrator",
    "Primeiro acesso: criar desenvolvedor": "First access: create developer",
    "Nome completo": "Full name",
    Usuario: "User",
    Senha: "Password",
    "Criar admin": "Create admin",
    "Criar desenvolvedor": "Create developer",
    "Entrar no sistema": "Sign in",
    Entrar: "Sign in",
    "Cadastro base (clientes, projetos e pessoas)": "Master data (clients, projects and people)",
    "Na abertura do sistema voce monta a estrutura do cliente e seus projetos; depois relaciona as pessoas envolvidas em cada projeto.":
      "At startup, set up client and project structure, then assign people to each project.",
    Clientes: "Clients",
    "Nome do cliente": "Client name",
    "Contato principal": "Main contact",
    Phone: "Phone",
    "Adicionar cliente": "Add client",
    "Projetos por cliente": "Projects by client",
    Client: "Client",
    "Nome do projeto": "Project name",
    "Codigo do projeto": "Project code",
    "Endereco / Job Site": "Address / Job Site",
    "Adicionar projeto": "Add project",
    "Pessoas no projeto": "People in project",
    Name: "Name",
    Funcao: "Role",
    Company: "Company",
    Project: "Project",
    "Adicionar pessoa": "Add person",
    "Containers do fabricante e schedule de chegada": "Manufacturer containers and arrival schedule",
    "Este modulo abastece o warehouse: manifesto de materiais do container, ETA de chegada e controle de envio/retencao por destino para evitar entrega errada ou excesso.":
      "This module feeds warehouse operations: container manifest, ETA, and release/hold control by destination.",
    Fabricante: "Manufacturer",
    "ETA de chegada": "Arrival ETA",
    "Status de chegada": "Arrival status",
    "Cozinhas no container": "Kitchens in container",
    "Vanities no container": "Vanities in container",
    "Med Cabinets no container": "Med Cabinets in container",
    "Countertops no container": "Countertops in container",
    Notas: "Notes",
    "Adicionar container": "Add container",
    "Sincronizacao na nuvem (multi-dispositivo)": "Cloud sync (multi-device)",
    "Modo offline-first: dados locais continuam funcionando sem internet. Quando online, use push/pull para sincronizar dispositivos.":
      "Offline-first mode: local data still works without internet. When online, use push/pull to sync devices.",
    "Tenant/Codigo da empresa": "Tenant/Company code",
    "Auto-sync (quando online)": "Auto-sync (when online)",
    Desligado: "Off",
    Ligado: "On",
    "Salvar configuracao": "Save settings",
    "Enviar para nuvem (Push)": "Send to cloud (Push)",
    "Baixar da nuvem (Pull)": "Download from cloud (Pull)",
    "Sem sincronizacao": "No sync",
    "Usuarios e perfis": "Users and roles",
    Perfil: "Role",
    "Adicionar usuario": "Add user",
    "Novo registro por unidade": "New unit record",
    "Nome do projeto": "Project name",
    "Pavimento": "Floor",
    Categoria: "Category",
    "Modelo da cozinha": "Kitchen model",
    "Shopdrawing (referencia/link)": "Shopdrawing (reference/link)",
    "Scope Work do projeto": "Project scope of work",
    "Salvar unidade": "Save unit",
    "Controle operacional": "Operations control",
    "Buscar por cliente, projeto, unidade, tipo, modelo...": "Search by client, project, unit, type, model...",
    Todas: "All",
    Pendentes: "Pending",
    "Com problema de qualidade": "With quality issue",
    "Instalacao bloqueada": "Installation blocked",
    "Relatorio PDF (Projeto)": "Project PDF report",
    "Exportar backup": "Export backup",
    "Importar backup": "Import backup",
    Excluir: "Delete",
    "Manifesto de materiais do fabricante": "Manufacturer material manifest",
    "Cod. item": "Item code",
    Description: "Description",
    Qtd: "Qty",
    "Unidade (pcs/box)": "Unit (pcs/box)",
    "Adicionar no manifesto": "Add to manifest",
    "Controle warehouse: envio ou retencao por destino": "Warehouse control: release or hold by destination",
    "Enviar ao destino": "Release to destination",
    "Reter no warehouse": "Hold in warehouse",
    "Destino (job site/unidade)": "Destination (job site/unit)",
    Nota: "Note",
    "Registrar movimento": "Register movement",
    "PDF unidade": "Unit PDF",
    "Qualidade da entrega": "Delivery quality",
    Pendente: "Pending",
    Aprovada: "Approved",
    Rejeitada: "Rejected",
    "Status da instalacao": "Installation status",
    "Nao iniciada": "Not started",
    "Em andamento": "In progress",
    Concluida: "Completed",
    Bloqueada: "Blocked",
    "Notas de qualidade/instalacao": "Quality/installation notes",
    "Material faltante / defeituoso / ajuste": "Missing / damaged / adjustment material",
    "Checklist detalhado de saida/recebimento": "Detailed outbound/receiving checklist",
    "Qtd. esperada": "Expected qty",
    "Qtd. conferida": "Checked qty",
    Faltante: "Missing",
    Defeito: "Damaged",
    Ajuste: "Adjustment",
    "Adicionar item": "Add item",
    "Anexar fotos": "Attach photos",
    "Sem ETA informado.": "No ETA informed.",
    "Container previsto para hoje.": "Container expected today.",
    "Sem itens no manifesto.": "No manifest items.",
    Envio: "Release",
    Retencao: "Hold",
    Data: "Date",
    Movimento: "Movement",
    Unit: "Unit",
    Destino: "Destination",
    Operador: "Operator",
    "Sem movimentos registrados.": "No movements recorded.",
    "Nenhum container cadastrado.": "No container registered.",
    "Descricao e quantidade sao obrigatorias.": "Description and quantity are required.",
    "Informe ao menos uma quantidade para registrar movimento.": "Enter at least one quantity to register movement.",
    "Quantidade excede o saldo disponivel no container.": "Quantity exceeds available container balance.",
    "No items registered.": "No items registered.",
    "Nenhum registro encontrado.": "No records found.",
    "Scope: nao informado": "Scope: not informed",
    "Shopdrawing: nao informado": "Shopdrawing: not informed",
    "Remover foto": "Remove photo",
    Atualizado: "Updated",
    "Ultimo sync:": "Last sync:",
    "Para cadastrar unidades, primeiro crie cliente e projeto no bloco de cadastro base.":
      "To register units, first create a client and project in master data.",
    "Total unidades": "Total units",
    "Instalacao concluida": "Installation completed",
    "Qualidade rejeitada": "Quality rejected",
    "Chegam em 7 dias": "Arrive in 7 days",
    "Containers atrasados": "Delayed containers",
    "Project for PDF (all)": "Project for PDF (all)",
    "Nenhum cliente cadastrado.": "No clients registered.",
    "Nenhum projeto cadastrado.": "No projects registered.",
    "Nenhuma pessoa cadastrada.": "No people registered.",
    "Nao e possivel excluir cliente com projetos vinculados.": "Cannot delete a client with linked projects.",
    "Nao e possivel excluir cliente com pessoas vinculadas.": "Cannot delete a client with linked people.",
    "Nao e possivel excluir cliente com containers vinculados.": "Cannot delete a client with linked containers.",
    "Nao e possivel excluir projeto com unidades vinculadas.": "Cannot delete project with linked units.",
    "Nao e possivel excluir projeto com pessoas vinculadas.": "Cannot delete project with linked people.",
    "Nao e possivel excluir projeto com containers vinculados.": "Cannot delete project with linked containers.",
    "Usuario ou senha invalidos.": "Invalid username or password.",
    "Usuario ja existe.": "User already exists.",
    "Cliente ja cadastrado.": "Client already exists.",
    "Cliente e nome do projeto sao obrigatorios.": "Client and project name are required.",
    "Projeto ja cadastrado para este cliente.": "Project already exists for this client.",
    "Container, cliente e projeto sao obrigatorios.": "Container, client and project are required.",
    "Container ID ja cadastrado.": "Container ID already exists.",
    "Selecione um projeto.": "Select a project.",
    "Falha ao importar backup. Verifique o arquivo.": "Failed to import backup. Check the file.",
    "Configurar Supabase URL/Key/Tenant": "Configure Supabase URL/Key/Tenant",
    "Sem internet para sincronizar": "No internet to sync",
    "Enviando dados...": "Sending data...",
    "Push concluido:": "Push completed:",
    "Falha no push. Verifique a configuracao.": "Push failed. Check configuration.",
    "Baixando dados...": "Downloading data...",
    "Pull concluido:": "Pull completed:",
    "Falha no pull. Verifique a configuracao.": "Pull failed. Check configuration.",
    "Configuracao salva.": "Settings saved.",
    Online: "Online",
    Offline: "Offline",
    "Nao foi possivel iniciar o banco local. Feche outras abas deste app e recarregue a pagina.":
      "Could not initialize local database. Close other app tabs and reload.",
    "Acesso total: cadastro base, containers, usuarios, sync, QA, instalacao e relatorios.":
      "Full access: master data, containers, users, sync, QA, installation and reports.",
    "Cria unidades, controla containers, manifesto e envio/retencao do warehouse.":
      "Creates units and controls containers, manifest, and warehouse release/hold.",
    "Atualiza transportation/site job e acompanha schedule de containers.":
      "Updates transportation/site job and tracks container schedule.",
    "Valida qualidade, pendencias e pode anexar fotos.": "Validates quality/issues and can attach photos.",
    "Controla distribuicao, instalacao, pendencias e fotos.":
      "Controls distribution, installation, issues and photos.",
    "Acesso desenvolvedor: controle total, inclusive configuracoes de sistema.":
      "Developer access: full control, including system settings.",
    "Acesso admin: gestao de dados, sem alterar estrutura do sistema.":
      "Admin access: data management without changing system structure.",
    "Acesso visitante: apenas consulta e relatorios. Aprovacao de funcoes por admin/developer.":
      "Visitor access: read-only and reports. Role approval by admin/developer.",
    "Acesso visitante: apenas consulta e relatorios. Aprovacao de funcoes somente por admin.":
      "Visitor access: read-only and reports. Role approval is admin-only.",
    "Nao ha unidades para gerar o relatorio.": "There are no units to generate the report.",
  },
  es: {
    "Verificando conexao...": "Verificando conexion...",
    Sair: "Salir",
    "Instalar App": "Instalar app",
    "Falha ao iniciar o app": "Fallo al iniciar la app",
    "Primeiro acesso: criar administrador": "Primer acceso: crear administrador",
    "Primeiro acesso: criar desenvolvedor": "Primer acceso: crear desarrollador",
    "Nome completo": "Nombre completo",
    Usuario: "Usuario",
    Senha: "Contrasena",
    "Criar admin": "Crear admin",
    "Criar desenvolvedor": "Crear desarrollador",
    "Entrar no sistema": "Entrar al sistema",
    Entrar: "Entrar",
    "Cadastro base (clientes, projetos e pessoas)": "Datos base (clientes, proyectos y personas)",
    Clientes: "Clientes",
    "Nome do cliente": "Nombre del cliente",
    "Contato principal": "Contacto principal",
    Phone: "Telefono",
    "Adicionar cliente": "Agregar cliente",
    "Projetos por cliente": "Proyectos por cliente",
    Client: "Cliente",
    "Nome do projeto": "Nombre del proyecto",
    "Codigo do projeto": "Codigo del proyecto",
    "Endereco / Job Site": "Direccion / Obra",
    "Adicionar projeto": "Agregar proyecto",
    "Pessoas no projeto": "Personas del proyecto",
    Name: "Nombre",
    Funcao: "Rol",
    Company: "Empresa",
    Project: "Proyecto",
    "Adicionar pessoa": "Agregar persona",
    "Containers do fabricante e schedule de chegada": "Contenedores del fabricante y calendario de llegada",
    Fabricante: "Fabricante",
    "ETA de chegada": "ETA de llegada",
    "Status de chegada": "Estado de llegada",
    "Cozinhas no container": "Cocinas en contenedor",
    "Vanities no container": "Vanities en contenedor",
    "Med Cabinets no container": "Med Cabinets en contenedor",
    "Countertops no container": "Countertops en contenedor",
    Notas: "Notas",
    "Adicionar container": "Agregar contenedor",
    "Sincronizacao na nuvem (multi-dispositivo)": "Sincronizacion en la nube (multi-dispositivo)",
    "Salvar configuracao": "Guardar configuracion",
    "Enviar para nuvem (Push)": "Enviar a la nube (Push)",
    "Baixar da nuvem (Pull)": "Bajar de la nube (Pull)",
    "Sem sincronizacao": "Sin sincronizacion",
    "Usuarios e perfis": "Usuarios y perfiles",
    Perfil: "Perfil",
    "Adicionar usuario": "Agregar usuario",
    "Novo registro por unidade": "Nuevo registro por unidad",
    "Pavimento": "Piso",
    Categoria: "Categoria",
    "Modelo da cozinha": "Modelo de cocina",
    "Scope Work do projeto": "Alcance del proyecto",
    "Salvar unidade": "Guardar unidad",
    "Controle operacional": "Control operativo",
    Todas: "Todas",
    Pendentes: "Pendientes",
    "Com problema de qualidade": "Con problema de calidad",
    "Instalacao bloqueada": "Instalacion bloqueada",
    "Relatorio PDF (Projeto)": "Reporte PDF (Proyecto)",
    "Exportar backup": "Exportar backup",
    "Importar backup": "Importar backup",
    Excluir: "Eliminar",
    "Manifesto de materiais do fabricante": "Manifiesto de materiales del fabricante",
    Description: "Descripcion",
    Qtd: "Cant",
    "Adicionar no manifesto": "Agregar al manifiesto",
    "Controle warehouse: envio ou retencao por destino": "Control de almacen: envio o retencion por destino",
    "Enviar ao destino": "Enviar al destino",
    "Reter no warehouse": "Retener en almacen",
    "Registrar movimento": "Registrar movimiento",
    "PDF unidade": "PDF unidad",
    "Qualidade da entrega": "Calidad de entrega",
    Pendente: "Pendiente",
    Aprovada: "Aprobada",
    Rejeitada: "Rechazada",
    "Status da instalacao": "Estado de instalacion",
    "Nao iniciada": "No iniciada",
    "Em andamento": "En progreso",
    Concluida: "Concluida",
    Bloqueada: "Bloqueada",
    "Notas de qualidade/instalacao": "Notas de calidad/instalacion",
    "Material faltante / defeituoso / ajuste": "Material faltante / defectuoso / ajuste",
    "Checklist detalhado de saida/recebimento": "Checklist detallado de salida/recepcion",
    "Qtd. esperada": "Cant. esperada",
    "Qtd. conferida": "Cant. revisada",
    Faltante: "Faltante",
    Defeito: "Defecto",
    Ajuste: "Ajuste",
    "Adicionar item": "Agregar item",
    "Anexar fotos": "Adjuntar fotos",
    Data: "Fecha",
    Movimento: "Movimiento",
    Unit: "Unidad",
    Destino: "Destino",
    Operador: "Operador",
    Atualizado: "Actualizado",
    Online: "Online",
    Offline: "Offline",
    "Nao foi possivel iniciar o banco local. Feche outras abas deste app e recarregue a pagina.":
      "No fue posible iniciar la base local. Cierre otras pestanas y recargue la pagina.",
    "Configurar Supabase URL/Key/Tenant": "Configurar Supabase URL/Key/Tenant",
    "Sem internet para sincronizar": "Sin internet para sincronizar",
    "Enviando dados...": "Enviando datos...",
    "Baixando dados...": "Descargando datos...",
    "Push concluido:": "Push completado:",
    "Pull concluido:": "Pull completado:",
    "Falha no push. Verifique a configuracao.": "Fallo en push. Verifique la configuracion.",
    "Falha no pull. Verifique a configuracao.": "Fallo en pull. Verifique la configuracion.",
    "Configuracao salva.": "Configuracion guardada.",
    "Acesso total: cadastro base, containers, usuarios, sync, QA, instalacao e relatorios.":
      "Acceso total: datos base, contenedores, usuarios, sync, calidad, instalacion y reportes.",
    "Cria unidades, controla containers, manifesto e envio/retencao do warehouse.":
      "Crea unidades y controla contenedores, manifiesto y envio/retencion de almacen.",
    "Atualiza transportation/site job e acompanha schedule de containers.":
      "Actualiza transporte/obra y acompana el calendario de contenedores.",
    "Valida qualidade, pendencias e pode anexar fotos.": "Valida calidad, pendientes y puede adjuntar fotos.",
    "Controla distribuicao, instalacao, pendencias e fotos.":
      "Controla distribucion, instalacion, pendientes y fotos.",
    "Acesso desenvolvedor: controle total, inclusive configuracoes de sistema.":
      "Acceso desarrollador: control total, incluidas configuraciones del sistema.",
    "Acesso admin: gestao de dados, sem alterar estrutura do sistema.":
      "Acceso admin: gestion de datos sin alterar la estructura del sistema.",
    "Acesso visitante: apenas consulta e relatorios. Aprovacao de funcoes por admin/developer.":
      "Acceso visitante: solo consulta y reportes. Aprobacion de roles por admin/desarrollador.",
    "Acesso visitante: apenas consulta e relatorios. Aprovacao de funcoes somente por admin.":
      "Acceso visitante: solo consulta y reportes. La aprobacion de roles es solo de admin.",
    "Nao ha unidades para gerar o relatorio.": "No hay unidades para generar el reporte.",
  },
};

let currentLang = DEFAULT_LANG;
const textNodeSource = new WeakMap();
const attrSource = new WeakMap();
const PT_EN_REGEX_REPLACEMENTS = [
  [/\\bUsuario\\b/g, "User"],
  [/\\busuarios\\b/gi, "users"],
  [/\\bSenha\\b/g, "Password"],
  [/\\bPerfil\\b/g, "Role"],
  [/\\bCliente\\b/g, "Client"],
  [/\\bclientes\\b/gi, "clients"],
  [/\\bProjeto\\b/g, "Project"],
  [/\\bprojetos\\b/gi, "projects"],
  [/\\bUnidade\\b/g, "Unit"],
  [/\\bunidades\\b/gi, "units"],
  [/\\bFornecedor\\b/g, "Supplier"],
  [/\\bFabricante\\b/g, "Manufacturer"],
  [/\\bEndereco\\b/g, "Address"],
  [/\\bTelefone\\b/g, "Phone"],
  [/\\bNome\\b/g, "Name"],
  [/\\bFuncao\\b/g, "Role"],
  [/\\bQualidade\\b/g, "Quality"],
  [/\\bInstalacao\\b/g, "Installation"],
  [/\\bConcluida\\b/g, "Completed"],
  [/\\bPendente\\b/g, "Pending"],
  [/\\bRejeitada\\b/g, "Rejected"],
  [/\\bAprovada\\b/g, "Approved"],
  [/\\bSem\\b/g, "No"],
  [/\\bNao\\b/g, "Not"],
  [/\\bEditar\\b/g, "Edit"],
  [/\\bExcluir\\b/g, "Delete"],
  [/\\bAdicionar\\b/g, "Add"],
  [/\\bRelatorio\\b/g, "Report"],
  [/\\bSincronizacao\\b/g, "Sync"],
  [/\\bAcoes\\b/g, "Actions"],
  [/\\bAcao\\b/g, "Action"],
  [/\\bData\\b/g, "Date"],
  [/\\bOrigem\\b/g, "Source"],
  [/\\bDetalhe\\b/g, "Detail"],
  [/\\bUltimo sync:/g, "Last sync:"],
  [/\\bVerificando conexao\\.\\.\\./g, "Checking connection..."],
  [/\\bFalha ao iniciar o app\\b/g, "App startup failed"],
  [/\\bSem sincronizacao\\b/g, "No sync"],
  [/\\bSem eventos\\./g, "No events."],
  [/\\bSem registros\\./g, "No records."],
  [/\\bSem itens cadastrados\\./g, "No items registered."],
  [/\\bSem itens no manifesto\\./g, "No manifest items."],
  [/\\bSem movimentos registrados\\./g, "No movement records."],
  [/\\bNenhum container cadastrado\\./g, "No containers registered."],
  [/\\bNenhum projeto cadastrado\\./g, "No projects registered."],
  [/\\bSem alteracoes auditadas\\./g, "No audited changes."],
  [/\\bSem tarefas de envio\\./g, "No dispatch tasks."],
  [/\\bSem recebimentos registrados\\./g, "No receipts recorded."],
  [/\\bNo dispatch signature\\b/g, "No dispatch signature"],
  [/\\bSem COI vencendo nos proximos 30 dias\\./g, "No COI expiring in the next 30 days."],
  [/\\bLembrete\\b/g, "Reminder"],
  [/\\bVence hoje\\b/g, "Expires today"],
  [/\\bVence em\\b/g, "Expires in"],
  [/\\bVencido ha\\b/g, "Expired for"],
];

function t(text) {
  if (currentLang === "pt") return text;
  return I18N[currentLang]?.[text] || text;
}

function roleLabel(role) {
  return ROLE_LABELS[currentLang]?.[role] || ROLE_LABEL[role] || role;
}

function contactRoleLabel(role) {
  return CONTACT_ROLE_LABELS[currentLang]?.[role] || CONTACT_ROLE_LABEL[role] || role;
}

function containerStatusLabel(status) {
  return CONTAINER_STATUS_LABELS[currentLang]?.[status] || CONTAINER_STATUS_LABEL[status] || status;
}

function stageLabel(key) {
  return STAGE_LABELS[currentLang]?.[key] || STAGE_LABELS.en[key] || key;
}

function translateDynamicText(text) {
  if (currentLang !== "en") return text;

  const lineUser = text.match(/^Usuario:\s(.+)\s\|\sPerfil:\s(.+)$/);
  if (lineUser) {
    return `User: ${lineUser[1]} | Role: ${lineUser[2]}`;
  }

  const lineSimple = text.match(/^(Cliente|Projeto|Fornecedor|Fabricante|ETA|Saida porto):\s(.+)$/);
  if (lineSimple) {
    const labelsEn = {
      Client: "Client",
      Project: "Project",
      Fornecedor: "Supplier",
      Fabricante: "Manufacturer",
      ETA: "ETA",
      "Saida porto": "Port departure",
    };
    return `${labelsEn[lineSimple[1]]}: ${lineSimple[2]}`;
  }

  const progress = text.match(/^(\d+)% concluido \((\d+)\/(\d+)\)$/);
  if (progress) {
    return `${progress[1]}% completed (${progress[2]}/${progress[3]})`;
  }

  const checked = text.match(/^Baixado:\s(.+)$/);
  if (checked) {
    return `Checked: ${checked[1]}`;
  }

  const items = text.match(/^Itens:\s(\d+)\s\|\sFaltantes:\s(\d+)\s\|\sDefeito:\s(\d+)\s\|\sAjuste:\s(\d+)$/);
  if (items) {
    return `Items: ${items[1]} | Missing: ${items[2]} | Damaged: ${items[3]} | Adjustment: ${items[4]}`;
  }

  const etaDays = text.match(/^ETA em (\d+) dia\(s\)\.$/);
  if (etaDays) {
    return `ETA in ${etaDays[1]} day(s).`;
  }

  const etaSoon = text.match(/^Container chega em (\d+) dia\(s\)\.$/);
  if (etaSoon) {
    return `Container arrives in ${etaSoon[1]} day(s).`;
  }

  const delayed = text.match(/^Atencao: container atrasado \((\d+) dia\(s\) de atraso\)\.$/);
  if (delayed) {
    return `Warning: delayed container (${delayed[1]} day(s) late).`;
  }

  const arrived = text.match(/^Container recebido em status Arrived\. ETA:\s(.+)\.$/);
  if (arrived) {
    return `Container received with status Arrived. ETA: ${arrived[1]}.`;
  }

  let out = text;
  PT_EN_REGEX_REPLACEMENTS.forEach(([pattern, replacement]) => {
    out = out.replace(pattern, replacement);
  });
  return out;
}

function translateRaw(text) {
  return translateDynamicText(t(text));
}

function translatePreservingWhitespace(value) {
  const match = value.match(/^(\s*)(.*?)(\s*)$/s);
  if (!match) return value;
  return `${match[1]}${translateRaw(match[2])}${match[3]}`;
}

function applyLanguageToRoot(root) {
  if (!root) return;

  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const textNodes = [];
  let current = walker.nextNode();
  while (current) {
    const parentTag = current.parentElement?.tagName;
    if (parentTag && !["SCRIPT", "STYLE", "NOSCRIPT"].includes(parentTag)) {
      textNodes.push(current);
    }
    current = walker.nextNode();
  }

  textNodes.forEach((node) => {
    if (!textNodeSource.has(node)) textNodeSource.set(node, node.nodeValue);
    const base = textNodeSource.get(node);
    node.nodeValue = translatePreservingWhitespace(base);
  });

  const attrs = ["placeholder", "title", "aria-label"];
  root.querySelectorAll("*").forEach((el) => {
    if (!attrSource.has(el)) attrSource.set(el, {});
    const saved = attrSource.get(el);
    attrs.forEach((attr) => {
      if (!el.hasAttribute(attr)) return;
      if (!saved[attr]) saved[attr] = el.getAttribute(attr);
      el.setAttribute(attr, translateRaw(saved[attr]));
    });
  });
}

function applyLanguageToUi() {
  applyLanguageToRoot(document.body);
  document.querySelectorAll("template").forEach((templateEl) => applyLanguageToRoot(templateEl.content));
}

function setLanguage(lang, { rerender = true } = {}) {
  const normalized = SUPPORTED_LANGS.includes(lang) ? lang : DEFAULT_LANG;
  currentLang = normalized;
  document.documentElement.lang = "en";
  if (languageSelect) languageSelect.value = normalized;

  if (rerender && currentUser) {
    render();
  } else {
    applyLanguageToUi();
  }

  updateConnectivity();
}

const nativeAlert = window.alert.bind(window);
const nativePrompt = window.prompt.bind(window);
const nativeConfirm = window.confirm.bind(window);

window.alert = (message) => nativeAlert(translateRaw(String(message ?? "")));
window.prompt = (message, defaultValue = "") => nativePrompt(translateRaw(String(message ?? "")), defaultValue);
window.confirm = (message) => nativeConfirm(translateRaw(String(message ?? "")));

let db;
let units = [];
let photos = [];
let users = [];
let clients = [];
let projects = [];
let contacts = [];
let containers = [];
let materials = [];
let syncConfig = { ...DEFAULT_SYNC };
let photosByUnit = new Map();
let currentUser = null;
let deferredPrompt = null;
let autoSyncTimer = null;
let currentView = "home";
let editingUserId = "";
let userEditReturnView = "home";
let adminEditingUserId = "";
let lastQrLookupCode = "";
let currentProjectSector = "fabrica";
let usersViewFilter = "all";
let hasAutoDispatchedCoiReminder = false;
let selectedClientId = "";
let clientSearchQuery = "";
let projectsViewMode = "overview";
let appAuditLog = loadAppAuditLog();

const authView = document.getElementById("authView");
const bootErrorPanel = document.getElementById("bootErrorPanel");
const bootErrorMessage = document.getElementById("bootErrorMessage");
const setupPanel = document.getElementById("setupPanel");
const loginPanel = document.getElementById("loginPanel");
const signupPanel = document.getElementById("signupPanel");
const setupForm = document.getElementById("setupForm");
const loginForm = document.getElementById("loginForm");
const signupForm = document.getElementById("signupForm");
const openSignupBtn = document.getElementById("openSignupBtn");
const backToLoginBtn = document.getElementById("backToLoginBtn");
const signupPhoneInput = document.getElementById("signupPhone");
const signupCellPhoneInput = document.getElementById("signupCellPhone");
const signupEmploymentTypeSelect = document.getElementById("signupEmploymentType");

const appMain = document.getElementById("appMain");
const roleLine = document.getElementById("roleLine");
const permissionLine = document.getElementById("permissionLine");
const homePanel = document.getElementById("homePanel");
const coiReminderPanel = document.getElementById("coiReminderPanel");
const quickNavPanel = document.getElementById("quickNavPanel");

const masterDataPanel = document.getElementById("masterDataPanel");
const clientsSubpanel = document.getElementById("clientsSubpanel");
const projectsSubpanel = document.getElementById("projectsSubpanel");
const contactsSubpanel = document.getElementById("contactsSubpanel");
const materialsSubpanel = document.getElementById("materialsSubpanel");
const clientForm = document.getElementById("clientForm");
const projectForm = document.getElementById("projectForm");
const contactForm = document.getElementById("contactForm");
const materialForm = document.getElementById("materialForm");
const clientsSummaryScroll = document.getElementById("clientsSummaryScroll");
const clientDetailsPanel = document.getElementById("clientDetailsPanel");
const clientAdminTarget = document.getElementById("clientAdminTarget");
const clientLogoPreview = document.getElementById("clientLogoPreview");
const clientFormNewBtn = document.getElementById("clientFormNewBtn");
const clientFormDeleteBtn = document.getElementById("clientFormDeleteBtn");
const goProjectsFromClientBtn = document.getElementById("goProjectsFromClientBtn");
const goPeopleFromClientBtn = document.getElementById("goPeopleFromClientBtn");
const clientSearchInput = document.getElementById("clientSearchInput");
const projectsTable = document.getElementById("projectsTable");
const contactsTable = document.getElementById("contactsTable");
const materialsTable = document.getElementById("materialsTable");
const projectClientSelect = document.getElementById("projectClientSelect");
const contactClientSelect = document.getElementById("contactClientSelect");
const contactProjectSelect = document.getElementById("contactProjectSelect");

const containerPanel = document.getElementById("containerPanel");
const containerForm = document.getElementById("containerForm");
const containerClientSelect = document.getElementById("containerClientSelect");
const containerProjectSelect = document.getElementById("containerProjectSelect");
const containersBoard = document.getElementById("containersBoard");
const containerTemplate = document.getElementById("containerTemplate");
const projectSectorPanel = document.getElementById("projectSectorPanel");

const unitForm = document.getElementById("unitForm");
const unitEntryPanel = document.getElementById("unitEntryPanel");
const unitProjectSelect = document.getElementById("unitProjectSelect");
const unitClientName = document.getElementById("unitClientName");
const unitProjectName = document.getElementById("unitProjectName");
const unitJobSite = document.getElementById("unitJobSite");
const unitProjectHint = document.getElementById("unitProjectHint");

const unitsContainer = document.getElementById("unitsContainer");
const unitTemplate = document.getElementById("unitTemplate");
const statsPanel = document.getElementById("statsPanel");
const unitsToolbarPanel = document.getElementById("unitsToolbarPanel");
const searchInput = document.getElementById("searchInput");
const filterSelect = document.getElementById("filterSelect");
const qrLookupPanel = document.getElementById("qrLookupPanel");
const qrLookupForm = document.getElementById("qrLookupForm");
const qrLookupInput = document.getElementById("qrLookupInput");
const qrLookupResult = document.getElementById("qrLookupResult");

const exportBtn = document.getElementById("exportBtn");
const importInput = document.getElementById("importInput");
const projectReportSelect = document.getElementById("projectReportSelect");
const projectReportBtn = document.getElementById("projectReportBtn");

const syncPanel = document.getElementById("syncPanel");
const syncConfigForm = document.getElementById("syncConfigForm");
const pushSyncBtn = document.getElementById("pushSyncBtn");
const pullSyncBtn = document.getElementById("pullSyncBtn");
const syncStatus = document.getElementById("syncStatus");
const developerAuditPanel = document.getElementById("developerAuditPanel");

const usersPanel = document.getElementById("usersPanel");
const userForm = document.getElementById("userForm");
const usersNameRail = document.getElementById("usersNameRail");
const usersTable = document.getElementById("usersTable");
const userAdminTarget = document.getElementById("userAdminTarget");
const userAdminPhotoPreview = document.getElementById("userAdminPhotoPreview");
const userAdminEmploymentTypeSelect = document.getElementById("userAdminEmploymentType");
const userAdminPhoneInput = document.getElementById("userAdminPhone");
const userAdminCellPhoneInput = document.getElementById("userAdminCellPhone");
const userAdminCoiFileStatus = document.getElementById("userAdminCoiFileStatus");
const openUserAdminCoiFileBtn = document.getElementById("openUserAdminCoiFileBtn");
const userFormNewBtn = document.getElementById("userFormNewBtn");
const userFormDeleteBtn = document.getElementById("userFormDeleteBtn");
const userEditPanel = document.getElementById("userEditPanel");
const userEditForm = document.getElementById("userEditForm");
const userEditTarget = document.getElementById("userEditTarget");
const userEditPhotoPreview = document.getElementById("userEditPhotoPreview");
const userEditCoiFileStatus = document.getElementById("userEditCoiFileStatus");
const openUserEditCoiFileBtn = document.getElementById("openUserEditCoiFileBtn");
const cancelUserEditBtn = document.getElementById("cancelUserEditBtn");
const userEditPhoneInput = document.getElementById("userEditPhone");
const userEditCellPhoneInput = document.getElementById("userEditCellPhone");
const userEditEmploymentTypeSelect = document.getElementById("userEditEmploymentType");

const connectionBadge = document.getElementById("connectionBadge");
const userBadge = document.getElementById("userBadge");
const editProfileBtn = document.getElementById("editProfileBtn");
const logoutBtn = document.getElementById("logoutBtn");
const installBtn = document.getElementById("installBtn");
const languageSelect = document.getElementById("languageSelect");
if (languageSelect) languageSelect.disabled = SUPPORTED_LANGS.length === 1;

const fmtDate = (iso) => {
  if (!iso) return "-";
  return new Date(iso).toLocaleString("en-US", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
};

const uid = () => {
  if (crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
};

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function toNumber(value) {
  const num = Number(value || 0);
  return Number.isFinite(num) && num > 0 ? num : 0;
}

function maskPhoneValue(raw) {
  const digits = String(raw || "")
    .replace(/\D/g, "")
    .slice(0, 12);
  if (!digits) return "";
  const parts = [];
  if (digits.length <= 2) return digits;
  parts.push(digits.slice(0, 2));
  if (digits.length <= 5) {
    parts.push(digits.slice(2));
    return parts.join("-");
  }
  parts.push(digits.slice(2, 5));
  if (digits.length <= 8) {
    parts.push(digits.slice(5));
    return parts.join("-");
  }
  parts.push(digits.slice(5, 8));
  parts.push(digits.slice(8));
  return parts.join("-");
}

function collectCheckedValues(form, name) {
  if (!form) return [];
  return Array.from(form.querySelectorAll(`input[name="${name}"]:checked`))
    .map((entry) => entry.value)
    .filter(Boolean);
}

function setCheckedValues(form, name, values) {
  if (!form) return;
  const target = new Set(ensureArray(values));
  form.querySelectorAll(`input[name="${name}"]`).forEach((entry) => {
    entry.checked = target.has(entry.value);
  });
}

function toggleSubcontractorExtras(scope, employmentType) {
  const show = employmentType === "subcontractor";
  const requiredFields = new Set(["contractorCoi", "contractorCoiExpiry", "contractorW9"]);
  document.querySelectorAll(`[data-subcontractor-fields="${scope}"]`).forEach((wrapper) => {
    wrapper.classList.toggle("hidden", !show);
    wrapper.querySelectorAll("input,select,textarea").forEach((field) => {
      if (requiredFields.has(field.name)) field.required = show;
      if (field.name === "contractorCoiFile") field.required = show && scope === "signup";
      if (!show && field.type === "checkbox") field.checked = false;
      if (!show && field.type === "file") field.value = "";
      if (!show && field.type !== "checkbox") field.value = "";
    });
  });
}

function validateCoiFile(file) {
  if (!file) return { ok: false, message: "Attach the COI file in PDF or JPG format." };
  const fileType = (file.type || "").toLowerCase();
  const fileName = (file.name || "").toLowerCase();
  const validType = fileType === "application/pdf" || fileType === "image/jpeg";
  const validExt = /\.(pdf|jpe?g)$/.test(fileName);
  if (!validType && !validExt) {
    return { ok: false, message: "Invalid COI file. Use PDF or JPG." };
  }
  if ((file.size || 0) > COI_MAX_FILE_SIZE) {
    return { ok: false, message: "COI file is too large. Limit is 8 MB." };
  }
  return { ok: true, message: "" };
}

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const now = new Date();
  const target = new Date(`${dateStr}T00:00:00`);
  const diff = target.getTime() - new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

function fmtDateOnly(iso) {
  if (!iso) return "-";
  const value = iso.includes("T") ? iso : `${iso}T00:00:00`;
  return new Date(value).toLocaleDateString("en-US");
}

function elapsedFrom(startIso) {
  if (!startIso) return "-";
  const ms = Date.now() - new Date(startIso).getTime();
  if (!Number.isFinite(ms) || ms < 0) return "-";
  const totalMinutes = Math.floor(ms / 60000);
  const days = Math.floor(totalMinutes / 1440);
  const hours = Math.floor((totalMinutes % 1440) / 60);
  const minutes = totalMinutes % 60;
  return `${days}d ${hours}h ${minutes}m`;
}

function ensureArray(value) {
  return Array.isArray(value) ? value : [];
}

function coiReminderWindowUsers() {
  return users
    .filter((user) => user.employmentType === "subcontractor" && user.email && user.contractorCoiExpiry && canAccessEmployeeRecord(user))
    .map((user) => {
      const daysLeft = daysUntil(user.contractorCoiExpiry);
      return {
        ...user,
        daysLeft,
        reminderSentForCurrentExpiry:
          user.contractorCoiReminderForExpiry === user.contractorCoiExpiry && Boolean(user.contractorCoiReminderSentAt),
      };
    })
    .filter((user) => user.daysLeft !== null && user.daysLeft <= COI_REMINDER_DAYS)
    .sort((a, b) => a.daysLeft - b.daysLeft);
}

function pendingCoiReminderUsers() {
  return coiReminderWindowUsers().filter((user) => !user.reminderSentForCurrentExpiry);
}

function coiReminderLabel(daysLeft) {
  if (daysLeft < 0) return `Expired for ${Math.abs(daysLeft)} day(s)`;
  if (daysLeft === 0) return "Expires today";
  return `Expires in ${daysLeft} day(s)`;
}

function buildCoiReminderMailto(user) {
  const dueDate = fmtDateOnly(user.contractorCoiExpiry);
  const subject = `COI Renewal Required - ${user.companyName || user.name || "Sub Contractor"}`;
  const body = [
    `Hello ${user.firstName || user.name || "team"},`,
    "",
    `Your COI document is set to expire on ${dueDate}.`,
    "Please send the updated COI before the expiration date to keep your access active.",
    "",
    "Thank you.",
    "TAG Team",
  ].join("\n");
  return `mailto:${encodeURIComponent(user.email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

async function dispatchCoiReminder(userId, { automatic = false } = {}) {
  if (!currentUser || (!isAdmin() && !isDeveloper())) return false;
  const targetUser = users.find((entry) => entry.id === userId);
  if (!targetUser) return false;
  if (!canAccessEmployeeRecord(targetUser)) return false;
  if (targetUser.employmentType !== "subcontractor" || !targetUser.email || !targetUser.contractorCoiExpiry) return false;
  if (
    targetUser.contractorCoiReminderForExpiry === targetUser.contractorCoiExpiry &&
    targetUser.contractorCoiReminderSentAt
  ) {
    if (!automatic) alert("COI reminder already sent for this expiration date.");
    return false;
  }

  const daysLeft = daysUntil(targetUser.contractorCoiExpiry);
  if (daysLeft === null || daysLeft > COI_REMINDER_DAYS) {
    if (!automatic) alert("COI ainda nao esta na janela de aviso (30 dias).");
    return false;
  }

  window.location.href = buildCoiReminderMailto(targetUser);

  await put(
    USER_STORE,
    normalizeUser({
      ...targetUser,
      contractorCoiReminderSentAt: new Date().toISOString(),
      contractorCoiReminderForExpiry: targetUser.contractorCoiExpiry,
      updatedAt: new Date().toISOString(),
    })
  );
  pushEntityAudit(
    "Users",
    "updated",
    `${targetUser.name || targetUser.username}: COI reminder sent for ${targetUser.contractorCoiExpiry}`,
    "users"
  );
  await loadAll();
  if (currentUser) currentUser = users.find((entry) => entry.id === currentUser.id) || currentUser;
  render();
  queueAutoSync();
  if (!automatic) alert("Email de renovacao do COI disparado para o app de email.");
  return true;
}

async function maybeAutoDispatchCoiReminder() {
  if (hasAutoDispatchedCoiReminder || !currentUser) return;
  if (!isAdmin() && !isDeveloper()) return;
  hasAutoDispatchedCoiReminder = true;

  const pending = pendingCoiReminderUsers();
  if (!pending.length) return;
  await dispatchCoiReminder(pending[0].id, { automatic: true });
}

function pushAudit(log, message) {
  const entry = {
    id: uid(),
    at: new Date().toISOString(),
    byUserId: currentUser?.id || "",
    byName: currentUser?.name || "System",
    message,
  };
  const out = ensureArray(log);
  out.push(entry);
  return out;
}

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const dbRef = event.target.result;
      if (!dbRef.objectStoreNames.contains(UNIT_STORE)) dbRef.createObjectStore(UNIT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(PHOTO_STORE)) {
        const store = dbRef.createObjectStore(PHOTO_STORE, { keyPath: "id" });
        store.createIndex("unitId", "unitId", { unique: false });
      }
      if (!dbRef.objectStoreNames.contains(USER_STORE)) dbRef.createObjectStore(USER_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(SETTINGS_STORE)) dbRef.createObjectStore(SETTINGS_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(CLIENT_STORE)) dbRef.createObjectStore(CLIENT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(PROJECT_STORE)) dbRef.createObjectStore(PROJECT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(CONTACT_STORE)) dbRef.createObjectStore(CONTACT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(CONTAINER_STORE)) dbRef.createObjectStore(CONTAINER_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(MATERIAL_STORE)) dbRef.createObjectStore(MATERIAL_STORE, { keyPath: "id" });
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("Banco local bloqueado por outra aba/dispositivo."));
  });
}

function showBootError(message) {
  authView.classList.remove("hidden");
  appMain.classList.add("hidden");
  userBadge.classList.add("hidden");
  editProfileBtn?.classList.add("hidden");
  logoutBtn.classList.add("hidden");
  setupPanel.classList.add("hidden");
  loginPanel.classList.add("hidden");
  signupPanel?.classList.add("hidden");
  bootErrorPanel.classList.remove("hidden");
  bootErrorMessage.textContent = message;
  applyLanguageToUi();
}

function tx(storeName, mode = "readonly") {
  return db.transaction(storeName, mode).objectStore(storeName);
}

function requestToPromise(req) {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function getAll(storeName) {
  return requestToPromise(tx(storeName).getAll());
}

async function put(storeName, value) {
  return requestToPromise(tx(storeName, "readwrite").put(value));
}

async function del(storeName, id) {
  return requestToPromise(tx(storeName, "readwrite").delete(id));
}

function normalizeUnit(unit) {
  const normalized = { ...unit };
  normalized.stages = normalized.stages || {};
  for (const stage of STAGES) {
    normalized.stages[stage.key] = normalized.stages[stage.key] || { done: false, at: null };
  }
  normalized.checkItems = ensureArray(normalized.checkItems);
  normalized.dispatchTasks = ensureArray(normalized.dispatchTasks);
  normalized.siteReceipts = ensureArray(normalized.siteReceipts).map((entry) => ({
    id: entry.id || uid(),
    qrCode: entry.qrCode || "",
    item: entry.item || "",
    qty: toNumber(entry.qty || 0),
    issue: entry.issue || "ok",
    note: entry.note || "",
    receivedByUserId: entry.receivedByUserId || "",
    receivedBy: entry.receivedBy || "",
    createdAt: entry.createdAt || new Date().toISOString(),
    updatedAt: entry.updatedAt || entry.createdAt || new Date().toISOString(),
  }));
  normalized.auditLog = ensureArray(normalized.auditLog);
  normalized.dispatchSignature = normalized.dispatchSignature || null;
  normalized.deliveryQuality = normalized.deliveryQuality || "pending";
  normalized.installationStatus = normalized.installationStatus || "not-started";
  normalized.deliveryNotes = normalized.deliveryNotes || "";
  normalized.issuesText = normalized.issuesText || "";
  normalized.clientId = normalized.clientId || "";
  normalized.projectId = normalized.projectId || "";
  normalized.clientName = normalized.clientName || "";
  normalized.projectName = normalized.projectName || "";
  normalized.createdAt = normalized.createdAt || new Date().toISOString();
  normalized.updatedAt = normalized.updatedAt || normalized.createdAt;
  return normalized;
}

function normalizeClient(client) {
  return {
    id: client.id,
    name: client.name || "",
    address: client.address || "",
    officePhone: client.officePhone || client.phone || "",
    logoDataUrl: client.logoDataUrl || "",
    ownerName: client.ownerName || "",
    contactOwner: client.contactOwner || "",
    contactProjectManagers: client.contactProjectManagers || "",
    contactSeniorProjectManager: client.contactSeniorProjectManager || "",
    seniorProjectManagerPhone: client.seniorProjectManagerPhone || "",
    contactPerson: client.contactPerson || "",
    offices: client.offices || "",
    yearsInBusiness: client.yearsInBusiness || "",
    phone: client.phone || client.officePhone || "",
    email: client.email || "",
    createdAt: client.createdAt || new Date().toISOString(),
    updatedAt: client.updatedAt || client.createdAt || new Date().toISOString(),
  };
}

function normalizeProject(project) {
  return {
    id: project.id,
    clientId: project.clientId || "",
    name: project.name || "",
    code: project.code || "",
    address: project.address || "",
    createdAt: project.createdAt || new Date().toISOString(),
    updatedAt: project.updatedAt || project.createdAt || new Date().toISOString(),
  };
}

function normalizeContact(contact) {
  return {
    id: contact.id,
    name: contact.name || "",
    role: contact.role || "other",
    company: contact.company || "",
    clientId: contact.clientId || "",
    projectId: contact.projectId || "",
    phone: contact.phone || "",
    email: contact.email || "",
    createdAt: contact.createdAt || new Date().toISOString(),
    updatedAt: contact.updatedAt || contact.createdAt || new Date().toISOString(),
  };
}

function normalizeUser(user) {
  const firstName = user.firstName || "";
  const lastName = user.lastName || "";
  const composedName = `${firstName} ${lastName}`.trim();
  const normalizedCoiFile =
    user.contractorCoiFile && typeof user.contractorCoiFile === "object"
      ? {
          name: user.contractorCoiFile.name || "",
          type: user.contractorCoiFile.type || "",
          dataUrl: user.contractorCoiFile.dataUrl || "",
          uploadedAt: user.contractorCoiFile.uploadedAt || "",
        }
      : null;
  return {
    id: user.id,
    name: user.name || composedName || "",
    firstName,
    lastName,
    companyName: user.companyName || "",
    jobTitle: user.jobTitle || "",
    employmentType: user.employmentType || "",
    contractorCoi: user.contractorCoi || "",
    contractorCoiExpiry: user.contractorCoiExpiry || "",
    contractorCoiFile: normalizedCoiFile,
    contractorW9: user.contractorW9 || "",
    contractorAreas: ensureArray(user.contractorAreas),
    contractorCoiReminderSentAt: user.contractorCoiReminderSentAt || "",
    contractorCoiReminderForExpiry: user.contractorCoiReminderForExpiry || "",
    address: user.address || "",
    phone: user.phone || "",
    cellPhone: user.cellPhone || "",
    email: user.email || "",
    photoDataUrl: user.photoDataUrl || "",
    username: (user.username || "").toLowerCase(),
    passwordHash: user.passwordHash || "",
    role: user.role || "visitor",
    createdAt: user.createdAt || new Date().toISOString(),
    updatedAt: user.updatedAt || user.createdAt || new Date().toISOString(),
  };
}

function normalizeMaterial(material) {
  return {
    id: material.id,
    sku: material.sku || "",
    description: material.description || "",
    category: material.category || "other",
    unit: material.unit || "pcs",
    kitchenType: material.kitchenType || "",
    createdAt: material.createdAt || new Date().toISOString(),
    updatedAt: material.updatedAt || material.createdAt || new Date().toISOString(),
  };
}

function normalizeContainer(container) {
  const materialItems = ensureArray(container.materialItems).map((item) => ({
    id: item.id || uid(),
    materialId: item.materialId || "",
    code: item.code || "",
    description: item.description || "",
    category: item.category || "other",
    qty: toNumber(item.qty),
    unit: item.unit || "",
    issueType: item.issueType || "ok",
    issueNote: item.issueNote || "",
    issueRoute: item.issueRoute || "",
    createdAt: item.createdAt || new Date().toISOString(),
    updatedAt: item.updatedAt || item.createdAt || new Date().toISOString(),
  }));
  const qrItems = ensureArray(container.qrItems).map((item) => ({
    id: item.id || uid(),
    manifestItemId: item.manifestItemId || "",
    qrCode: item.qrCode || "",
    materialId: item.materialId || "",
    code: item.code || "",
    description: item.description || "",
    category: item.category || "other",
    unit: item.unit || "",
    status: item.status || "in-warehouse",
    clientId: item.clientId || "",
    projectId: item.projectId || "",
    unitId: item.unitId || "",
    unitLabel: item.unitLabel || "",
    kitchenType: item.kitchenType || "",
    assignedAt: item.assignedAt || "",
    assignedBy: item.assignedBy || "",
    deliveredAt: item.deliveredAt || "",
    issueType: item.issueType || "",
    issueNote: item.issueNote || "",
    updatedAt: item.updatedAt || item.createdAt || new Date().toISOString(),
    createdAt: item.createdAt || new Date().toISOString(),
  }));
  return {
    id: container.id,
    containerCode: container.containerCode || "",
    supplier: container.supplier || "",
    manufacturer: container.manufacturer || "",
    clientId: container.clientId || "",
    projectId: container.projectId || "",
    etaDate: container.etaDate || "",
    departureAt: container.departureAt || "",
    arrivalStatus: container.arrivalStatus || "scheduled",
    qtyKitchens: toNumber(container.qtyKitchens),
    qtyVanities: toNumber(container.qtyVanities),
    qtyMedCabinets: toNumber(container.qtyMedCabinets),
    qtyCountertops: toNumber(container.qtyCountertops),
    notes: container.notes || "",
    loosePartsNotes: container.loosePartsNotes || "",
    materialItems,
    qrItems,
    movements: ensureArray(container.movements),
    replacementQueue: ensureArray(container.replacementQueue),
    factoryMissingQueue: ensureArray(container.factoryMissingQueue),
    auditLog: ensureArray(container.auditLog),
    createdAt: container.createdAt || new Date().toISOString(),
    updatedAt: container.updatedAt || container.createdAt || new Date().toISOString(),
  };
}

async function loadAll() {
  const [unitRows, photoRows, userRows, settingsRows, clientRows, projectRows, contactRows, containerRows, materialRows] =
    await Promise.all([
      getAll(UNIT_STORE),
      getAll(PHOTO_STORE),
      getAll(USER_STORE),
      getAll(SETTINGS_STORE),
      getAll(CLIENT_STORE),
      getAll(PROJECT_STORE),
      getAll(CONTACT_STORE),
      getAll(CONTAINER_STORE),
      getAll(MATERIAL_STORE),
    ]);

  units = unitRows.map(normalizeUnit).sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
  photos = photoRows;
  users = userRows.map(normalizeUser).sort((a, b) => a.username.localeCompare(b.username));
  clients = clientRows.map(normalizeClient).sort((a, b) => a.name.localeCompare(b.name));
  projects = projectRows.map(normalizeProject).sort((a, b) => a.name.localeCompare(b.name));
  contacts = contactRows.map(normalizeContact).sort((a, b) => a.name.localeCompare(b.name));
  containers = containerRows.map(normalizeContainer).sort((a, b) => {
    const ad = a.etaDate ? new Date(a.etaDate).getTime() : 0;
    const bd = b.etaDate ? new Date(b.etaDate).getTime() : 0;
    return ad - bd;
  });
  materials = materialRows.map(normalizeMaterial).sort((a, b) => a.sku.localeCompare(b.sku));

  const savedSync = settingsRows.find((row) => row.id === "syncConfig") || {};
  syncConfig = {
    ...DEFAULT_SYNC,
    ...savedSync,
    autoSync: Boolean(savedSync.autoSync),
  };

  photosByUnit = new Map();
  for (const photo of photos) {
    if (!photosByUnit.has(photo.unitId)) photosByUnit.set(photo.unitId, []);
    photosByUnit.get(photo.unitId).push(photo);
  }
}

function normalizePersonName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ");
}

function isPrimaryDeveloperUser(user) {
  const fullName = user.name || `${user.firstName || ""} ${user.lastName || ""}`;
  return normalizePersonName(fullName) === normalizePersonName(PRIMARY_DEVELOPER_NAME);
}

async function ensureDeveloperRolePresence() {
  const preferredDeveloper = users.find((user) => isPrimaryDeveloperUser(user));

  if (preferredDeveloper && preferredDeveloper.role !== "developer") {
    await put(
      USER_STORE,
      normalizeUser({
        ...preferredDeveloper,
        role: "developer",
        updatedAt: new Date().toISOString(),
      })
    );
    await loadAll();
    return;
  }

  if (users.some((user) => user.role === "developer")) return;
  const fallbackAdmin = users.find((user) => user.role === "admin");
  if (!fallbackAdmin) return;
  await put(
    USER_STORE,
    normalizeUser({
      ...fallbackAdmin,
      role: "developer",
      updatedAt: new Date().toISOString(),
    })
  );
  await loadAll();
}

function clientById(id) {
  return clients.find((client) => client.id === id);
}

function projectById(id) {
  return projects.find((project) => project.id === id);
}

function materialById(id) {
  return materials.find((material) => material.id === id);
}

function contactsForProject(projectId) {
  return contacts.filter((contact) => contact.projectId === projectId);
}

function nextQrSequence(container, materialCode) {
  const seq = container.qrItems.filter((item) => item.code === materialCode).length + 1;
  return String(seq).padStart(4, "0");
}

function makeQrCode(container, materialCode) {
  const seq = nextQrSequence(container, materialCode);
  return `QR|${container.containerCode}|${materialCode}|${seq}|${uid().slice(0, 8)}`;
}

function qrImageUrl(qrCode, size = 300) {
  const normalized = encodeURIComponent(String(qrCode || ""));
  return `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${normalized}`;
}

function qrLabelFilename(item) {
  const base = String(item?.qrCode || "qr-label").replace(/[^a-zA-Z0-9_-]+/g, "_");
  return `${base}.png`;
}

function openQrLabelWindow(container, item, { autoPrint = false } = {}) {
  if (!item?.qrCode) return;
  const client = clientById(container.clientId);
  const project = projectById(container.projectId);
  const unit = units.find((entry) => entry.id === item.unitId);
  const imageUrl = qrImageUrl(item.qrCode, 320);
  const win = window.open("", "_blank");
  if (!win) return;
  const html = `<!doctype html>
  <html lang="en">
    <head>
      <meta charset="utf-8" />
      <title>QR Label ${escapeHtml(item.code || "")}</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; color: #1f2937; }
        .label { width: 340px; border: 1px solid #d1d5db; border-radius: 12px; padding: 14px; }
        h1 { margin: 0 0 8px; font-size: 14px; letter-spacing: 0.02em; }
        img { width: 220px; height: 220px; border: 1px solid #e5e7eb; display: block; margin: 10px auto; }
        p { margin: 4px 0; font-size: 12px; }
        .qr-code { font-family: "Courier New", monospace; word-break: break-all; font-size: 11px; }
        .actions { margin-top: 12px; display: flex; gap: 8px; }
        button { border: 1px solid #cbd5e1; border-radius: 999px; background: #fff; padding: 6px 10px; cursor: pointer; }
        @media print { .actions { display: none; } body { margin: 0; } .label { border: none; border-radius: 0; } }
      </style>
    </head>
    <body>
      <div class="label">
        <h1>${escapeHtml(item.code || "-")} | ${escapeHtml(item.description || "-")}</h1>
        <img src="${imageUrl}" alt="QR ${escapeHtml(item.qrCode)}" />
        <p><strong>QR:</strong> <span class="qr-code">${escapeHtml(item.qrCode)}</span></p>
        <p><strong>Container:</strong> ${escapeHtml(container.containerCode)}</p>
        <p><strong>Client:</strong> ${escapeHtml(client?.name || "-")}</p>
        <p><strong>Project:</strong> ${escapeHtml(project?.name || "-")}</p>
        <p><strong>Unit:</strong> ${escapeHtml(item.unitLabel || unit?.unitCode || "-")}</p>
        <p><strong>Kitchen Type:</strong> ${escapeHtml(item.kitchenType || "-")}</p>
        <div class="actions">
          <button onclick="window.print()">Print / Save PDF</button>
        </div>
      </div>
    </body>
  </html>`;
  win.document.open();
  win.document.write(html);
  win.document.close();
  if (autoPrint) setTimeout(() => win.print(), 450);
}

async function downloadQrLabelPng(item) {
  if (!item?.qrCode) return;
  const url = qrImageUrl(item.qrCode, 512);
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error("QR image download failed");
    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = objectUrl;
    a.download = qrLabelFilename(item);
    a.click();
    URL.revokeObjectURL(objectUrl);
  } catch {
    window.open(url, "_blank");
  }
}

function buildQrHistory(qrCode) {
  const container = findQrItemContainer(qrCode);
  if (!container) return null;
  const item = container.qrItems.find((entry) => entry.qrCode === qrCode);
  if (!item) return null;

  const client = clientById(container.clientId);
  const project = projectById(container.projectId);
  const linkedUnit = units.find((entry) => entry.id === item.unitId) || null;
  const receiptUnits = units.filter((entry) => entry.siteReceipts.some((receipt) => receipt.qrCode === qrCode));
  const unit = linkedUnit || receiptUnits[0] || null;
  const events = [];

  events.push({
    at: item.createdAt,
    source: "Manifesto",
    action: "QR gerado",
    detail: `${item.code || "-"} | ${item.description || "-"}`,
  });
  if (item.assignedAt) {
    events.push({
      at: item.assignedAt,
      source: "Dispatch",
      action: "Item direcionado",
      detail: `${item.unitLabel || "-"}${item.kitchenType ? ` | ${item.kitchenType}` : ""}`,
    });
  }
  if (item.deliveredAt) {
    events.push({
      at: item.deliveredAt,
      source: "Site Job",
      action: item.issueType ? "Recebido com problema" : "Recebido",
      detail: item.issueType ? `${item.issueType}${item.issueNote ? ` | ${item.issueNote}` : ""}` : "OK",
    });
  }

  ensureArray(container.replacementQueue)
    .filter((entry) => entry.qrCode === qrCode)
    .forEach((entry) => {
      events.push({
        at: entry.createdAt,
        source: "Warehouse",
        action: "Fila de substituicao",
        detail: `${entry.issueType}${entry.note ? ` | ${entry.note}` : ""}`,
      });
    });

  ensureArray(container.factoryMissingQueue)
    .filter((entry) => entry.qrCode === qrCode)
    .forEach((entry) => {
      events.push({
        at: entry.createdAt,
        source: "Factory Claim",
        action: "Fila de faltantes da fabrica",
        detail: `${entry.issueType}${entry.note ? ` | ${entry.note}` : ""}`,
      });
    });

  receiptUnits.forEach((receiptUnit) => {
    receiptUnit.siteReceipts
      .filter((entry) => entry.qrCode === qrCode)
      .forEach((entry) => {
        events.push({
          at: entry.createdAt,
          source: "Site Job",
          action: "Lancamento de recebimento",
          detail: `${entry.item} (${entry.qty}) - ${entry.issue}${entry.note ? ` | ${entry.note}` : ""}`,
        });
      });
  });

  ensureArray(container.auditLog)
    .filter((entry) => String(entry.message || "").includes(qrCode))
    .forEach((entry) => {
      events.push({
        at: entry.at,
        source: entry.byName || "Audit",
        action: "Audit",
        detail: entry.message || "",
      });
    });

  if (unit) {
    ensureArray(unit.auditLog)
      .filter(
        (entry) =>
          String(entry.message || "").includes(qrCode) ||
          String(entry.message || "").includes("Site job receipt by QR") ||
          String(entry.message || "").includes("Recebimento no site job por QR")
      )
      .forEach((entry) => {
        events.push({
          at: entry.at,
          source: entry.byName || "Unit Audit",
          action: "Unit audit",
          detail: entry.message || "",
        });
      });
  }

  events.sort((a, b) => new Date(b.at || 0).getTime() - new Date(a.at || 0).getTime());
  return { container, item, client, project, unit, events };
}

function renderQrTrackTable(wrapper, container) {
  const rows = ensureArray(container.qrItems)
    .slice()
    .reverse()
    .map(
      (item) => `<tr>
      <td>${escapeHtml(item.qrCode)}</td>
      <td>${escapeHtml(item.code)}</td>
      <td>${escapeHtml(item.description)}</td>
      <td>${escapeHtml(item.status)}</td>
      <td>${escapeHtml(item.unitLabel || "-")}</td>
      <td>${escapeHtml(item.kitchenType || "-")}</td>
      <td>${escapeHtml(fmtDate(item.updatedAt || item.createdAt))}</td>
      <td><div class="actions-inline"><button class="secondary xs-btn" data-qr-print="${item.id}">PDF/Print</button><button class="secondary xs-btn" data-qr-png="${item.id}">PNG</button></div></td>
    </tr>`
    )
    .join("");
  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>QR</th><th>SKU</th><th>Item</th><th>Status</th><th>Unit</th><th>Kitchen Type</th><th>Updated</th><th>Label</th></tr></thead><tbody>${
    rows || '<tr><td colspan="8">No QR items generated.</td></tr>'
  }</tbody></table>`;
}

function renderQueueTable(wrapper, list) {
  const rows = ensureArray(list)
    .slice()
    .reverse()
    .map(
      (entry) => `<tr>
      <td>${escapeHtml(fmtDate(entry.createdAt))}</td>
      <td>${escapeHtml(entry.qrCode)}</td>
      <td>${escapeHtml(entry.code)}</td>
      <td>${escapeHtml(entry.description)}</td>
      <td>${escapeHtml(entry.unitLabel || "-")}</td>
      <td>${escapeHtml(entry.issueType)}</td>
      <td>${escapeHtml(entry.note || "-")}</td>
    </tr>`
    )
    .join("");
  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>Date</th><th>QR</th><th>SKU</th><th>Description</th><th>Unit</th><th>Issue</th><th>Detail</th></tr></thead><tbody>${
    rows || '<tr><td colspan="7">No records.</td></tr>'
  }</tbody></table>`;
}

function findQrItemContainer(qrCode, projectId = "") {
  return containers.find((container) => {
    if (projectId && container.projectId !== projectId) return false;
    return ensureArray(container.qrItems).some((item) => item.qrCode === qrCode);
  });
}

function qrIssueRoutesToFactory(issueType) {
  return issueType === "nao-chegou" || issueType === "faltando-pecas";
}

function isAdmin() {
  return currentUser?.role === "admin";
}

function isDeveloper() {
  return currentUser?.role === "developer";
}

function normalizeCompanyName(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function isTagAdminUser(user = currentUser) {
  if (!user || user.role !== "admin") return false;
  const employmentType = String(user.employmentType || "").toLowerCase();
  const company = normalizeCompanyName(user.companyName);
  const isTagCompany = company.includes("tag");
  const isNoLimitCompany = company.includes("no limit") || company.includes("nolimit");
  return employmentType === "tag" || isTagCompany || isNoLimitCompany;
}

function canViewAllEmployees() {
  if (isDeveloper()) return true;
  return isTagAdminUser(currentUser);
}

function canAccessEmployeeRecord(targetUser) {
  if (!currentUser || !targetUser) return false;
  if (isDeveloper()) return true;
  if (!isAdmin()) return targetUser.id === currentUser.id;
  if (isTagAdminUser(currentUser)) return true;

  const myCompany = normalizeCompanyName(currentUser.companyName);
  if (!myCompany) return targetUser.id === currentUser.id;
  return normalizeCompanyName(targetUser.companyName) === myCompany || targetUser.id === currentUser.id;
}

function can(action, stageKey = "") {
  if (!currentUser) return false;
  if (isDeveloper()) return true;
  if (isAdmin()) {
    if (action === "sync") return false;
    if (action === "report") return true;
    if (action === "toggleStage") return true;
    return [
      "manageUsers",
      "manageCatalog",
      "manageContainers",
      "manageContainerManifest",
      "manageContainerRelease",
      "manageContainerIssues",
      "dispatchTasks",
      "siteReceive",
      "deleteContainer",
      "deleteUnit",
      "createUnit",
      "quality",
      "install",
      "notes",
      "issues",
      "photos",
      "checklist",
    ].includes(action);
  }
  if (currentUser.role === "visitor") return action === "report";

  if (action === "manageUsers") return false;
  if (action === "sync") return false;
  if (action === "manageCatalog") return false;
  if (action === "manageContainers") return ["warehouse", "transport"].includes(currentUser.role);
  if (action === "manageContainerManifest") return ["warehouse", "transport"].includes(currentUser.role);
  if (action === "manageContainerRelease") return currentUser.role === "warehouse";
  if (action === "manageContainerIssues") return ["warehouse", "transport", "qa"].includes(currentUser.role);
  if (action === "dispatchTasks") return currentUser.role === "warehouse";
  if (action === "siteReceive") return ["transport", "qa", "installer"].includes(currentUser.role);
  if (action === "deleteContainer") return false;
  if (action === "deleteUnit") return false;
  if (action === "createUnit") return currentUser.role === "warehouse";
  if (action === "toggleStage") return STAGE_ROLE_ACCESS[currentUser.role]?.includes(stageKey) || false;
  if (action === "quality") return currentUser.role === "qa";
  if (action === "install") return currentUser.role === "installer";
  if (action === "notes") return currentUser.role === "qa" || currentUser.role === "installer";
  if (action === "issues") return currentUser.role === "qa" || currentUser.role === "installer";
  if (action === "photos") return currentUser.role === "qa" || currentUser.role === "installer";
  if (action === "checklist") return ["warehouse", "transport", "qa"].includes(currentUser.role);
  if (action === "report") return true;

  return false;
}

function rolePermissionsSummary(role) {
  if (role === "developer") return t("Acesso desenvolvedor: controle total, inclusive configuracoes de sistema.");
  if (role === "admin") return t("Acesso admin: gestao de dados, sem alterar estrutura do sistema.");
  if (role === "warehouse") return t("Cria unidades, controla containers, manifesto e envio/retencao do warehouse.");
  if (role === "transport") return t("Atualiza transportation/site job e acompanha schedule de containers.");
  if (role === "qa") return t("Valida qualidade, pendencias e pode anexar fotos.");
  if (role === "installer") return t("Controla distribuicao, instalacao, pendencias e fotos.");
  if (role === "visitor") return t("Acesso visitante: apenas consulta e relatorios. Aprovacao de funcoes por admin/developer.");
  return "";
}

function stageDoneCount(unit) {
  return STAGES.filter((stage) => unit.stages[stage.key]?.done).length;
}

function statusBadge(type) {
  if (type === "ok") return "OK";
  if (type === "missing") return "Faltante";
  if (type === "damaged") return "Defeito";
  if (type === "adjustment") return "Ajuste";
  return "-";
}

async function hashPassword(password) {
  const encoded = new TextEncoder().encode(password);
  const buffer = await crypto.subtle.digest("SHA-256", encoded);
  return Array.from(new Uint8Array(buffer))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function loadAppAuditLog() {
  try {
    const raw = localStorage.getItem(APP_AUDIT_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function persistAppAuditLog() {
  try {
    localStorage.setItem(APP_AUDIT_KEY, JSON.stringify(appAuditLog.slice(-APP_AUDIT_MAX)));
  } catch {
    // Ignore storage errors to avoid blocking user actions.
  }
}

function auditValue(value, max = 120) {
  const text = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return "-";
  return text.length > max ? `${text.slice(0, max)}...` : text;
}

function auditChange(field, beforeValue, afterValue) {
  return `${field}: "${auditValue(beforeValue)}" -> "${auditValue(afterValue)}"`;
}

function pushAppAudit(message, category = "app", scope = "-") {
  if (!currentUser || !message) return;
  appAuditLog.push({
    id: uid(),
    at: new Date().toISOString(),
    byUserId: currentUser.id,
    byName: currentUser.name || currentUser.username || "User",
    category,
    scope,
    message,
    view: currentView,
    sector: currentProjectSector,
  });
  if (appAuditLog.length > APP_AUDIT_MAX) appAuditLog = appAuditLog.slice(-APP_AUDIT_MAX);
  persistAppAuditLog();
}

function auditComparableValue(value) {
  if (Array.isArray(value)) return value.join(", ");
  if (value && typeof value === "object") {
    if ("name" in value) return value.name || "file";
    return JSON.stringify(value);
  }
  return String(value ?? "");
}

function collectAuditChanges(beforeObj, afterObj, fields) {
  return fields
    .map((field) => {
      const key = typeof field === "string" ? field : field.key;
      const label = typeof field === "string" ? field : field.label || key;
      const mapValue = typeof field === "string" ? null : field.map;
      const beforeRaw = mapValue ? mapValue(beforeObj?.[key], beforeObj) : beforeObj?.[key];
      const afterRaw = mapValue ? mapValue(afterObj?.[key], afterObj) : afterObj?.[key];
      const beforeValue = auditComparableValue(beforeRaw);
      const afterValue = auditComparableValue(afterRaw);
      if (beforeValue === afterValue) return "";
      return auditChange(label, beforeValue, afterValue);
    })
    .filter(Boolean);
}

function pushEntityAudit(entity, action, detail, scope = "-") {
  pushAppAudit(`${entity} ${action}: ${detail}`, "data-change", scope);
}

function pushContainerAudit(container, message) {
  container.auditLog = pushAudit(container.auditLog, message);
  const linkedClient = clientById(container.clientId);
  const linkedProject = projectById(container.projectId);
  const scope = [container.containerCode || container.id, linkedClient?.name, linkedProject?.name].filter(Boolean).join(" | ") || "-";
  pushAppAudit(`[Container ${container.containerCode || container.id}] ${message}`, "container-change", scope);
}

function sessionUserId() {
  return localStorage.getItem(SESSION_KEY);
}

function setSession(userId) {
  localStorage.setItem(SESSION_KEY, userId);
}

function clearSession() {
  localStorage.removeItem(SESSION_KEY);
}

async function tryRestoreSession() {
  const id = sessionUserId();
  if (!id) return;
  const found = users.find((user) => user.id === id);
  if (found) currentUser = found;
}

function showSignupMode(show) {
  if (!loginPanel || !signupPanel) return;
  loginPanel.classList.toggle("hidden", show);
  signupPanel.classList.toggle("hidden", !show);
}

function renderAuth() {
  const noUsers = users.length === 0;

  if (currentUser) {
    authView.classList.add("hidden");
    appMain.classList.remove("hidden");
    userBadge.classList.remove("hidden");
    editProfileBtn?.classList.remove("hidden");
    logoutBtn.classList.remove("hidden");
    userBadge.textContent = `${currentUser.name} (${roleLabel(currentUser.role)})`;
    currentView = viewFromHash();
    render();
    void maybeAutoDispatchCoiReminder();
    return;
  }

  appMain.classList.add("hidden");
  authView.classList.remove("hidden");
  userBadge.classList.add("hidden");
  editProfileBtn?.classList.add("hidden");
  logoutBtn.classList.add("hidden");

  setupPanel.classList.toggle("hidden", !noUsers);
  if (noUsers) {
    loginPanel.classList.add("hidden");
    signupPanel?.classList.add("hidden");
  } else {
    showSignupMode(false);
  }
  applyLanguageToUi();
}

function setFormEnabled(form, enabled) {
  form.querySelectorAll("input, select, textarea, button").forEach((field) => {
    field.disabled = !enabled;
  });
}

function viewFromHash() {
  const raw = window.location.hash.replace(/^#/, "").trim();
  if (!raw) return "home";
  const allowed = new Set(["home", "sync", "users", "clients", "projects", "userEdit"]);
  return allowed.has(raw) ? raw : "home";
}

function canOpenView(view) {
  if (!currentUser) return false;
  if (view === "home") return true;
  if (view === "sync") return can("sync");
  if (view === "users") return can("manageUsers");
  if (view === "userEdit") return true;
  if (view === "clients") return can("manageCatalog");
  if (view === "projects") return true;
  return false;
}

function allowedProjectSectorsForRole(role) {
  if (role === "developer" || role === "admin" || role === "visitor") return PROJECT_SECTORS;
  if (role === "warehouse") return ["warehouse"];
  if (role === "transport") return ["delivery"];
  if (role === "installer") return ["distribuicao", "instalacao"];
  if (role === "qa") return ["punchlist"];
  return ["delivery"];
}

function allowedProjectSectors() {
  return allowedProjectSectorsForRole(currentUser?.role || "visitor");
}

function ensureProjectSector() {
  const allowed = allowedProjectSectors();
  if (!allowed.includes(currentProjectSector)) currentProjectSector = allowed[0] || "delivery";
}

function stageVisibleForSector(stageKey) {
  if (currentView !== "projects") return true;
  const allowedStages = SECTOR_STAGE_MAP[currentProjectSector] || [];
  return allowedStages.includes(stageKey);
}

function nodeVisibleForSector(node) {
  const list = String(node?.dataset?.sectors || "")
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
  if (!list.length) return true;
  return list.includes(currentProjectSector);
}

function applySectorVisibility(root) {
  if (!root) return;
  root.querySelectorAll("[data-sectors]").forEach((node) => {
    node.classList.toggle("hidden-view", !nodeVisibleForSector(node));
  });
}

function applyMasterSubpanelMode() {
  if (!clientsSubpanel || !projectsSubpanel || !contactsSubpanel) return;
  const showClients = currentView === "clients";
  const showProjects = currentView === "projects";

  if (showClients) {
    clientsSubpanel.classList.remove("hidden-view");
    projectsSubpanel.classList.add("hidden-view");
    contactsSubpanel.classList.add("hidden-view");
    return;
  }

  if (showProjects) {
    clientsSubpanel.classList.add("hidden-view");
    const showProjectCatalog = projectsViewMode === "overview";
    projectsSubpanel.classList.toggle("hidden-view", !showProjectCatalog);
    contactsSubpanel.classList.toggle("hidden-view", !showProjectCatalog);
    return;
  }

  clientsSubpanel.classList.remove("hidden-view");
  projectsSubpanel.classList.remove("hidden-view");
  contactsSubpanel.classList.remove("hidden-view");
}

function renderProjectSectorPanel() {
  if (!projectSectorPanel) return;
  const allowed = allowedProjectSectors();
  projectSectorPanel.querySelectorAll("[data-project-sector]").forEach((button) => {
    const sector = button.dataset.projectSector;
    const enabled = allowed.includes(sector);
    button.classList.toggle("hidden", !enabled);
    button.classList.toggle("active", enabled && sector === currentProjectSector);
    button.disabled = !enabled;
  });
}

function applyProjectSectorLayout() {
  const inProjects = currentView === "projects";
  projectSectorPanel?.classList.toggle("hidden-view", !inProjects);
  if (!inProjects) return;

  const containerVisible = ["fabrica", "warehouse"].includes(currentProjectSector);
  const unitEntryVisible = ["warehouse"].includes(currentProjectSector);
  const unitsVisible = ["warehouse", "delivery", "distribuicao", "instalacao", "punchlist"].includes(currentProjectSector);

  containerPanel.classList.toggle("hidden-view", !containerVisible);
  unitEntryPanel.classList.toggle("hidden-view", !unitEntryVisible);
  unitsToolbarPanel.classList.toggle("hidden-view", !unitsVisible);
  qrLookupPanel.classList.toggle("hidden-view", !unitsVisible);
  unitsContainer.classList.toggle("hidden-view", !unitsVisible);
  applySectorVisibility(containerPanel);
}

function applyViewMode() {
  if (!currentUser) return;
  if (!canOpenView(currentView)) currentView = "home";
  if (currentView === "userEdit" && !editingUserId) currentView = can("manageUsers") ? "users" : "home";
  ensureProjectSector();
  renderProjectSectorPanel();

  const groups = {
    home: [homePanel],
    sync: [syncPanel],
    users: [usersPanel],
    userEdit: [userEditPanel],
    clients: [masterDataPanel],
    projects: [masterDataPanel, projectSectorPanel, containerPanel, unitEntryPanel, unitsToolbarPanel, qrLookupPanel, unitsContainer],
  };

  const allPanels = [
    homePanel,
    syncPanel,
    usersPanel,
    userEditPanel,
    masterDataPanel,
    projectSectorPanel,
    containerPanel,
    unitEntryPanel,
    unitsToolbarPanel,
    qrLookupPanel,
    unitsContainer,
  ];
  allPanels.forEach((panel) => panel?.classList.add("hidden-view"));
  (groups[currentView] || []).forEach((panel) => panel?.classList.remove("hidden-view"));

  applyMasterSubpanelMode();
  applyProjectSectorLayout();

  const showStats = currentView === "projects" && currentProjectSector === "warehouse";
  statsPanel.classList.toggle("hidden-view", !showStats);
  quickNavPanel?.classList.toggle("hidden-view", false);
}

function setView(view, { updateHash = true } = {}) {
  const previousView = currentView;
  currentView = view;
  applyViewMode();
  if (previousView !== currentView) {
    pushAppAudit(`Navigation view: ${previousView || "-"} -> ${currentView}`, "navigation", `sector:${currentProjectSector}`);
  }
  if (updateHash) {
    const hash = `#${currentView}`;
    if (window.location.hash !== hash) window.history.replaceState(null, "", hash);
  }
}

function setProjectSector(sector, { rerender = true } = {}) {
  if (!PROJECT_SECTORS.includes(sector)) return;
  if (!allowedProjectSectors().includes(sector)) return;
  const previousSector = currentProjectSector;
  currentProjectSector = sector;
  if (previousSector !== currentProjectSector) {
    pushAppAudit(
      `Navigation sector: ${previousSector || "-"} -> ${currentProjectSector}`,
      "navigation",
      `view:${currentView}`
    );
  }
  if (rerender && currentUser) render();
}

function usersFilterMatches(user) {
  if (usersViewFilter === "subcontractor") return user.employmentType === "subcontractor";
  if (usersViewFilter === "partner") return user.employmentType === "supplier";
  return true;
}

function setUsersViewFilter(filter = "all") {
  usersViewFilter = filter;
}

function openProjectsSector(sector) {
  if (!canOpenView("projects")) {
    setView("home");
    return;
  }
  projectsViewMode = "operations";
  setUsersViewFilter("all");
  setView("projects");
  setProjectSector(sector);
}

function openProjectReportsView() {
  if (!canOpenView("projects") || !can("report")) {
    setView("home");
    return;
  }
  projectsViewMode = "operations";
  setUsersViewFilter("all");
  setView("projects");
  const visibleSectors = ["delivery", "distribuicao", "instalacao", "punchlist"].filter((sector) =>
    allowedProjectSectors().includes(sector)
  );
  if (visibleSectors.length && !visibleSectors.includes(currentProjectSector)) {
    setProjectSector(visibleSectors[0]);
  }
  projectReportSelect?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function handleQuickNavAction(action) {
  if (action === "home") {
    setUsersViewFilter("all");
    setView("home");
    return;
  }

  if (action === "sync") {
    setUsersViewFilter("all");
    setView(can("sync") ? "sync" : "home");
    return;
  }

  if (action === "developer") {
    setUsersViewFilter("all");
    setView(isDeveloper() ? "sync" : "home");
    return;
  }

  if (action === "users") {
    setUsersViewFilter("all");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "clients") {
    setUsersViewFilter("all");
    setView(can("manageCatalog") ? "clients" : "home");
    return;
  }

  if (action === "projects") {
    projectsViewMode = "overview";
    setUsersViewFilter("all");
    setView(canOpenView("projects") ? "projects" : "home");
    return;
  }

  if (action === "subcontractors") {
    setUsersViewFilter("subcontractor");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "partners") {
    setUsersViewFilter("partner");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "warehouse") {
    openProjectsSector("warehouse");
    return;
  }

  if (action === "delivery") {
    openProjectsSector("delivery");
    return;
  }

  if (action === "distribution") {
    openProjectsSector("distribuicao");
    return;
  }

  if (action === "installation") {
    openProjectsSector("instalacao");
    return;
  }

  if (action === "changeOrders") {
    openProjectsSector("punchlist");
    return;
  }

  if (action === "projectReports") {
    openProjectReportsView();
    return;
  }
}

function renderStats() {
  const total = units.length;
  const completed = units.filter((unit) => unit.installationStatus === "completed").length;
  const pending = units.filter((unit) => stageDoneCount(unit) < STAGES.length).length;
  const blocked = units.filter((unit) => unit.installationStatus === "blocked").length;
  const qualityIssues = units.filter((unit) => unit.deliveryQuality === "rejected").length;

  const arriving7 = containers.filter((container) => {
    if (!container.etaDate || container.arrivalStatus === "arrived") return false;
    const d = daysUntil(container.etaDate);
    return d !== null && d >= 0 && d <= 7;
  }).length;

  const delayed = containers.filter((container) => {
    if (container.arrivalStatus === "delayed") return true;
    if (!container.etaDate || container.arrivalStatus === "arrived") return false;
    const d = daysUntil(container.etaDate);
    return d !== null && d < 0;
  }).length;

  const cards = [
    ["Total unidades", total],
    ["Pendentes", pending],
    ["Instalacao concluida", completed],
    ["Instalacao bloqueada", blocked],
    ["Qualidade rejeitada", qualityIssues],
    ["Containers", containers.length],
    ["Chegam em 7 dias", arriving7],
    ["Containers atrasados", delayed],
  ];

  statsPanel.innerHTML = cards
    .map(([label, value]) => `<div class="stat-item"><span>${label}</span><strong>${value}</strong></div>`)
    .join("");
}

function populateSelect(
  select,
  list,
  {
    placeholder = "Select",
    includeEmpty = true,
    labelFn = (item) => item.name,
  } = {}
) {
  const previous = select.value;
  const options = [];
  if (includeEmpty) options.push(`<option value="">${escapeHtml(t(placeholder))}</option>`);
  for (const item of list) options.push(`<option value="${escapeHtml(item.id)}">${escapeHtml(labelFn(item))}</option>`);

  select.innerHTML = options.join("");
  if (list.some((item) => item.id === previous)) select.value = previous;
}

function projectsForClient(clientId) {
  if (!clientId) return projects;
  return projects.filter((project) => project.clientId === clientId);
}

function populateProjectSelectByClient(select, clientId, placeholder, includeEmpty = true) {
  populateSelect(select, projectsForClient(clientId), {
    includeEmpty,
    placeholder,
    labelFn: (project) => `${project.name}${project.code ? ` (${project.code})` : ""}`,
  });
}

function syncUnitProjectInfo() {
  const projectId = unitProjectSelect.value;
  if (!projectId) {
    unitClientName.value = "";
    unitProjectName.value = "";
    unitJobSite.value = "";
    return;
  }

  const project = projectById(projectId);
  const client = project ? clientById(project.clientId) : null;
  unitProjectName.value = project?.name || "";
  unitClientName.value = client?.name || "";
  unitJobSite.value = project?.address || "";
}

function syncContainerProjectSelect() {
  const clientId = containerClientSelect.value;
  populateProjectSelectByClient(
    containerProjectSelect,
    clientId,
    clientId ? "Select a project" : "Select a client",
    true
  );
}

function syncContactProjectSelect() {
  const clientId = contactClientSelect.value;
  populateProjectSelectByClient(contactProjectSelect, clientId, "Project (optional)", true);
}

function renderMasterSelects() {
  populateSelect(projectClientSelect, clients, {
    includeEmpty: true,
    placeholder: clients.length ? "Select the client" : "Register a client",
  });

  populateSelect(contactClientSelect, clients, {
    includeEmpty: true,
    placeholder: "Client (optional)",
  });

  syncContactProjectSelect();

  populateSelect(unitProjectSelect, projects, {
    includeEmpty: true,
    placeholder: projects.length ? "Select a project" : "Register client and project first",
    labelFn: (project) => {
      const client = clientById(project.clientId);
      return `${project.name} - ${client?.name || "No client"}`;
    },
  });

  populateSelect(containerClientSelect, clients, {
    includeEmpty: true,
    placeholder: clients.length ? "Select the client" : "Cadastre cliente primeiro",
  });

  syncContainerProjectSelect();

  const previousReport = projectReportSelect.value;
  projectReportSelect.innerHTML = `<option value="all">Project for PDF (all)</option>${projects
    .map((project) => `<option value="${escapeHtml(project.id)}">${escapeHtml(project.name)}</option>`)
    .join("")}`;
  if (projects.some((project) => project.id === previousReport)) projectReportSelect.value = previousReport;

  syncUnitProjectInfo();
}

function renderClientsTable() {
  if (!clientsSummaryScroll || !clientDetailsPanel || !clientForm) return;

  const query = clientSearchQuery.trim().toLowerCase();
  const visibleClients = clients.filter((client) => {
    if (!query) return true;
    return String(client.name || "").toLowerCase().includes(query);
  });

  if (selectedClientId && !clients.some((entry) => entry.id === selectedClientId)) {
    selectedClientId = "";
  }
  if (selectedClientId && !visibleClients.some((entry) => entry.id === selectedClientId)) {
    selectedClientId = visibleClients[0]?.id || "";
  }
  if (!selectedClientId && visibleClients.length) selectedClientId = visibleClients[0].id;

  const cards = visibleClients
    .map((client) => {
      const officePhone = client.officePhone || client.phone || "-";
      return `<button class="client-summary-card${client.id === selectedClientId ? " active" : ""}" type="button" data-client-summary="${
        client.id
      }">
        <strong>${escapeHtml(client.name || "-")}</strong>
        <span>${escapeHtml(client.address || "No address")}</span>
        <small>Office: ${escapeHtml(officePhone)}</small>
      </button>`;
    })
    .join("");

  clientsSummaryScroll.innerHTML = cards || '<p class="hint">No matching clients.</p>';

  clientsSummaryScroll.querySelectorAll("[data-client-summary]").forEach((button) => {
    button.addEventListener("click", () => {
      selectedClientId = button.dataset.clientSummary || "";
      const selectedClient = clients.find((entry) => entry.id === selectedClientId) || null;
      renderClientsTable();
      if (selectedClient) populateClientForm(selectedClient);
    });
  });

  const selectedClient = clients.find((entry) => entry.id === selectedClientId) || null;
  renderClientDetails(selectedClient);

  if (!selectedClient) {
    resetClientForm();
    return;
  }

  const currentEditId = clientForm.elements.clientId?.value || "";
  if (!currentEditId) populateClientForm(selectedClient);
}

function splitLines(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function setClientLogoPreview(dataUrl) {
  if (!clientLogoPreview) return;
  if (dataUrl) {
    clientLogoPreview.src = dataUrl;
    clientLogoPreview.classList.remove("hidden");
  } else {
    clientLogoPreview.removeAttribute("src");
    clientLogoPreview.classList.add("hidden");
  }
}

function resetClientForm() {
  if (!clientForm) return;
  clientForm.reset();
  if (clientForm.elements.clientId) clientForm.elements.clientId.value = "";
  if (clientAdminTarget) clientAdminTarget.textContent = "Creating a new client.";
  setClientLogoPreview("");
}

function populateClientForm(client) {
  if (!clientForm || !client) return;
  if (clientForm.elements.clientId) clientForm.elements.clientId.value = client.id;
  clientForm.elements.name.value = client.name || "";
  clientForm.elements.address.value = client.address || "";
  clientForm.elements.officePhone.value = client.officePhone || client.phone || "";
  clientForm.elements.email.value = client.email || "";
  clientForm.elements.ownerName.value = client.ownerName || "";
  clientForm.elements.contactOwner.value = client.contactOwner || "";
  clientForm.elements.contactSeniorProjectManager.value = client.contactSeniorProjectManager || "";
  clientForm.elements.seniorProjectManagerPhone.value = client.seniorProjectManagerPhone || "";
  clientForm.elements.yearsInBusiness.value = client.yearsInBusiness || "";
  clientForm.elements.offices.value = client.offices || "";
  if (clientForm.elements.logo) clientForm.elements.logo.value = "";
  if (clientAdminTarget) clientAdminTarget.textContent = `Editing: ${client.name || "Client"}`;
  setClientLogoPreview(client.logoDataUrl || "");
}

function openClientRegistrationShortcut(target = "projects") {
  if (!canOpenView("projects")) {
    setView("home");
    return;
  }
  projectsViewMode = "overview";
  setView("projects");

  const selectedId = clientForm?.elements?.clientId?.value || selectedClientId || "";
  if (target === "projects") {
    if (selectedId && projectClientSelect) projectClientSelect.value = selectedId;
    projectsSubpanel?.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  if (selectedId && contactClientSelect) {
    contactClientSelect.value = selectedId;
    syncContactProjectSelect();
  }
  contactsSubpanel?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function renderClientDetails(client) {
  if (!clientDetailsPanel) return;
  if (!client) {
    clientDetailsPanel.innerHTML = '<p class="hint">Select a client to view company details.</p>';
    return;
  }

  const linkedProjects = projects.filter((project) => project.clientId === client.id);
  const officeLines = splitLines(client.offices);
  const logoBlock = client.logoDataUrl
    ? `<img class="client-logo" src="${escapeHtml(client.logoDataUrl)}" alt="${escapeHtml(client.name)} logo" />`
    : '<div class="client-logo placeholder">No logo</div>';
  const locationsHtml = officeLines.length
    ? `<ul>${officeLines.map((entry) => `<li>${escapeHtml(entry)}</li>`).join("")}</ul>`
    : "<p>-</p>";
  const projectsText = linkedProjects.length ? linkedProjects.map((project) => project.name).join(", ") : "-";

  clientDetailsPanel.innerHTML = `
    <div class="client-details-head">
      ${logoBlock}
      <div class="client-details-main">
        <h4>${escapeHtml(client.name || "-")}</h4>
        <p><strong>Address:</strong> ${escapeHtml(client.address || "-")}</p>
        <p><strong>Office Phone:</strong> ${escapeHtml(client.officePhone || client.phone || "-")}</p>
        <p><strong>General Email:</strong> ${escapeHtml(client.email || "-")}</p>
      </div>
    </div>
    <div class="client-details-grid">
      <div>
        <h5>Business Information</h5>
        <p><strong>Owner:</strong> ${escapeHtml(client.ownerName || "-")}</p>
        <p><strong>Senior Project Manager:</strong> ${escapeHtml(client.contactSeniorProjectManager || "-")}</p>
        <p><strong>Senior PM Phone:</strong> ${escapeHtml(client.seniorProjectManagerPhone || "-")}</p>
        <p><strong>Years in Business:</strong> ${escapeHtml(client.yearsInBusiness || "-")}</p>
        <p><strong>Projects Involved:</strong> ${escapeHtml(projectsText)}</p>
      </div>
      <div>
        <h5>Other Locations</h5>
        ${locationsHtml}
      </div>
    </div>
  `;
}

function renderProjectsTable() {
  const canManage = can("manageCatalog");

  const rows = projects
    .map((project) => {
      const client = clientById(project.clientId);
      return `<tr>
        <td>${escapeHtml(project.name)}</td>
        <td>${escapeHtml(client?.name || "-")}</td>
        <td>${escapeHtml(project.code)}</td>
        <td>${escapeHtml(project.address)}</td>
        <td>${
          canManage
            ? `<div class="actions-inline"><button class="secondary xs-btn" data-edit-project="${project.id}">Edit</button><button class="danger xs-btn" data-del-project="${project.id}">Delete</button></div>`
            : "-"
        }</td>
      </tr>`;
    })
    .join("");

  projectsTable.innerHTML = `<table class="data-table"><thead><tr><th>Project</th><th>Client</th><th>Code</th><th>Address</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="5">Nenhum projeto cadastrado.</td></tr>'
  }</tbody></table>`;

  if (!canManage) return;

  projectsTable.querySelectorAll("[data-edit-project]").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = button.dataset.editProject;
      const project = projects.find((entry) => entry.id === id);
      if (!project) return;
      const name = prompt("Project name:", project.name)?.trim();
      if (!name) return;
      const code = prompt("Codigo do projeto:", project.code || "")?.trim() || "";
      const address = prompt("Endereco / Job Site:", project.address || "")?.trim() || "";
      const updatedProject = normalizeProject({
        ...project,
        name,
        code,
        address,
        updatedAt: new Date().toISOString(),
      });
      await put(PROJECT_STORE, updatedProject);
      const changes = collectAuditChanges(project, updatedProject, [
        { key: "name", label: "Project name" },
        { key: "code", label: "Project code" },
        { key: "address", label: "Project address" },
      ]);
      pushEntityAudit(
        "Projects",
        "updated",
        `${updatedProject.name}${changes.length ? ` | ${changes.join("; ")}` : " | no field changes"}`,
        "projects"
      );
      await loadAll();
      render();
      queueAutoSync();
    });
  });

  projectsTable.querySelectorAll("[data-del-project]").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = button.dataset.delProject;
      if (units.some((unit) => unit.projectId === id)) {
        alert(t("Nao e possivel excluir projeto com unidades vinculadas."));
        return;
      }
      if (contacts.some((contact) => contact.projectId === id)) {
        alert(t("Nao e possivel excluir projeto com pessoas vinculadas."));
        return;
      }
      if (containers.some((container) => container.projectId === id)) {
        alert(t("Nao e possivel excluir projeto com containers vinculados."));
        return;
      }
      const targetProject = projects.find((entry) => entry.id === id);
      if (targetProject) pushEntityAudit("Projects", "deleted", `${targetProject.name}`, "projects");
      await del(PROJECT_STORE, id);
      await loadAll();
      render();
      queueAutoSync();
    });
  });
}

function renderContactsTable() {
  const canManage = can("manageCatalog");

  const rows = contacts
    .map((contact) => {
      const client = clientById(contact.clientId);
      const project = projectById(contact.projectId);
      return `<tr>
        <td>${escapeHtml(contact.name)}</td>
        <td>${escapeHtml(contactRoleLabel(contact.role))}</td>
        <td>${escapeHtml(contact.company)}</td>
        <td>${escapeHtml(client?.name || "-")}</td>
        <td>${escapeHtml(project?.name || "-")}</td>
        <td>${escapeHtml(contact.phone)}</td>
        <td>${escapeHtml(contact.email)}</td>
        <td>${
          canManage
            ? `<div class="actions-inline"><button class="secondary xs-btn" data-edit-contact="${contact.id}">Edit</button><button class="danger xs-btn" data-del-contact="${contact.id}">Delete</button></div>`
            : "-"
        }</td>
      </tr>`;
    })
    .join("");

  contactsTable.innerHTML = `<table class="data-table"><thead><tr><th>Name</th><th>Role</th><th>Company</th><th>Client</th><th>Project</th><th>Phone</th><th>Email</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="8">Nenhuma pessoa cadastrada.</td></tr>'
  }</tbody></table>`;

  if (!canManage) return;
  contactsTable.querySelectorAll("[data-edit-contact]").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = button.dataset.editContact;
      const contact = contacts.find((entry) => entry.id === id);
      if (!contact) return;
      const name = prompt("Name:", contact.name)?.trim();
      if (!name) return;
      const company = prompt("Company:", contact.company || "")?.trim() || "";
      const phone = prompt("Phone:", contact.phone || "")?.trim() || "";
      const email = prompt("Email:", contact.email || "")?.trim() || "";
      const updatedContact = normalizeContact({
        ...contact,
        name,
        company,
        phone,
        email,
        updatedAt: new Date().toISOString(),
      });
      await put(CONTACT_STORE, updatedContact);
      const changes = collectAuditChanges(contact, updatedContact, [
        { key: "name", label: "Contact name" },
        { key: "company", label: "Company" },
        { key: "phone", label: "Phone" },
        { key: "email", label: "Email" },
      ]);
      pushEntityAudit(
        "Contacts",
        "updated",
        `${updatedContact.name}${changes.length ? ` | ${changes.join("; ")}` : " | no field changes"}`,
        "contacts"
      );
      await loadAll();
      render();
      queueAutoSync();
    });
  });

  contactsTable.querySelectorAll("[data-del-contact]").forEach((button) => {
    button.addEventListener("click", async () => {
      const targetContact = contacts.find((entry) => entry.id === button.dataset.delContact);
      if (targetContact) pushEntityAudit("Contacts", "deleted", `${targetContact.name}`, "contacts");
      await del(CONTACT_STORE, button.dataset.delContact);
      await loadAll();
      render();
      queueAutoSync();
    });
  });
}

function renderMaterialsTable() {
  const canManage = can("manageCatalog");
  const rows = materials
    .map(
      (material) => `<tr>
      <td>${escapeHtml(material.sku)}</td>
      <td>${escapeHtml(material.description)}</td>
      <td>${escapeHtml(material.category)}</td>
      <td>${escapeHtml(material.unit)}</td>
      <td>${escapeHtml(material.kitchenType || "-")}</td>
      <td>${escapeHtml(fmtDate(material.updatedAt))}</td>
      <td>${
        canManage
          ? `<div class="actions-inline"><button class="secondary xs-btn" data-edit-material="${material.id}">Edit</button><button class="danger xs-btn" data-del-material="${material.id}">Delete</button></div>`
          : "-"
      }</td>
    </tr>`
    )
    .join("");

  materialsTable.innerHTML = `<table class="data-table"><thead><tr><th>SKU</th><th>Description</th><th>Category</th><th>Unit</th><th>Kitchen Type</th><th>Updated</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="7">No materials registered.</td></tr>'
  }</tbody></table>`;

  if (!canManage) return;

  materialsTable.querySelectorAll("[data-edit-material]").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = button.dataset.editMaterial;
      const material = materials.find((entry) => entry.id === id);
      if (!material) return;

      const sku = prompt("SKU/Codigo:", material.sku)?.trim();
      if (!sku) return;
      const description = prompt("Description:", material.description)?.trim();
      if (!description) return;
      const category = prompt("Category (kitchen, vanity, med-cabinet, countertop, other):", material.category)?.trim() || "other";
      const unit = prompt("Unit:", material.unit || "pcs")?.trim() || "pcs";
      const kitchenType = prompt("Tipo de cozinha:", material.kitchenType || "")?.trim() || "";

      const updatedMaterial = normalizeMaterial({
        ...material,
        sku,
        description,
        category,
        unit,
        kitchenType,
        updatedAt: new Date().toISOString(),
      });
      await put(MATERIAL_STORE, updatedMaterial);
      const changes = collectAuditChanges(material, updatedMaterial, [
        { key: "sku", label: "SKU" },
        { key: "description", label: "Description" },
        { key: "category", label: "Category" },
        { key: "unit", label: "Unit" },
        { key: "kitchenType", label: "Kitchen type" },
      ]);
      pushEntityAudit(
        "Materials",
        "updated",
        `${updatedMaterial.sku}${changes.length ? ` | ${changes.join("; ")}` : " | no field changes"}`,
        "materials"
      );
      await loadAll();
      render();
      queueAutoSync();
    });
  });

  materialsTable.querySelectorAll("[data-del-material]").forEach((button) => {
    button.addEventListener("click", async () => {
      const id = button.dataset.delMaterial;
      const linked = containers.some((container) => container.materialItems.some((item) => item.materialId === id));
      if (linked) {
        alert("Cannot delete material already used in a container manifest.");
        return;
      }
      const targetMaterial = materials.find((entry) => entry.id === id);
      if (targetMaterial) pushEntityAudit("Materials", "deleted", `${targetMaterial.sku} | ${targetMaterial.description}`, "materials");
      await del(MATERIAL_STORE, id);
      await loadAll();
      render();
      queueAutoSync();
    });
  });
}

function renderMasterData() {
  const canManage = can("manageCatalog");
  masterDataPanel.classList.toggle("hidden", !canManage);

  setFormEnabled(clientForm, canManage);
  setFormEnabled(projectForm, canManage);
  setFormEnabled(contactForm, canManage);
  setFormEnabled(materialForm, canManage);

  renderMasterSelects();
  if (!canManage) return;
  renderClientsTable();
  renderProjectsTable();
  renderContactsTable();
  renderMaterialsTable();
}

function containerReleasedQty(container) {
  const totals = { kitchen: 0, vanity: 0, medCabinet: 0, countertop: 0 };
  for (const move of container.movements) {
    if (move.action !== "release") continue;
    totals.kitchen += toNumber(move.kitchens);
    totals.vanity += toNumber(move.vanities);
    totals.medCabinet += toNumber(move.medCabinets);
    totals.countertop += toNumber(move.countertops);
  }
  return totals;
}

function containerAvailableQty(container) {
  const released = containerReleasedQty(container);
  return {
    kitchen: container.qtyKitchens - released.kitchen,
    vanity: container.qtyVanities - released.vanity,
    medCabinet: container.qtyMedCabinets - released.medCabinet,
    countertop: container.qtyCountertops - released.countertop,
  };
}

function containerAlertLine(container) {
  if (!container.etaDate) return "No ETA provided.";
  const d = daysUntil(container.etaDate);
  if (container.arrivalStatus === "arrived") return `Container received with status Arrived. ETA: ${fmtDateOnly(container.etaDate)}.`;
  if (container.arrivalStatus === "delayed" || d < 0) return `Warning: delayed container (${Math.abs(d)} day(s) late).`;
  if (d === 0) return "Container expected today.";
  if (d <= 7) return `Container arrives in ${d} day(s).`;
  return `ETA in ${d} day(s).`;
}

async function saveContainer(container) {
  container.updatedAt = new Date().toISOString();
  await put(CONTAINER_STORE, normalizeContainer(container));
  await loadAll();
  render();
  queueAutoSync();
}

function transactionDonePromise(transaction) {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error("Transaction failed"));
    transaction.onabort = () => reject(transaction.error || new Error("Transaction aborted"));
  });
}

async function saveUnitAndContainer(unit, container, unitAuditMessage = "") {
  const now = new Date().toISOString();
  if (unitAuditMessage) unit.auditLog = pushAudit(unit.auditLog, unitAuditMessage);
  if (unitAuditMessage) {
    pushAppAudit(`[Unit ${unit.unitCode || unit.id}] ${unitAuditMessage}`, "unit-change", unit.projectName || "-");
  }
  unit.updatedAt = now;
  container.updatedAt = now;
  const transaction = db.transaction([UNIT_STORE, CONTAINER_STORE], "readwrite");
  transaction.objectStore(UNIT_STORE).put(normalizeUnit(unit));
  transaction.objectStore(CONTAINER_STORE).put(normalizeContainer(container));
  await transactionDonePromise(transaction);
  await loadAll();
  render();
  queueAutoSync();
}

function renderManifestTable(wrapper, container, editable) {
  const rows = container.materialItems
    .map(
      (item) => `<tr>
      <td>${escapeHtml(item.code)}</td>
      <td>${escapeHtml(item.description)}</td>
      <td>${escapeHtml(item.category)}</td>
      <td>${escapeHtml(item.qty)}</td>
      <td>${escapeHtml(item.unit)}</td>
      <td>
        ${
          editable
            ? `<select data-manifest-field="issueType" data-manifest-id="${item.id}">
            <option value="ok" ${item.issueType === "ok" ? "selected" : ""}>OK</option>
            <option value="missing" ${item.issueType === "missing" ? "selected" : ""}>Missing</option>
            <option value="damaged" ${item.issueType === "damaged" ? "selected" : ""}>Damage</option>
            <option value="misplaced" ${item.issueType === "misplaced" ? "selected" : ""}>Misplaced</option>
          </select>`
            : escapeHtml(item.issueType || "ok")
        }
      </td>
      <td>
        ${
          editable
            ? `<input data-manifest-field="issueRoute" data-manifest-id="${item.id}" value="${escapeHtml(item.issueRoute || "")}" placeholder="Route to..." />`
            : escapeHtml(item.issueRoute || "-")
        }
      </td>
      <td>
        ${
          editable
            ? `<input data-manifest-field="issueNote" data-manifest-id="${item.id}" value="${escapeHtml(item.issueNote || "")}" placeholder="Issue note" />`
            : escapeHtml(item.issueNote || "-")
        }
      </td>
      <td>${editable ? `<button class="danger xs-btn" data-manifest-del="${item.id}">x</button>` : "-"}</td>
    </tr>`
    )
    .join("");

  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>Code</th><th>Description</th><th>Category</th><th>Qty</th><th>Unit</th><th>Issue</th><th>Route To</th><th>Detail</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="9">No manifest items.</td></tr>'
  }</tbody></table>`;
}

function renderReleaseTable(wrapper, container) {
  const rows = container.movements
    .slice()
    .reverse()
    .map((move) => {
      return `<tr>
      <td>${escapeHtml(fmtDate(move.createdAt))}</td>
      <td>${escapeHtml(move.action === "release" ? "Dispatch" : "Hold")}</td>
      <td>${escapeHtml(move.unitLabel || "-")}</td>
      <td>${escapeHtml(move.destination || "-")}</td>
      <td>K:${toNumber(move.kitchens)} V:${toNumber(move.vanities)} M:${toNumber(move.medCabinets)} C:${toNumber(move.countertops)}</td>
      <td>${escapeHtml(move.operatorName || "-")}</td>
      <td>${escapeHtml(move.note || "-")}</td>
    </tr>`;
    })
    .join("");

  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>Date</th><th>Movement</th><th>Unit</th><th>Destination</th><th>Qty</th><th>Operator</th><th>Note</th></tr></thead><tbody>${
    rows || '<tr><td colspan="7">No movement records.</td></tr>'
  }</tbody></table>`;
}

function renderContainers() {
  containerPanel.classList.toggle("hidden", !currentUser);
  containersBoard.innerHTML = "";

  if (!containers.length) {
    containersBoard.innerHTML = `<div class="panel empty">No containers registered.</div>`;
    return;
  }

  for (const container of containers) {
    const node = containerTemplate.content.firstElementChild.cloneNode(true);
    const client = clientById(container.clientId);
    const project = projectById(container.projectId);
    const available = containerAvailableQty(container);
    const released = containerReleasedQty(container);

    node.querySelector(".container-title").textContent = `${container.containerCode} - ${containerStatusLabel(container.arrivalStatus)}`;
    node.querySelector(".container-meta").textContent = [
      client?.name && `Client: ${client.name}`,
      project?.name && `Project: ${project.name}`,
      container.supplier && `Supplier: ${container.supplier}`,
      container.manufacturer && `Manufacturer: ${container.manufacturer}`,
      container.etaDate && `ETA: ${fmtDateOnly(container.etaDate)}`,
      container.departureAt && `Port departure: ${fmtDate(container.departureAt)}`,
    ]
      .filter(Boolean)
      .join(" | ");

    node.querySelector(".container-kpi").innerHTML = `
      <div class="kpi-box"><span>Planned (K/V/M/C)</span><strong>${container.qtyKitchens}/${container.qtyVanities}/${container.qtyMedCabinets}/${container.qtyCountertops}</strong></div>
      <div class="kpi-box"><span>Dispatched (K/V/M/C)</span><strong>${released.kitchen}/${released.vanity}/${released.medCabinet}/${released.countertop}</strong></div>
      <div class="kpi-box"><span>Balance (K/V/M/C)</span><strong>${available.kitchen}/${available.vanity}/${available.medCabinet}/${available.countertop}</strong></div>
      <div class="kpi-box"><span>Counter since port departure</span><strong>${elapsedFrom(container.departureAt)}</strong></div>
    `;

    const alertLine = containerAlertLine(container);
    node.querySelector(".container-alert").textContent = container.loosePartsNotes
      ? `${alertLine} | Loose parts: ${container.loosePartsNotes}`
      : alertLine;

    const deleteBtn = node.querySelector(".container-delete-btn");
    if (can("manageContainers")) {
      const editBtn = document.createElement("button");
      editBtn.type = "button";
      editBtn.className = "secondary";
      editBtn.textContent = "Edit";
      deleteBtn.parentElement.appendChild(editBtn);
      editBtn.addEventListener("click", async () => {
        const prevSupplier = container.supplier || "";
        const prevManufacturer = container.manufacturer || "";
        const prevNotes = container.notes || "";
        const prevLoose = container.loosePartsNotes || "";
        const supplier = prompt("Supplier:", container.supplier || "")?.trim();
        if (supplier === null) return;
        const manufacturer = prompt("Manufacturer:", container.manufacturer || "")?.trim();
        if (manufacturer === null) return;
        const notes = prompt("Notes:", container.notes || "")?.trim();
        if (notes === null) return;
        const loosePartsNotes = prompt("Loose parts:", container.loosePartsNotes || "")?.trim();
        if (loosePartsNotes === null) return;
        container.supplier = supplier;
        container.manufacturer = manufacturer;
        container.notes = notes;
        container.loosePartsNotes = loosePartsNotes;
        const changes = [
          prevSupplier !== supplier ? `supplier "${auditValue(prevSupplier)}" -> "${auditValue(supplier)}"` : "",
          prevManufacturer !== manufacturer
            ? `manufacturer "${auditValue(prevManufacturer)}" -> "${auditValue(manufacturer)}"`
            : "",
          prevNotes !== notes ? `notes "${auditValue(prevNotes)}" -> "${auditValue(notes)}"` : "",
          prevLoose !== loosePartsNotes
            ? `loose parts "${auditValue(prevLoose)}" -> "${auditValue(loosePartsNotes)}"`
            : "",
        ].filter(Boolean);
        pushContainerAudit(
          container,
          changes.length ? `Container updated (${container.containerCode}): ${changes.join("; ")}` : `Container updated (${container.containerCode}) with no field changes`
        );
        await saveContainer(container);
      });
    }
    deleteBtn.classList.toggle("hidden", !can("deleteContainer"));
    deleteBtn.addEventListener("click", async () => {
      if (!can("deleteContainer")) return;
      pushContainerAudit(container, `Container deleted: ${container.containerCode}`);
      await del(CONTAINER_STORE, container.id);
      await loadAll();
      render();
      queueAutoSync();
    });

    const manifestEditable = can("manageContainerManifest");
    const manifestForm = node.querySelector(".manifest-form");
    const manifestMaterialSelect = manifestForm.querySelector(".mf-material");
    const manifestCodeInput = manifestForm.querySelector(".mf-code");
    const manifestDescInput = manifestForm.querySelector(".mf-desc");
    const manifestCategorySelect = manifestForm.querySelector(".mf-category");
    const manifestQtyInput = manifestForm.querySelector(".mf-qty");
    const manifestUnitInput = manifestForm.querySelector(".mf-unit");

    manifestMaterialSelect.innerHTML = `<option value="">Catalog (optional)</option>${materials
      .map((material) => `<option value="${escapeHtml(material.id)}">${escapeHtml(material.sku)} - ${escapeHtml(material.description)}</option>`)
      .join("")}`;

    manifestMaterialSelect.addEventListener("change", () => {
      const selected = materialById(manifestMaterialSelect.value);
      if (!selected) return;
      manifestCodeInput.value = selected.sku;
      manifestDescInput.value = selected.description;
      manifestCategorySelect.value = selected.category || "other";
      manifestUnitInput.value = selected.unit || "pcs";
    });

    if (!manifestEditable) manifestForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));

    const manifestTable = node.querySelector(".manifest-table");
    renderManifestTable(manifestTable, container, manifestEditable);

    manifestForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!manifestEditable) return;

      const materialId = manifestMaterialSelect.value;
      const selectedMaterial = materialById(materialId);
      const code = manifestCodeInput.value.trim() || selectedMaterial?.sku || "";
      const description = manifestDescInput.value.trim() || selectedMaterial?.description || "";
      const category = manifestCategorySelect.value || selectedMaterial?.category || "other";
      const qty = Math.trunc(toNumber(manifestQtyInput.value));
      const unit = manifestUnitInput.value.trim() || selectedMaterial?.unit || "pcs";

      if (!description || qty <= 0) {
        alert(t("Descricao e quantidade sao obrigatorias."));
        return;
      }

      const now = new Date().toISOString();
      const manifestItemId = uid();
      const resolvedCode = code || `ITEM-${String(container.materialItems.length + 1).padStart(3, "0")}`;
      container.materialItems.push({
        id: manifestItemId,
        materialId: materialId || "",
        code: resolvedCode,
        description,
        category,
        qty,
        unit,
        issueType: "ok",
        issueNote: "",
        issueRoute: "",
        createdAt: now,
        updatedAt: now,
      });

      for (let index = 0; index < qty; index += 1) {
        container.qrItems.push({
          id: uid(),
          manifestItemId,
          qrCode: makeQrCode(container, resolvedCode || "ITEM"),
          materialId: materialId || "",
          code: resolvedCode,
          description,
          category,
          unit,
          status: "in-warehouse",
          clientId: "",
          projectId: "",
          unitId: "",
          unitLabel: "",
          kitchenType: selectedMaterial?.kitchenType || "",
          assignedAt: "",
          assignedBy: "",
          deliveredAt: "",
          issueType: "",
          issueNote: "",
          createdAt: now,
          updatedAt: now,
        });
      }

      pushContainerAudit(container, `Manifest item added: ${description} (${qty} ${unit || "un"})`);
      pushContainerAudit(container, `QR generated for ${resolvedCode}: ${qty} unit(s)`);

      manifestForm.reset();
      await saveContainer(container);
    });

    if (manifestEditable) {
      manifestTable.querySelectorAll("[data-manifest-field]").forEach((field) => {
        field.addEventListener("change", async () => {
          const item = container.materialItems.find((entry) => entry.id === field.dataset.manifestId);
          if (!item) return;
          const key = field.dataset.manifestField;
          if (key === "issueNote" || key === "issueRoute") return;
          const previous = item[key] || "";
          const next = field.value;
          if (previous === next) return;
          item[key] = field.value;
          item.updatedAt = new Date().toISOString();
          pushContainerAudit(
            container,
            `Manifest item ${item.description || item.code || item.id} updated: ${key} "${auditValue(previous)}" -> "${auditValue(next)}"`
          );
          await saveContainer(container);
        });
      });

      manifestTable.querySelectorAll('[data-manifest-field="issueNote"],[data-manifest-field="issueRoute"]').forEach((field) => {
        field.addEventListener("blur", async () => {
          const item = container.materialItems.find((entry) => entry.id === field.dataset.manifestId);
          if (!item) return;
          const key = field.dataset.manifestField;
          const previous = item[key] || "";
          const next = field.value.trim();
          if (previous === next) return;
          item[key] = next;
          item.updatedAt = new Date().toISOString();
          pushContainerAudit(
            container,
            `Manifest item ${item.description || item.code || item.id} updated: ${key} "${auditValue(previous)}" -> "${auditValue(next)}"`
          );
          await saveContainer(container);
        });
      });

      manifestTable.querySelectorAll("[data-manifest-del]").forEach((btn) => {
        btn.addEventListener("click", async () => {
          const removed = container.materialItems.find((item) => item.id === btn.dataset.manifestDel);
          container.materialItems = container.materialItems.filter((item) => item.id !== btn.dataset.manifestDel);
          container.qrItems = container.qrItems.filter(
            (item) => item.manifestItemId !== btn.dataset.manifestDel || (item.status !== "in-warehouse" && item.status !== "dispatched")
          );
          if (removed) pushContainerAudit(container, `Manifest item removed: ${removed.description || removed.code}`);
          await saveContainer(container);
        });
      });
    }

    const releaseEditable = can("manageContainerRelease");
    const releaseForm = node.querySelector(".release-form");
    const releaseUnitSelect = releaseForm.querySelector(".rf-unit");
    const releaseDestination = releaseForm.querySelector(".rf-destination");

    const unitOptions = units.filter((u) => u.projectId === container.projectId || u.projectName === project?.name);
    releaseUnitSelect.innerHTML = `<option value="">Destination unit (optional)</option>${unitOptions
      .map((u) => `<option value="${escapeHtml(u.id)}">${escapeHtml(u.unitCode)} - ${escapeHtml(u.jobSite || u.projectName)}</option>`)
      .join("")}`;

    releaseUnitSelect.addEventListener("change", () => {
      const selected = unitOptions.find((u) => u.id === releaseUnitSelect.value);
      if (selected) releaseDestination.value = selected.jobSite || selected.unitCode;
    });

    if (!releaseEditable) releaseForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));

    const releaseTable = node.querySelector(".release-table");
    renderReleaseTable(releaseTable, container);

    releaseForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!releaseEditable) return;

      const unitId = releaseUnitSelect.value;
      const selectedUnit = units.find((u) => u.id === unitId);
      const action = releaseForm.querySelector(".rf-action").value;
      const destination = releaseDestination.value.trim();
      const kitchens = toNumber(releaseForm.querySelector(".rf-k").value);
      const vanities = toNumber(releaseForm.querySelector(".rf-v").value);
      const medCabinets = toNumber(releaseForm.querySelector(".rf-m").value);
      const countertops = toNumber(releaseForm.querySelector(".rf-c").value);
      const note = releaseForm.querySelector(".rf-note").value.trim();

      const totalMoved = kitchens + vanities + medCabinets + countertops;
      if (totalMoved <= 0) {
        alert(t("Informe ao menos uma quantidade para registrar movimento."));
        return;
      }

      if (action === "release") {
        const balance = containerAvailableQty(container);
        if (
          kitchens > balance.kitchen ||
          vanities > balance.vanity ||
          medCabinets > balance.medCabinet ||
          countertops > balance.countertop
        ) {
          alert(t("Quantidade excede o saldo disponivel no container."));
          return;
        }
      }

      const now = new Date().toISOString();
      container.movements.push({
        id: uid(),
        unitId: unitId || "",
        unitLabel: selectedUnit ? selectedUnit.unitCode : "",
        action,
        destination,
        kitchens,
        vanities,
        medCabinets,
        countertops,
        note,
        operatorUserId: currentUser.id,
        operatorName: currentUser.name,
        createdAt: now,
      });
      pushContainerAudit(
        container,
        `${action === "release" ? "Dispatch" : "Hold"} recorded: K${kitchens} V${vanities} M${medCabinets} C${countertops} -> ${
          destination || "no destination"
        }`
      );

      container.updatedAt = now;
      await put(CONTAINER_STORE, normalizeContainer(container));

      if (action === "release" && selectedUnit) {
        selectedUnit.stages.warehouse = { done: true, at: now };
        const noteLine = `Warehouse release ${container.containerCode}: K${kitchens} V${vanities} M${medCabinets} C${countertops} (${fmtDate(now)})`;
        selectedUnit.deliveryNotes = selectedUnit.deliveryNotes
          ? `${selectedUnit.deliveryNotes}\n${noteLine}`
          : noteLine;
        selectedUnit.updatedAt = now;
        await put(UNIT_STORE, normalizeUnit(selectedUnit));
      }

      releaseForm.reset();
      releaseUnitSelect.value = "";
      releaseDestination.value = "";
      await loadAll();
      render();
      queueAutoSync();
    });

    const qrDispatchForm = node.querySelector(".qr-dispatch-form");
    const qrUnitSelect = qrDispatchForm.querySelector(".qd-unit");
    const qrKitchenInput = qrDispatchForm.querySelector(".qd-kitchen");
    const qrScanInput = qrDispatchForm.querySelector(".qd-scan");
    const qrTrackTable = node.querySelector(".qr-track-table");
    const replaceQueueTable = node.querySelector(".replace-queue-table");
    const factoryQueueTable = node.querySelector(".factory-queue-table");
    renderQrTrackTable(qrTrackTable, container);
    qrTrackTable.querySelectorAll("[data-qr-print]").forEach((button) => {
      button.addEventListener("click", () => {
        const item = container.qrItems.find((entry) => entry.id === button.dataset.qrPrint);
        if (!item) return;
        openQrLabelWindow(container, item, { autoPrint: true });
      });
    });
    qrTrackTable.querySelectorAll("[data-qr-png]").forEach((button) => {
      button.addEventListener("click", () => {
        const item = container.qrItems.find((entry) => entry.id === button.dataset.qrPng);
        if (!item) return;
        downloadQrLabelPng(item);
      });
    });
    renderQueueTable(replaceQueueTable, container.replacementQueue);
    renderQueueTable(factoryQueueTable, container.factoryMissingQueue);

    qrUnitSelect.innerHTML = `<option value="">Required unit</option>${unitOptions
      .map((u) => `<option value="${escapeHtml(u.id)}">${escapeHtml(u.unitCode)} - ${escapeHtml(u.jobSite || u.projectName)}</option>`)
      .join("")}`;

    const qrDispatchEditable = can("manageContainerRelease");
    if (!qrDispatchEditable) qrDispatchForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));

    qrDispatchForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!qrDispatchEditable) return;
      const qrCode = qrScanInput.value.trim();
      const unitId = qrUnitSelect.value;
      const targetUnit = unitOptions.find((entry) => entry.id === unitId);
      if (!qrCode || !targetUnit) {
        alert("Select a unit and scan the QR.");
        return;
      }
      const qrItem = container.qrItems.find((entry) => entry.qrCode === qrCode);
      if (!qrItem) {
        alert("QR nao encontrado neste container.");
        return;
      }
      const now = new Date().toISOString();
      qrItem.clientId = container.clientId;
      qrItem.projectId = container.projectId;
      qrItem.unitId = targetUnit.id;
      qrItem.unitLabel = targetUnit.unitCode;
      qrItem.kitchenType = qrKitchenInput.value.trim() || qrItem.kitchenType || targetUnit.unitType || targetUnit.kitchenModel || "";
      qrItem.assignedAt = now;
      qrItem.assignedBy = currentUser.name;
      qrItem.status = "dispatched";
      qrItem.updatedAt = now;
      pushContainerAudit(container, `QR routed: ${qrItem.qrCode} -> ${targetUnit.unitCode}`);
      qrDispatchForm.reset();
      await saveContainer(container);
    });

    applySectorVisibility(node);
    containersBoard.appendChild(node);
  }
}

function matchesFilter(unit, query, filter) {
  const q = query.trim().toLowerCase();
  const text = [
    unit.clientName,
    unit.projectName,
    unit.jobSite,
    unit.unitCode,
    unit.building,
    unit.floor,
    unit.category,
    unit.unitType,
    unit.kitchenModel,
  ]
    .join(" ")
    .toLowerCase();

  const passQuery = !q || text.includes(q);
  if (!passQuery) return false;

  if (filter === "pending") return stageDoneCount(unit) < STAGES.length;
  if (filter === "qualityIssues") return unit.deliveryQuality === "rejected";
  if (filter === "blockedInstall") return unit.installationStatus === "blocked";
  return true;
}

async function saveSetting(row) {
  await put(SETTINGS_STORE, row);
  await loadAll();
  render();
}

function updateSyncStatus(message, isError = false) {
  syncStatus.textContent = message;
  syncStatus.style.borderColor = isError ? "#e7b6b6" : "#94c5a8";
  syncStatus.style.background = isError ? "#fff5f5" : "#f4fcf7";
}

async function saveUnit(unit) {
  unit.updatedAt = new Date().toISOString();
  await put(UNIT_STORE, normalizeUnit(unit));
  await loadAll();
  render();
  queueAutoSync();
}

async function saveUnitWithAudit(unit, message) {
  if (message) unit.auditLog = pushAudit(unit.auditLog, message);
  if (message) pushAppAudit(`[Unit ${unit.unitCode || unit.id}] ${message}`, "unit-change", unit.projectName || "-");
  await saveUnit(unit);
}

async function addPhotos(unitId, files) {
  const fileNames = [];
  for (const file of files) {
    const dataUrl = await fileToDataUrl(file);
    fileNames.push(file.name || "photo");
    await put(PHOTO_STORE, {
      id: uid(),
      unitId,
      name: file.name,
      type: file.type,
      dataUrl,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
  }
  const targetUnit = units.find((entry) => entry.id === unitId);
  if (fileNames.length) {
    pushAppAudit(
      `[Unit ${targetUnit?.unitCode || unitId}] Photos uploaded (${fileNames.length}): ${fileNames.join(", ")}`,
      "unit-change",
      targetUnit?.projectName || "-"
    );
  }
  await loadAll();
  render();
  queueAutoSync();
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function checklistSummary(checkItems) {
  const total = checkItems.length;
  const missing = checkItems.filter((item) => item.status === "missing").length;
  const damaged = checkItems.filter((item) => item.status === "damaged").length;
  const adjustment = checkItems.filter((item) => item.status === "adjustment").length;
  return `Items: ${total} | Missing: ${missing} | Damaged: ${damaged} | Adjustment: ${adjustment}`;
}

function renderChecklistTable(wrapper, unit) {
  const editable = can("checklist");
  if (!unit.checkItems.length) {
    wrapper.innerHTML = `<p class="hint">No items registered.</p>`;
    return;
  }

  wrapper.innerHTML = `
    <table class="data-table">
      <thead>
        <tr>
          <th>Code</th>
          <th>Description</th>
          <th>Expected</th>
          <th>Checked</th>
          <th>Status</th>
          <th>Notes</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        ${unit.checkItems
          .map(
            (item) => `
            <tr>
              <td>${escapeHtml(item.code)}</td>
              <td>${escapeHtml(item.description)}</td>
              <td>${escapeHtml(item.expectedQty)}</td>
              <td><input data-item-action="got" data-item-id="${item.id}" type="number" min="0" value="${Number(item.checkedQty || 0)}" ${
                editable ? "" : "disabled"
              } /></td>
              <td>
                <select data-item-action="status" data-item-id="${item.id}" ${editable ? "" : "disabled"}>
                  <option value="ok" ${item.status === "ok" ? "selected" : ""}>OK</option>
                  <option value="missing" ${item.status === "missing" ? "selected" : ""}>Missing</option>
                  <option value="damaged" ${item.status === "damaged" ? "selected" : ""}>Damaged</option>
                  <option value="adjustment" ${item.status === "adjustment" ? "selected" : ""}>Adjustment</option>
                </select>
              </td>
              <td><input data-item-action="note" data-item-id="${item.id}" value="${escapeHtml(item.note || "")}" ${
                editable ? "" : "disabled"
              } /></td>
              <td><button data-item-remove="${item.id}" class="danger xs-btn" ${editable ? "" : "disabled"}>x</button></td>
            </tr>`
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function bindChecklistEvents(node, unit) {
  const editable = can("checklist");
  const wrapper = node.querySelector(".checklist-table");
  renderChecklistTable(wrapper, unit);
  node.querySelector(".checklist-summary").textContent = checklistSummary(unit.checkItems);

  if (!editable) return;

  wrapper.querySelectorAll("[data-item-action]").forEach((el) => {
    el.addEventListener("change", () => {
      const item = unit.checkItems.find((entry) => entry.id === el.dataset.itemId);
      if (!item) return;
      if (el.dataset.itemAction === "got") {
        const previous = Number(item.checkedQty || 0);
        const next = Number(el.value || 0);
        if (previous === next) return;
        item.checkedQty = next;
        item.updatedAt = new Date().toISOString();
        saveUnitWithAudit(unit, `Checklist "${item.description || item.code}" checked qty: ${previous} -> ${next}`);
        return;
      }
      if (el.dataset.itemAction === "status") {
        const previous = item.status || "ok";
        const next = el.value;
        if (previous === next) return;
        item.status = next;
        item.updatedAt = new Date().toISOString();
        saveUnitWithAudit(unit, `Checklist "${item.description || item.code}" status: ${previous} -> ${next}`);
        return;
      }
      if (el.dataset.itemAction === "note") {
        const previous = item.note || "";
        const next = el.value;
        if (previous === next) return;
        item.note = next;
        item.updatedAt = new Date().toISOString();
        saveUnitWithAudit(unit, auditChange(`Checklist "${item.description || item.code}" note`, previous, next));
        return;
      }
      item.updatedAt = new Date().toISOString();
    });
  });

  wrapper.querySelectorAll("[data-item-remove]").forEach((button) => {
    button.addEventListener("click", () => {
      unit.checkItems = unit.checkItems.filter((entry) => entry.id !== button.dataset.itemRemove);
      saveUnitWithAudit(unit, `Item removido do checklist: ${button.dataset.itemRemove}`);
    });
  });
}

function renderDispatchTable(wrapper, unit, editable) {
  const rows = unit.dispatchTasks
    .map(
      (task) => `<tr>
      <td>${escapeHtml(task.item)}</td>
      <td>${toNumber(task.qty)}</td>
      <td>${escapeHtml(task.createdBy || "-")}</td>
      <td>${escapeHtml(fmtDate(task.createdAt))}</td>
      <td>${
        editable
          ? `<div class="actions-inline"><button class="secondary xs-btn" data-dispatch-edit="${task.id}">Edit</button><button class="danger xs-btn" data-dispatch-del="${task.id}">x</button></div>`
          : "-"
      }</td>
    </tr>`
    )
    .join("");

  const signatureLine = unit.dispatchSignature
    ? `Signed by ${escapeHtml(unit.dispatchSignature.byName)} at ${escapeHtml(fmtDate(unit.dispatchSignature.at))}`
    : "No dispatch signature";

  wrapper.innerHTML = `<p class="hint">${signatureLine}</p><table class="data-table"><thead><tr><th>Item</th><th>Qty</th><th>Logged by</th><th>Date</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="5">No dispatch tasks.</td></tr>'
  }</tbody></table>`;
}

function renderSiteReceiveTable(wrapper, unit, editable) {
  const rows = unit.siteReceipts
    .slice()
    .reverse()
    .map(
      (entry) => `<tr>
      <td>${escapeHtml(entry.qrCode || "-")}</td>
      <td>${escapeHtml(entry.item)}</td>
      <td>${toNumber(entry.qty)}</td>
      <td>${escapeHtml(entry.issue || "ok")}</td>
      <td>${escapeHtml(entry.note || "-")}</td>
      <td>${escapeHtml(entry.receivedBy || "-")}</td>
      <td>${escapeHtml(fmtDate(entry.createdAt))}</td>
      <td>${
        editable
          ? `<div class="actions-inline"><button class="secondary xs-btn" data-receive-edit="${entry.id}">Edit</button><button class="danger xs-btn" data-receive-del="${entry.id}">x</button></div>`
          : "-"
      }</td>
    </tr>`
    )
    .join("");

  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>QR</th><th>Item</th><th>Qty</th><th>Issue</th><th>Detail</th><th>Received by</th><th>Date</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="8">No receipts recorded.</td></tr>'
  }</tbody></table>`;
}

function renderAuditTable(wrapper, unit) {
  const rows = unit.auditLog
    .slice()
    .reverse()
    .slice(0, 120)
    .map(
      (entry) => `<tr>
      <td>${escapeHtml(fmtDate(entry.at))}</td>
      <td>${escapeHtml(entry.byName || "-")}</td>
      <td>${escapeHtml(entry.message || "")}</td>
    </tr>`
    )
    .join("");

  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>Date</th><th>User</th><th>Action</th></tr></thead><tbody>${
    rows || '<tr><td colspan="3">No audited changes.</td></tr>'
  }</tbody></table>`;
}

function renderUnits() {
  const query = searchInput.value;
  const filter = filterSelect.value;
  const filtered = units.filter((unit) => matchesFilter(unit, query, filter));

  unitsContainer.innerHTML = "";

  if (!filtered.length) {
    unitsContainer.innerHTML = `<div class="panel empty">Nenhum registro encontrado.</div>`;
    return;
  }

  for (const unit of filtered) {
    const node = unitTemplate.content.firstElementChild.cloneNode(true);
    const progress = Math.round((stageDoneCount(unit) / STAGES.length) * 100);

    node.querySelector(".unit-title").textContent = `${unit.unitCode} - ${unit.category}`;
    node.querySelector(".unit-meta").textContent = [
      unit.clientName && `Client: ${unit.clientName}`,
      unit.projectName && `Project: ${unit.projectName}`,
      unit.jobSite,
      unit.building && `Building: ${unit.building}`,
      unit.floor && `Floor: ${unit.floor}`,
      unit.unitType && `Type: ${unit.unitType}`,
      unit.kitchenModel && `Kitchen: ${unit.kitchenModel}`,
    ]
      .filter(Boolean)
      .join(" | ");

    node.querySelector(".progress-bar").style.width = `${progress}%`;
    node.querySelector(".progress-label").textContent =
      currentLang === "en"
        ? `${progress}% completed (${stageDoneCount(unit)}/${STAGES.length})`
        : currentLang === "es"
        ? `${progress}% completado (${stageDoneCount(unit)}/${STAGES.length})`
        : `${progress}% concluido (${stageDoneCount(unit)}/${STAGES.length})`;

    const stageGrid = node.querySelector(".stage-grid");
    for (const stage of STAGES) {
      if (!stageVisibleForSector(stage.key)) continue;
      const stageValue = unit.stages[stage.key] || { done: false, at: null };
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `stage-btn ${stageValue.done ? "done" : ""}`;
      btn.disabled = !can("toggleStage", stage.key);
      btn.innerHTML = `<strong>${stageLabel(stage.key)}</strong><span>${
        stageValue.done
          ? currentLang === "en"
            ? `Checked: ${fmtDate(stageValue.at)}`
            : currentLang === "es"
            ? `Bajado: ${fmtDate(stageValue.at)}`
            : `Baixado: ${fmtDate(stageValue.at)}`
          : t("Pendente")
      }</span>`;
      btn.addEventListener("click", () => {
        if (!can("toggleStage", stage.key)) return;
        const latest = unit.stages[stage.key];
        unit.stages[stage.key] = {
          done: !latest.done,
          at: !latest.done ? new Date().toISOString() : null,
        };
        saveUnitWithAudit(unit, `Etapa ${stageLabel(stage.key)} ${unit.stages[stage.key].done ? "concluida" : "reaberta"}`);
      });
      stageGrid.appendChild(btn);
    }

    const qualitySelect = node.querySelector(".quality-status");
    const installSelect = node.querySelector(".install-status");
    const notesInput = node.querySelector(".delivery-notes");
    const issuesInput = node.querySelector(".issues-text");

    qualitySelect.value = unit.deliveryQuality;
    installSelect.value = unit.installationStatus;
    notesInput.value = unit.deliveryNotes;
    issuesInput.value = unit.issuesText;

    qualitySelect.disabled = !can("quality");
    installSelect.disabled = !can("install");
    notesInput.disabled = !can("notes");
    issuesInput.disabled = !can("issues");

    qualitySelect.addEventListener("change", () => {
      if (!can("quality")) return;
      const previous = unit.deliveryQuality;
      const next = qualitySelect.value;
      if (previous === next) return;
      unit.deliveryQuality = qualitySelect.value;
      saveUnitWithAudit(unit, auditChange("Delivery quality", previous, next));
    });

    installSelect.addEventListener("change", () => {
      if (!can("install")) return;
      const previous = unit.installationStatus;
      const next = installSelect.value;
      if (previous === next) return;
      unit.installationStatus = installSelect.value;
      saveUnitWithAudit(unit, auditChange("Installation status", previous, next));
    });

    notesInput.addEventListener("blur", () => {
      if (!can("notes")) return;
      const previous = unit.deliveryNotes || "";
      const next = notesInput.value.trim();
      if (previous === next) return;
      unit.deliveryNotes = next;
      saveUnitWithAudit(unit, auditChange("Quality/installation notes", previous, next));
    });

    issuesInput.addEventListener("blur", () => {
      if (!can("issues")) return;
      const previous = unit.issuesText || "";
      const next = issuesInput.value.trim();
      if (previous === next) return;
      unit.issuesText = next;
      saveUnitWithAudit(unit, auditChange("Issues text", previous, next));
    });

    const dispatchForm = node.querySelector(".dispatch-form");
    const dispatchTable = node.querySelector(".dispatch-table");
    const signDispatchBtn = node.querySelector(".sign-dispatch-btn");
    const canDispatch = can("dispatchTasks");
    if (!canDispatch) dispatchForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));
    renderDispatchTable(dispatchTable, unit, canDispatch);
    signDispatchBtn.disabled = !canDispatch;
    signDispatchBtn.textContent = unit.dispatchSignature ? "Reassinar envio (warehouse)" : "Assinar envio (warehouse)";

    dispatchForm.addEventListener("submit", (event) => {
      event.preventDefault();
      if (!canDispatch) return;
      const item = dispatchForm.querySelector(".df-item").value.trim();
      const qty = toNumber(dispatchForm.querySelector(".df-qty").value);
      if (!item || qty <= 0) return;
      unit.dispatchTasks.push({
        id: uid(),
        item,
        qty,
        createdByUserId: currentUser.id,
        createdBy: currentUser.name,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      dispatchForm.reset();
      saveUnitWithAudit(unit, `Tarefa de envio adicionada: ${item} (${qty})`);
    });

    if (canDispatch) {
      dispatchTable.querySelectorAll("[data-dispatch-edit]").forEach((button) => {
        button.addEventListener("click", () => {
          const row = unit.dispatchTasks.find((entry) => entry.id === button.dataset.dispatchEdit);
          if (!row) return;
          const previousItem = row.item;
          const previousQty = row.qty;
          const item = prompt("Item do delivery:", row.item)?.trim();
          if (!item) return;
          const qty = toNumber(prompt("Quantidade:", String(row.qty)) || row.qty);
          if (previousItem === item && Number(previousQty) === Number(qty)) return;
          row.item = item;
          row.qty = qty;
          row.updatedAt = new Date().toISOString();
          saveUnitWithAudit(
            unit,
            `Dispatch task edited: item "${auditValue(previousItem)}" -> "${auditValue(item)}"; qty ${previousQty} -> ${qty}`
          );
        });
      });

      dispatchTable.querySelectorAll("[data-dispatch-del]").forEach((button) => {
        button.addEventListener("click", () => {
          const row = unit.dispatchTasks.find((entry) => entry.id === button.dataset.dispatchDel);
          unit.dispatchTasks = unit.dispatchTasks.filter((entry) => entry.id !== button.dataset.dispatchDel);
          saveUnitWithAudit(unit, `Tarefa de envio removida: ${row?.item || button.dataset.dispatchDel}`);
        });
      });
    }

    signDispatchBtn.addEventListener("click", () => {
      if (!canDispatch) return;
      if (!unit.dispatchTasks.length) {
        alert("Add at least one task before signing.");
        return;
      }
      unit.dispatchSignature = {
        byUserId: currentUser.id,
        byName: currentUser.name,
        at: new Date().toISOString(),
      };
      saveUnitWithAudit(unit, "Assinatura digital de envio registrada no warehouse");
    });

    const siteForm = node.querySelector(".site-receive-form");
    const siteTable = node.querySelector(".site-receive-table");
    const canSiteReceive = can("siteReceive");
    if (!canSiteReceive) siteForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));
    renderSiteReceiveTable(siteTable, unit, canSiteReceive);

    siteForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!canSiteReceive) return;
      const qrCode = siteForm.querySelector(".sr-qr").value.trim();
      const item = siteForm.querySelector(".sr-item").value.trim();
      const qty = toNumber(siteForm.querySelector(".sr-qty").value);
      const issue = siteForm.querySelector(".sr-issue").value;
      const note = siteForm.querySelector(".sr-note").value.trim();

      if (qrCode) {
        const qrContainer = findQrItemContainer(qrCode, unit.projectId);
        if (!qrContainer) {
          alert("QR nao encontrado para este projeto.");
          return;
        }
        const qrItem = qrContainer.qrItems.find((entry) => entry.qrCode === qrCode);
        if (!qrItem) {
          alert("QR nao encontrado.");
          return;
        }

        const now = new Date().toISOString();
        const effectiveItem = item || qrItem.description || qrItem.code || qrCode;
        const effectiveQty = qty > 0 ? qty : 1;
        unit.siteReceipts.push({
          id: uid(),
          qrCode,
          item: effectiveItem,
          qty: effectiveQty,
          issue,
          note,
          receivedByUserId: currentUser.id,
          receivedBy: currentUser.name,
          createdAt: now,
          updatedAt: now,
        });

        qrItem.clientId = unit.clientId || qrContainer.clientId;
        qrItem.projectId = unit.projectId || qrContainer.projectId;
        qrItem.unitId = unit.id;
        qrItem.unitLabel = unit.unitCode;
        qrItem.kitchenType = qrItem.kitchenType || unit.unitType || unit.kitchenModel || "";
        qrItem.assignedAt = qrItem.assignedAt || now;
        qrItem.assignedBy = qrItem.assignedBy || currentUser.name;
        qrItem.deliveredAt = now;
        qrItem.issueType = issue === "ok" ? "" : issue;
        qrItem.issueNote = issue === "ok" ? "" : note;
        qrItem.status = issue === "ok" ? "received" : "issue";
        qrItem.updatedAt = now;

        if (issue !== "ok") {
          const issueLine = `${effectiveItem} (${effectiveQty}) -> ${issue}${note ? ` | ${note}` : ""}`;
          unit.issuesText = unit.issuesText ? `${unit.issuesText}\n${issueLine}` : issueLine;
          const queueEntry = {
            id: uid(),
            createdAt: now,
            qrCode: qrItem.qrCode,
            code: qrItem.code,
            description: qrItem.description || effectiveItem,
            unitId: unit.id,
            unitLabel: unit.unitCode,
            issueType: issue,
            note,
            kitchenType: qrItem.kitchenType || "",
            clientId: qrContainer.clientId,
            projectId: qrContainer.projectId,
            byUserId: currentUser.id,
            byName: currentUser.name,
          };
          if (qrIssueRoutesToFactory(issue)) {
            qrContainer.factoryMissingQueue.push(queueEntry);
          } else {
            qrContainer.replacementQueue.push(queueEntry);
          }
        }

        qrContainer.auditLog = pushAudit(
          qrContainer.auditLog,
          `QR receipt ${qrItem.qrCode} at unit ${unit.unitCode} with status ${issue}${note ? ` (${note})` : ""}`
        );

        siteForm.reset();
        await saveUnitAndContainer(unit, qrContainer, `Site job receipt by QR: ${effectiveItem} (${effectiveQty}) - ${issue}`);
        return;
      }

      if (!item || qty <= 0) return;
      unit.siteReceipts.push({
        id: uid(),
        qrCode: "",
        item,
        qty,
        issue,
        note,
        receivedByUserId: currentUser.id,
        receivedBy: currentUser.name,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
      siteForm.reset();
      if (issue !== "ok") {
        const issueLine = `${item} (${qty}) -> ${issue}${note ? ` | ${note}` : ""}`;
        unit.issuesText = unit.issuesText ? `${unit.issuesText}\n${issueLine}` : issueLine;
      }
      await saveUnitWithAudit(unit, `Site job receipt: ${item} (${qty}) - ${issue}`);
    });

    if (canSiteReceive) {
      siteTable.querySelectorAll("[data-receive-edit]").forEach((button) => {
        button.addEventListener("click", () => {
          const row = unit.siteReceipts.find((entry) => entry.id === button.dataset.receiveEdit);
          if (!row) return;
          const previousItem = row.item;
          const previousQty = row.qty;
          const previousIssue = row.issue;
          const previousNote = row.note || "";
          const item = prompt("Received item:", row.item)?.trim();
          if (!item) return;
          const qty = toNumber(prompt("Quantity:", String(row.qty)) || row.qty);
          const issue = prompt("Issue (ok, nao-chegou, quebrado, faltando-pecas, medida-diferente, cor-errada):", row.issue || "ok")?.trim() || row.issue;
          const note = prompt("Detail:", row.note || "")?.trim() || "";
          if (
            previousItem === item &&
            Number(previousQty) === Number(qty) &&
            previousIssue === issue &&
            previousNote === note
          ) {
            return;
          }
          row.item = item;
          row.qty = qty;
          row.issue = issue;
          row.note = note;
          row.updatedAt = new Date().toISOString();
          saveUnitWithAudit(
            unit,
            `Site receipt edited: item "${auditValue(previousItem)}" -> "${auditValue(item)}"; qty ${previousQty} -> ${qty}; issue ${previousIssue} -> ${issue}; note "${auditValue(
              previousNote
            )}" -> "${auditValue(note)}"`
          );
        });
      });

      siteTable.querySelectorAll("[data-receive-del]").forEach((button) => {
        button.addEventListener("click", () => {
          const row = unit.siteReceipts.find((entry) => entry.id === button.dataset.receiveDel);
          unit.siteReceipts = unit.siteReceipts.filter((entry) => entry.id !== button.dataset.receiveDel);
          saveUnitWithAudit(unit, `Receipt removed: ${row?.item || button.dataset.receiveDel}`);
        });
      });
    }

    const auditTable = node.querySelector(".audit-table");
    if (auditTable) renderAuditTable(auditTable, unit);

    const checkForm = node.querySelector(".check-item-form");
    if (!can("checklist")) checkForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));

    checkForm.addEventListener("submit", (event) => {
      event.preventDefault();
      if (!can("checklist")) return;

      const code = checkForm.querySelector(".ci-code").value.trim();
      const description = checkForm.querySelector(".ci-desc").value.trim();
      const expectedQty = Number(checkForm.querySelector(".ci-exp").value || 0);
      const checkedQty = Number(checkForm.querySelector(".ci-got").value || 0);
      const status = checkForm.querySelector(".ci-status").value;

      if (!description) return;

      unit.checkItems.push({
        id: uid(),
        code,
        description,
        expectedQty,
        checkedQty,
        status,
        note: "",
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });

      checkForm.reset();
      saveUnitWithAudit(unit, `Item adicionado ao checklist: ${description}`);
    });

    bindChecklistEvents(node, unit);

    node.querySelector(".scope-line").textContent = unit.scopeWork ? `Scope: ${unit.scopeWork}` : "Scope: nao informado";
    node.querySelector(".shop-line").textContent = unit.shopdrawingRef
      ? `Shopdrawing: ${unit.shopdrawingRef}`
      : "Shopdrawing: nao informado";

    const photoInput = node.querySelector(".photo-input");
    const photoLabel = node.querySelector(".photo-label");
    photoLabel.classList.toggle("hidden", !can("photos"));

    photoInput.addEventListener("change", async () => {
      if (!can("photos")) return;
      if (!photoInput.files?.length) return;
      await addPhotos(unit.id, photoInput.files);
      photoInput.value = "";
    });

    const gallery = node.querySelector(".photo-gallery");
    const unitPhotos = photosByUnit.get(unit.id) || [];
    for (const photo of unitPhotos) {
      const p = document.createElement("div");
      p.className = "photo";
      p.innerHTML = `<img loading="lazy" src="${photo.dataUrl}" alt="${escapeHtml(photo.name)}" /><button type="button" title="Remover foto" ${
        can("photos") ? "" : "disabled"
      }>x</button>`;
      p.querySelector("button").addEventListener("click", async () => {
        if (!can("photos")) return;
        pushAppAudit(
          `[Unit ${unit.unitCode || unit.id}] Photo removed: ${photo.name || photo.id}`,
          "unit-change",
          unit.projectName || "-"
        );
        await del(PHOTO_STORE, photo.id);
        await loadAll();
        render();
        queueAutoSync();
      });
      gallery.appendChild(p);
    }

    const deleteBtn = node.querySelector(".delete-btn");
    deleteBtn.classList.toggle("hidden", !can("deleteUnit"));
    deleteBtn.addEventListener("click", async () => {
      if (!can("deleteUnit")) return;
      pushAppAudit(`[Unit ${unit.unitCode || unit.id}] Unit deleted`, "unit-change", unit.projectName || "-");
      await del(UNIT_STORE, unit.id);
      const toDelete = photosByUnit.get(unit.id) || [];
      await Promise.all(toDelete.map((photo) => del(PHOTO_STORE, photo.id)));
      await loadAll();
      render();
      queueAutoSync();
    });

    node.querySelector(".report-btn").addEventListener("click", () => {
      generateUnitReport(unit.id);
    });

    applySectorVisibility(node);
    unitsContainer.appendChild(node);
  }
}

function resetAdminUserForm() {
  if (!userForm) return;
  adminEditingUserId = "";
  userForm.reset();
  userForm.role.value = "visitor";
  userForm.employmentType.value = "tag";
  toggleSubcontractorExtras("admin", "tag");
  if (userAdminTarget) userAdminTarget.textContent = "Creating a new user.";
  if (userAdminPhotoPreview) {
    userAdminPhotoPreview.removeAttribute("src");
    userAdminPhotoPreview.classList.add("hidden");
  }
  if (userAdminCoiFileStatus) userAdminCoiFileStatus.textContent = "No COI file attached.";
  if (openUserAdminCoiFileBtn) openUserAdminCoiFileBtn.classList.add("hidden");
  if (userFormDeleteBtn) userFormDeleteBtn.disabled = true;
}

function populateAdminUserForm(user) {
  if (!userForm || !user) return;
  adminEditingUserId = user.id;
  userForm.firstName.value = user.firstName || "";
  userForm.lastName.value = user.lastName || "";
  userForm.companyName.value = user.companyName || "";
  userForm.jobTitle.value = user.jobTitle || "";
  userForm.employmentType.value = user.employmentType || "tag";
  userForm.contractorCoi.value = user.contractorCoi || "";
  userForm.contractorCoiExpiry.value = user.contractorCoiExpiry || "";
  userForm.contractorW9.value = user.contractorW9 || "";
  setCheckedValues(userForm, "contractorAreas", user.contractorAreas || []);
  toggleSubcontractorExtras("admin", user.employmentType || "tag");
  userForm.address.value = user.address || "";
  userForm.phone.value = maskPhoneValue(user.phone || "");
  userForm.cellPhone.value = maskPhoneValue(user.cellPhone || "");
  userForm.email.value = user.email || "";
  userForm.username.value = user.username || "";
  userForm.password.value = "";
  userForm.role.value = user.role || "visitor";
  const photoInput = userForm.querySelector('input[name="photo"]');
  if (photoInput) photoInput.value = "";
  const contractorFileInput = userForm.querySelector('input[name="contractorCoiFile"]');
  if (contractorFileInput) contractorFileInput.value = "";
  if (userAdminTarget) userAdminTarget.textContent = `Editing user: ${user.name || user.username}`;
  if (userAdminPhotoPreview) {
    if (user.photoDataUrl) {
      userAdminPhotoPreview.src = user.photoDataUrl;
      userAdminPhotoPreview.classList.remove("hidden");
    } else {
      userAdminPhotoPreview.removeAttribute("src");
      userAdminPhotoPreview.classList.add("hidden");
    }
  }
  if (userAdminCoiFileStatus) {
    userAdminCoiFileStatus.textContent = user.contractorCoiFile?.name
      ? `Current file: ${user.contractorCoiFile.name}`
      : "No COI file attached.";
  }
  if (openUserAdminCoiFileBtn) openUserAdminCoiFileBtn.classList.toggle("hidden", !user.contractorCoiFile?.dataUrl);
  if (userFormDeleteBtn) userFormDeleteBtn.disabled = user.id === currentUser?.id;
}

function syncSelectedUserRow() {
  if (!usersTable) return;
  usersTable.querySelectorAll("tr.user-selected").forEach((row) => row.classList.remove("user-selected"));
  usersNameRail?.querySelectorAll(".users-name-btn.active").forEach((button) => button.classList.remove("active"));
  if (!adminEditingUserId) return;
  const selectedRow = usersTable.querySelector(`[data-user-row="${adminEditingUserId}"]`);
  if (selectedRow) selectedRow.classList.add("user-selected");
  const selectedName = usersNameRail?.querySelector(`[data-user-rail="${adminEditingUserId}"]`);
  if (selectedName) selectedName.classList.add("active");
}

function selectAdminUser(userId) {
  const target = users.find((entry) => entry.id === userId);
  if (!target) return;
  populateAdminUserForm(target);
  syncSelectedUserRow();
}

function renderUsers() {
  if (!can("manageUsers")) {
    usersPanel.classList.add("hidden");
    return;
  }

  usersPanel.classList.remove("hidden");

  const visibleUsers = users.filter((user) => {
    if (!canAccessEmployeeRecord(user)) return false;
    if (isDeveloper()) return usersFilterMatches(user);
    if (user.role === "developer" || isPrimaryDeveloperUser(user)) return false;
    return usersFilterMatches(user);
  });

  const userRoleSelect = userForm?.querySelector('select[name="role"]');
  if (userRoleSelect) {
    const developerOption = userRoleSelect.querySelector('option[value="developer"]');
    if (developerOption) developerOption.hidden = !isDeveloper();
    if (!isDeveloper() && userRoleSelect.value === "developer") userRoleSelect.value = "admin";
  }

  if (!adminEditingUserId && userForm && !userForm.dataset.ready) {
    resetAdminUserForm();
    userForm.dataset.ready = "1";
  }

  if (usersNameRail) {
    const nameRows = visibleUsers
      .map(
        (user) => `<button class="users-name-btn" type="button" data-user-rail="${user.id}">
          ${escapeHtml(user.name || user.username || "-")}
          <small>${escapeHtml(user.role || "visitor")} • ${escapeHtml(user.companyName || "No company")}</small>
        </button>`
      )
      .join("");
    usersNameRail.innerHTML = nameRows || '<p class="hint">No users available.</p>';
    usersNameRail.querySelectorAll("[data-user-rail]").forEach((button) => {
      button.addEventListener("click", () => {
        selectAdminUser(button.dataset.userRail || "");
      });
    });
  }

  const rows = visibleUsers
    .map(
      (user) => `<tr data-user-row="${user.id}">
      <td>${user.photoDataUrl ? `<img class="avatar-xs" src="${user.photoDataUrl}" alt="${escapeHtml(user.name)}" />` : "-"}</td>
      <td>${escapeHtml(user.name)}</td>
      <td>${escapeHtml(user.companyName || "-")}</td>
      <td>${escapeHtml(user.jobTitle || "-")}</td>
      <td>${escapeHtml(user.employmentType || "-")}</td>
      <td>${escapeHtml(user.email || "-")}</td>
      <td>${escapeHtml(user.username)}</td>
      <td>${escapeHtml(roleLabel(user.role || "visitor"))}</td>
      <td>${fmtDate(user.updatedAt)}</td>
    </tr>`
    )
    .join("");

  usersTable.innerHTML = `<table class="data-table"><thead><tr><th>Photo</th><th>Name</th><th>Company</th><th>Role</th><th>Type</th><th>Email</th><th>User</th><th>Profile</th><th>Updated</th></tr></thead><tbody>${
    rows || '<tr><td colspan="9">No visible users.</td></tr>'
  }</tbody></table>`;

  usersTable.querySelectorAll("[data-user-row]").forEach((row) => {
    row.addEventListener("click", () => {
      selectAdminUser(row.dataset.userRow || "");
    });
    row.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      selectAdminUser(row.dataset.userRow || "");
    });
    row.tabIndex = 0;
  });
  if (adminEditingUserId) {
    const selected = users.find((entry) => entry.id === adminEditingUserId);
    if (selected) populateAdminUserForm(selected);
    else resetAdminUserForm();
  }
  syncSelectedUserRow();
}

function populateUserEditForm(user) {
  if (!userEditForm || !user) return;
  userEditForm.firstName.value = user.firstName || "";
  userEditForm.lastName.value = user.lastName || "";
  userEditForm.companyName.value = user.companyName || "";
  userEditForm.jobTitle.value = user.jobTitle || "";
  userEditForm.employmentType.value = user.employmentType || "tag";
  userEditForm.contractorCoi.value = user.contractorCoi || "";
  userEditForm.contractorCoiExpiry.value = user.contractorCoiExpiry || "";
  userEditForm.contractorW9.value = user.contractorW9 || "";
  setCheckedValues(userEditForm, "contractorAreas", user.contractorAreas || []);
  toggleSubcontractorExtras("edit", user.employmentType || "tag");
  userEditForm.address.value = user.address || "";
  userEditForm.phone.value = maskPhoneValue(user.phone || "");
  userEditForm.cellPhone.value = maskPhoneValue(user.cellPhone || "");
  userEditForm.email.value = user.email || "";
  const photoInput = userEditForm.querySelector('input[name="photo"]');
  if (photoInput) photoInput.value = "";
  const contractorFileInput = userEditForm.querySelector('input[name="contractorCoiFile"]');
  if (contractorFileInput) contractorFileInput.value = "";
  if (userEditTarget) userEditTarget.textContent = `Editando: ${user.name || user.username}`;
  if (userEditCoiFileStatus) {
    if (user.contractorCoiFile?.name) {
      userEditCoiFileStatus.textContent = `Current file: ${user.contractorCoiFile.name}`;
    } else {
      userEditCoiFileStatus.textContent = "No COI file attached.";
    }
  }
  if (openUserEditCoiFileBtn) {
    openUserEditCoiFileBtn.classList.toggle("hidden", !user.contractorCoiFile?.dataUrl);
  }
  if (userEditPhotoPreview) {
    if (user.photoDataUrl) {
      userEditPhotoPreview.src = user.photoDataUrl;
      userEditPhotoPreview.classList.remove("hidden");
    } else {
      userEditPhotoPreview.removeAttribute("src");
      userEditPhotoPreview.classList.add("hidden");
    }
  }
}

function openUserEdit(userId, returnView = "users") {
  if (!currentUser) return;
  if (currentUser.role === "warehouse" && userId !== currentUser.id) return;
  if (!isDeveloper() && !isAdmin() && userId !== currentUser.id) return;
  const user = users.find((entry) => entry.id === userId);
  if (!user) return;
  if (!canAccessEmployeeRecord(user)) return;
  if (!isDeveloper() && user.role === "developer" && user.id !== currentUser.id) {
    alert("Somente developer pode editar contas de developer.");
    return;
  }
  editingUserId = user.id;
  userEditReturnView = returnView;
  populateUserEditForm(user);
  setView("userEdit");
}

function renderUserEditPanel() {
  if (!userEditPanel) return;
  if (!currentUser) {
    userEditPanel.classList.add("hidden");
    return;
  }
  userEditPanel.classList.remove("hidden");
  const user = users.find((entry) => entry.id === editingUserId);
  if (!user) {
    if (userEditTarget) userEditTarget.textContent = "Select a user on the Users screen to edit.";
  }
}

function renderCoiReminderPanel() {
  if (!coiReminderPanel) return;
  if (!currentUser || (!isAdmin() && !isDeveloper())) {
    coiReminderPanel.classList.add("hidden");
    return;
  }

  const reminderUsers = coiReminderWindowUsers();
  coiReminderPanel.classList.remove("hidden");

  if (!reminderUsers.length) {
    coiReminderPanel.innerHTML = `
      <h3>COI Renewals</h3>
      <p class="hint">No COI expiring in the next 30 days.</p>
    `;
    return;
  }

  const rows = reminderUsers
    .map((user) => {
      const reminderState = user.reminderSentForCurrentExpiry
        ? `<span class="reminder-state-sent">Sent (${fmtDate(user.contractorCoiReminderSentAt)})</span>`
        : '<span class="reminder-state-pending">Pending</span>';
      return `<tr>
        <td>${escapeHtml(user.name || "-")}</td>
        <td>${escapeHtml(user.companyName || "-")}</td>
        <td>${escapeHtml(user.email || "-")}</td>
        <td>${escapeHtml(fmtDateOnly(user.contractorCoiExpiry))}</td>
        <td>${escapeHtml(coiReminderLabel(user.daysLeft))}</td>
        <td>${reminderState}</td>
        <td>
          <div class="actions-inline">
            <button class="secondary xs-btn" type="button" data-send-coi-reminder="${user.id}" ${
              user.reminderSentForCurrentExpiry ? "disabled" : ""
            }>Send email</button>
            <button class="secondary xs-btn" type="button" data-open-coi-file="${user.id}" ${
              user.contractorCoiFile?.dataUrl ? "" : "disabled"
            }>Open COI</button>
          </div>
        </td>
      </tr>`;
    })
    .join("");

  coiReminderPanel.innerHTML = `
    <h3>COI Renewals</h3>
    <p class="hint">When 30 days remain before expiration, the system prepares the renewal email reminder.</p>
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>Company</th>
            <th>Email</th>
            <th>COI Expiration</th>
            <th>Deadline Status</th>
            <th>Reminder</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderHomePanel() {
  if (!homePanel) return;
  const isCatalogAdmin = can("manageCatalog");
  const canManageUsers = can("manageUsers");
  const canProjects = canOpenView("projects");
  const canReports = can("report");
  const isDev = isDeveloper();

  const toggleNav = (action, visible) => {
    document.querySelectorAll(`[data-nav-action="${action}"]`).forEach((el) => {
      el.classList.toggle("hidden", !visible);
    });
  };

  toggleNav("developer", isDev);
  toggleNav("users", canManageUsers);
  toggleNav("subcontractors", canManageUsers);
  toggleNav("partners", canManageUsers);
  toggleNav("clients", isCatalogAdmin);
  toggleNav("projects", canProjects);
  toggleNav("warehouse", canProjects);
  toggleNav("delivery", canProjects);
  toggleNav("distribution", canProjects);
  toggleNav("installation", canProjects);
  toggleNav("changeOrders", canProjects);
  toggleNav("projectReports", canProjects && canReports);
  renderCoiReminderPanel();
}

function renderRoleStrip() {
  roleLine.textContent = `User: ${currentUser.name} | Role: ${roleLabel(currentUser.role)}`;
  permissionLine.textContent = rolePermissionsSummary(currentUser.role);
}

function renderDeveloperAuditPanel() {
  if (!developerAuditPanel) return;
  if (!isDeveloper()) {
    developerAuditPanel.classList.add("hidden");
    return;
  }

  const categoryLabel = (category) => {
    if (category === "navigation") return "Navigation";
    if (category === "data-change") return "Data Change";
    if (category === "unit-change") return "Unit Change";
    if (category === "container-change") return "Container Change";
    if (category === "session") return "Session";
    return "App";
  };

  const rows = ensureArray(appAuditLog)
    .slice()
    .sort((a, b) => new Date(b.at || 0).getTime() - new Date(a.at || 0).getTime())
    .slice(0, 800)
    .map(
      (entry) => `<tr>
        <td>${escapeHtml(fmtDate(entry.at))}</td>
        <td>${escapeHtml(entry.byName || "-")}</td>
        <td>${escapeHtml(categoryLabel(entry.category))}</td>
        <td>${escapeHtml(entry.scope || "-")}</td>
        <td>${escapeHtml(`${entry.view || "-"} / ${entry.sector || "-"}`)}</td>
        <td>${escapeHtml(entry.message || "-")}</td>
      </tr>`
    )
    .join("");

  developerAuditPanel.classList.remove("hidden");
  developerAuditPanel.innerHTML = `
    <h3>Audit Log (Developer)</h3>
    <p class="hint">Tracks exact field changes and navigation history by user.</p>
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Date</th>
            <th>User</th>
            <th>Type</th>
            <th>Scope</th>
            <th>Location</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td colspan="6">No audited changes found.</td></tr>'}</tbody>
      </table>
    </div>
  `;
}

function renderSyncPanel() {
  const allowed = can("sync");
  syncPanel.classList.toggle("hidden", !allowed);
  if (!allowed) return;

  syncConfigForm.supabaseUrl.value = syncConfig.supabaseUrl;
  syncConfigForm.supabaseAnonKey.value = syncConfig.supabaseAnonKey;
  syncConfigForm.tenant.value = syncConfig.tenant;
  syncConfigForm.autoSync.value = String(syncConfig.autoSync);

  const statusText = syncConfig.lastSyncAt ? `Last sync: ${fmtDate(syncConfig.lastSyncAt)}` : t("Sem sincronizacao");
  updateSyncStatus(statusText, false);
  renderDeveloperAuditPanel();
}

function renderQrLookupPanel() {
  if (!qrLookupPanel || !qrLookupResult) return;
  qrLookupPanel.classList.toggle("hidden", !currentUser);
  if (!currentUser) return;

  const lookupCode = String(lastQrLookupCode || "").trim();
  if (!lookupCode) {
    qrLookupResult.innerHTML = `<p class="hint">Busque um QR para ver status, destino e historico completo.</p>`;
    return;
  }

  const snapshot = buildQrHistory(lookupCode);
  if (!snapshot) {
    qrLookupResult.innerHTML = `<p class="hint">QR nao encontrado: ${escapeHtml(lookupCode)}</p>`;
    return;
  }

  const { container, item, client, project, unit, events } = snapshot;
  const eventRows = events
    .map(
      (entry) => `<tr>
      <td>${escapeHtml(fmtDate(entry.at))}</td>
      <td>${escapeHtml(entry.source || "-")}</td>
      <td>${escapeHtml(entry.action || "-")}</td>
      <td>${escapeHtml(entry.detail || "-")}</td>
    </tr>`
    )
    .join("");

  qrLookupResult.innerHTML = `
    <div class="kpi-grid">
      <div class="kpi-box"><span>QR</span><strong>${escapeHtml(item.qrCode)}</strong></div>
      <div class="kpi-box"><span>SKU</span><strong>${escapeHtml(item.code || "-")}</strong></div>
      <div class="kpi-box"><span>Status</span><strong>${escapeHtml(item.status || "-")}</strong></div>
      <div class="kpi-box"><span>Container</span><strong>${escapeHtml(container.containerCode || "-")}</strong></div>
      <div class="kpi-box"><span>Client</span><strong>${escapeHtml(client?.name || "-")}</strong></div>
      <div class="kpi-box"><span>Project</span><strong>${escapeHtml(project?.name || "-")}</strong></div>
      <div class="kpi-box"><span>Unit</span><strong>${escapeHtml(item.unitLabel || unit?.unitCode || "-")}</strong></div>
      <div class="kpi-box"><span>Kitchen Type</span><strong>${escapeHtml(item.kitchenType || "-")}</strong></div>
    </div>
    <div class="row">
      <button class="secondary" type="button" data-lookup-print="${escapeHtml(item.id)}">Imprimir etiqueta (PDF/Print)</button>
      <button class="secondary" type="button" data-lookup-png="${escapeHtml(item.id)}">Baixar etiqueta PNG</button>
    </div>
    <table class="data-table">
      <thead><tr><th>Date</th><th>Source</th><th>Action</th><th>Detail</th></tr></thead>
      <tbody>${eventRows || '<tr><td colspan="4">No events.</td></tr>'}</tbody>
    </table>
  `;

  const printBtn = qrLookupResult.querySelector("[data-lookup-print]");
  printBtn?.addEventListener("click", () => openQrLabelWindow(container, item, { autoPrint: true }));

  const pngBtn = qrLookupResult.querySelector("[data-lookup-png]");
  pngBtn?.addEventListener("click", () => {
    downloadQrLabelPng(item);
  });
}

function renderAccessControl() {
  unitEntryPanel.classList.toggle("hidden", !can("createUnit"));
  exportBtn.disabled = !isDeveloper();
  importInput.disabled = !isDeveloper();
  projectReportBtn.disabled = !can("report");

  const sectorAllowsUnitEntry = ["warehouse"].includes(currentProjectSector);
  if (can("createUnit")) {
    const hasProjects = projects.length > 0;
    setFormEnabled(unitForm, hasProjects && sectorAllowsUnitEntry);
    unitProjectHint.textContent = hasProjects
      ? ""
      : "Para cadastrar unidades, primeiro crie cliente e projeto no bloco de cadastro base.";
  } else {
    setFormEnabled(unitForm, false);
  }

  const canManageCont = can("manageContainers");
  const hasContBase = clients.length > 0 && projects.length > 0;
  const sectorAllowsContainer = ["fabrica", "warehouse"].includes(currentProjectSector);
  setFormEnabled(containerForm, canManageCont && hasContBase && sectorAllowsContainer);
}

function render() {
  if (!currentUser) return;
  renderHomePanel();
  renderRoleStrip();
  renderMasterData();
  renderContainers();
  renderStats();
  renderSyncPanel();
  renderQrLookupPanel();
  renderUsers();
  renderUserEditPanel();
  renderAccessControl();
  renderUnits();
  applyViewMode();
  applyLanguageToUi();
}

function toUnit(formData) {
  const projectId = formData.get("projectId")?.toString();
  const project = projectById(projectId);
  const client = project ? clientById(project.clientId) : null;

  return normalizeUnit({
    id: uid(),
    clientId: client?.id || "",
    projectId: project?.id || "",
    clientName: client?.name || formData.get("clientName")?.toString().trim(),
    projectName: project?.name || formData.get("projectName")?.toString().trim(),
    jobSite: formData.get("jobSite")?.toString().trim(),
    unitCode: formData.get("unitCode")?.toString().trim(),
    building: formData.get("building")?.toString().trim(),
    floor: formData.get("floor")?.toString().trim(),
    category: formData.get("category")?.toString().trim(),
    unitType: formData.get("unitType")?.toString().trim(),
    kitchenModel: formData.get("kitchenModel")?.toString().trim(),
    shopdrawingRef: formData.get("shopdrawingRef")?.toString().trim(),
    scopeWork: formData.get("scopeWork")?.toString().trim(),
    deliveryQuality: "pending",
    installationStatus: "not-started",
    deliveryNotes: "",
    issuesText: "",
    checkItems: [],
    dispatchTasks: [],
    siteReceipts: [],
    dispatchSignature: null,
    auditLog: [
      {
        id: uid(),
        at: new Date().toISOString(),
        byUserId: currentUser?.id || "",
        byName: currentUser?.name || "System",
        message: "Unit created",
      },
    ],
    stages: Object.fromEntries(STAGES.map((stage) => [stage.key, { done: false, at: null }])),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });
}

function syncEndpoint() {
  const url = syncConfig.supabaseUrl.trim().replace(/\/$/, "");
  if (!url || !syncConfig.supabaseAnonKey || !syncConfig.tenant) return null;
  return `${url}/rest/v1/app_records`;
}

function toCloudRecords() {
  const tenant = syncConfig.tenant.trim();

  const mapRecords = (arr, kind, updatedField = "updatedAt") =>
    arr.map((item) => ({
      tenant,
      kind,
      id: item.id,
      payload: item,
      updated_at: item[updatedField] || item.createdAt || new Date().toISOString(),
    }));

  return [
    ...mapRecords(units, "unit"),
    ...mapRecords(photos, "photo", "updatedAt"),
    ...mapRecords(users, "user"),
    ...mapRecords(clients, "client"),
    ...mapRecords(projects, "project"),
    ...mapRecords(contacts, "contact"),
    ...mapRecords(containers, "container"),
    ...mapRecords(materials, "material"),
  ];
}

async function pushCloud({ silent = false } = {}) {
  if (!can("sync")) return;
  const endpoint = syncEndpoint();
  if (!endpoint) {
    updateSyncStatus(t("Configurar Supabase URL/Key/Tenant"), true);
    return;
  }
  if (!navigator.onLine) {
    updateSyncStatus(t("Sem internet para sincronizar"), true);
    return;
  }

  try {
    updateSyncStatus(t("Enviando dados..."));

    const response = await fetch(`${endpoint}?on_conflict=tenant,kind,id`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: syncConfig.supabaseAnonKey,
        Authorization: `Bearer ${syncConfig.supabaseAnonKey}`,
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify(toCloudRecords()),
    });

    if (!response.ok) throw new Error(await response.text());

    syncConfig.lastSyncAt = new Date().toISOString();
    await saveSetting({ id: "syncConfig", ...syncConfig });
    updateSyncStatus(`${t("Push concluido:")} ${fmtDate(syncConfig.lastSyncAt)}`);
  } catch (error) {
    updateSyncStatus(t("Falha no push. Verifique a configuracao."), true);
    if (!silent) console.error(error);
  }
}

function newerThan(remoteTs, localTs) {
  return new Date(remoteTs || 0).getTime() > new Date(localTs || 0).getTime();
}

async function pullCloud() {
  if (!can("sync")) return;
  const endpoint = syncEndpoint();
  if (!endpoint) {
    updateSyncStatus(t("Configurar Supabase URL/Key/Tenant"), true);
    return;
  }
  if (!navigator.onLine) {
    updateSyncStatus(t("Sem internet para sincronizar"), true);
    return;
  }

  try {
    updateSyncStatus(t("Baixando dados..."));
    const qs = new URLSearchParams({
      select: "kind,id,payload,updated_at",
      tenant: `eq.${syncConfig.tenant.trim()}`,
      order: "updated_at.desc",
      limit: "20000",
    });

    const response = await fetch(`${endpoint}?${qs.toString()}`, {
      headers: {
        apikey: syncConfig.supabaseAnonKey,
        Authorization: `Bearer ${syncConfig.supabaseAnonKey}`,
      },
    });

    if (!response.ok) throw new Error(await response.text());

    const records = await response.json();
    const unitMap = new Map(units.map((item) => [item.id, item]));
    const photoMap = new Map(photos.map((item) => [item.id, item]));
    const userMap = new Map(users.map((item) => [item.id, item]));
    const clientMap = new Map(clients.map((item) => [item.id, item]));
    const projectMap = new Map(projects.map((item) => [item.id, item]));
    const contactMap = new Map(contacts.map((item) => [item.id, item]));
    const containerMap = new Map(containers.map((item) => [item.id, item]));
    const materialMap = new Map(materials.map((item) => [item.id, item]));

    for (const row of records) {
      const item = row.payload || {};
      const remoteUpdatedAt = item.updatedAt || row.updated_at || new Date().toISOString();
      item.updatedAt = remoteUpdatedAt;

      if (row.kind === "unit") {
        const local = unitMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(UNIT_STORE, normalizeUnit(item));
      }

      if (row.kind === "photo") {
        const local = photoMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt || local.createdAt)) await put(PHOTO_STORE, item);
      }

      if (row.kind === "user") {
        const local = userMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(USER_STORE, normalizeUser(item));
      }

      if (row.kind === "client") {
        const local = clientMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(CLIENT_STORE, normalizeClient(item));
      }

      if (row.kind === "project") {
        const local = projectMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(PROJECT_STORE, normalizeProject(item));
      }

      if (row.kind === "contact") {
        const local = contactMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(CONTACT_STORE, normalizeContact(item));
      }

      if (row.kind === "container") {
        const local = containerMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(CONTAINER_STORE, normalizeContainer(item));
      }

      if (row.kind === "material") {
        const local = materialMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) await put(MATERIAL_STORE, normalizeMaterial(item));
      }
    }

    syncConfig.lastSyncAt = new Date().toISOString();
    await saveSetting({ id: "syncConfig", ...syncConfig });
    await loadAll();

    if (currentUser) currentUser = users.find((user) => user.id === currentUser.id) || currentUser;

    render();
    updateSyncStatus(`${t("Pull concluido:")} ${fmtDate(syncConfig.lastSyncAt)}`);
  } catch (error) {
    updateSyncStatus(t("Falha no pull. Verifique a configuracao."), true);
    console.error(error);
  }
}

function queueAutoSync() {
  if (!can("sync")) return;
  if (!syncConfig.autoSync) return;
  if (!navigator.onLine) return;

  clearTimeout(autoSyncTimer);
  autoSyncTimer = setTimeout(() => {
    pushCloud({ silent: true });
  }, 1200);
}

function reportShell(title, bodyHtml) {
  return `
  <!doctype html>
  <html lang="en">
    <head>
      <meta charset="UTF-8" />
      <title>${escapeHtml(title)}</title>
      <style>
        body { font-family: Georgia, "Times New Roman", serif; margin: 24px; color: #222; }
        h1 { margin: 0 0 6px; }
        p { margin: 2px 0; }
        .meta { margin-bottom: 14px; color: #555; }
        table { width: 100%; border-collapse: collapse; margin-top: 12px; }
        th, td { border: 1px solid #bbb; padding: 6px; font-size: 12px; text-align: left; vertical-align: top; }
        th { background: #efefef; }
        .photos { margin-top: 16px; display: flex; flex-wrap: wrap; gap: 8px; }
        .photos img { width: 150px; height: 120px; object-fit: cover; border: 1px solid #aaa; }
      </style>
    </head>
    <body>${bodyHtml}</body>
  </html>`;
}

function warehouseMovesForUnit(unitId) {
  const out = [];
  for (const container of containers) {
    for (const move of container.movements) {
      if (move.unitId === unitId) {
        out.push({ ...move, containerCode: container.containerCode });
      }
    }
  }
  return out.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
}

function generateUnitReport(unitId) {
  const unit = units.find((entry) => entry.id === unitId);
  if (!unit) return;

  const projectPeople = contactsForProject(unit.projectId);
  const photosRows = photosByUnit.get(unit.id) || [];
  const warehouseMoves = warehouseMovesForUnit(unit.id);

  const stagesTable = STAGES.map((stage) => {
    const entry = unit.stages[stage.key] || { done: false, at: null };
    return `<tr><td>${escapeHtml(stageLabel(stage.key))}</td><td>${entry.done ? t("Concluida") : t("Pendente")}</td><td>${escapeHtml(
      fmtDate(entry.at)
    )}</td></tr>`;
  }).join("");

  const checklistRows = unit.checkItems
    .map(
      (item) => `<tr>
      <td>${escapeHtml(item.code)}</td>
      <td>${escapeHtml(item.description)}</td>
      <td>${escapeHtml(item.expectedQty)}</td>
      <td>${escapeHtml(item.checkedQty)}</td>
      <td>${escapeHtml(statusBadge(item.status))}</td>
      <td>${escapeHtml(item.note || "")}</td>
    </tr>`
    )
    .join("");

  const peopleRows = projectPeople
    .map(
      (person) => `<tr>
      <td>${escapeHtml(person.name)}</td>
      <td>${escapeHtml(contactRoleLabel(person.role))}</td>
      <td>${escapeHtml(person.company)}</td>
      <td>${escapeHtml(person.phone)}</td>
      <td>${escapeHtml(person.email)}</td>
    </tr>`
    )
    .join("");

  const warehouseRows = warehouseMoves
    .map(
      (move) => `<tr>
      <td>${escapeHtml(fmtDate(move.createdAt))}</td>
      <td>${escapeHtml(move.containerCode)}</td>
      <td>${escapeHtml(move.action === "release" ? "Dispatch" : "Hold")}</td>
      <td>K:${toNumber(move.kitchens)} V:${toNumber(move.vanities)} M:${toNumber(move.medCabinets)} C:${toNumber(move.countertops)}</td>
      <td>${escapeHtml(move.destination || "-")}</td>
      <td>${escapeHtml(move.operatorName || "-")}</td>
    </tr>`
    )
    .join("");

  const html = reportShell(
    `Unit report ${unit.unitCode}`,
    `
      <h1>Unit report ${escapeHtml(unit.unitCode)}</h1>
      <div class="meta">
        <p><strong>Client:</strong> ${escapeHtml(unit.clientName)}</p>
        <p><strong>Project:</strong> ${escapeHtml(unit.projectName)}</p>
        <p><strong>Job Site:</strong> ${escapeHtml(unit.jobSite)}</p>
        <p><strong>Category:</strong> ${escapeHtml(unit.category)}</p>
        <p><strong>Type:</strong> ${escapeHtml(unit.unitType)}</p>
        <p><strong>Kitchen model:</strong> ${escapeHtml(unit.kitchenModel)}</p>
        <p><strong>Shopdrawing:</strong> ${escapeHtml(unit.shopdrawingRef || "-")}</p>
        <p><strong>Scope:</strong> ${escapeHtml(unit.scopeWork || "-")}</p>
        <p><strong>Quality:</strong> ${escapeHtml(unit.deliveryQuality)}</p>
        <p><strong>Installation:</strong> ${escapeHtml(unit.installationStatus)}</p>
        <p><strong>Issues:</strong> ${escapeHtml(unit.issuesText || "-")}</p>
        <p><strong>Notes:</strong> ${escapeHtml(unit.deliveryNotes || "-")}</p>
        <p><strong>Issued at:</strong> ${escapeHtml(fmtDate(new Date().toISOString()))}</p>
      </div>

      <h3>Operational flow</h3>
      <table>
        <thead><tr><th>Stage</th><th>Status</th><th>Date/Time</th></tr></thead>
        <tbody>${stagesTable}</tbody>
      </table>

      <h3>Detailed checklist</h3>
      <table>
        <thead><tr><th>Code</th><th>Description</th><th>Expected Qty</th><th>Checked Qty</th><th>Status</th><th>Notes</th></tr></thead>
        <tbody>${checklistRows || '<tr><td colspan="6">No items.</td></tr>'}</tbody>
      </table>

      <h3>Warehouse movements (containers)</h3>
      <table>
        <thead><tr><th>Date</th><th>Container</th><th>Movement</th><th>Qty</th><th>Destination</th><th>Operator</th></tr></thead>
        <tbody>${warehouseRows || '<tr><td colspan="6">No linked movements.</td></tr>'}</tbody>
      </table>

      <h3>Project team</h3>
      <table>
        <thead><tr><th>Name</th><th>Role</th><th>Company</th><th>Phone</th><th>Email</th></tr></thead>
        <tbody>${peopleRows || '<tr><td colspan="5">No linked people.</td></tr>'}</tbody>
      </table>

      <h3>Photos</h3>
      <div class="photos">${photosRows.map((photo) => `<img src="${photo.dataUrl}" alt="${escapeHtml(photo.name)}" />`).join("") || "No photos"}</div>
    `
  );

  const win = window.open("", "_blank");
  if (!win) return;
  win.document.open();
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 300);
}

function generateProjectReport(projectId) {
  const project = projectById(projectId);
  const scopedUnits =
    projectId === "all"
      ? units
      : units.filter((unit) => unit.projectId === projectId || (project && unit.projectName === project.name));

  if (!scopedUnits.length) {
    alert(t("Nao ha unidades para gerar o relatorio."));
    return;
  }

  const scopedContainers =
    projectId === "all" ? containers : containers.filter((container) => container.projectId === projectId);

  const unitRows = scopedUnits
    .map(
      (unit) => `<tr>
      <td>${escapeHtml(unit.clientName)}</td>
      <td>${escapeHtml(unit.projectName)}</td>
      <td>${escapeHtml(unit.unitCode)}</td>
      <td>${escapeHtml(unit.category)}</td>
      <td>${escapeHtml(unit.deliveryQuality)}</td>
      <td>${escapeHtml(unit.installationStatus)}</td>
      <td>${stageDoneCount(unit)}/${STAGES.length}</td>
      <td>${escapeHtml(checklistSummary(unit.checkItems))}</td>
      <td>${escapeHtml(unit.issuesText || "-")}</td>
    </tr>`
    )
    .join("");

  const containerRows = scopedContainers
    .map((container) => {
      const client = clientById(container.clientId);
      const projectEntry = projectById(container.projectId);
      const av = containerAvailableQty(container);
      return `<tr>
      <td>${escapeHtml(container.containerCode)}</td>
      <td>${escapeHtml(client?.name || "-")}</td>
      <td>${escapeHtml(projectEntry?.name || "-")}</td>
      <td>${escapeHtml(containerStatusLabel(container.arrivalStatus))}</td>
      <td>${escapeHtml(container.etaDate || "-")}</td>
      <td>${container.qtyKitchens}/${container.qtyVanities}/${container.qtyMedCabinets}/${container.qtyCountertops}</td>
      <td>${av.kitchen}/${av.vanity}/${av.medCabinet}/${av.countertop}</td>
    </tr>`;
    })
    .join("");

  const title = projectId === "all" ? "All projects" : project?.name || "Project";

  const html = reportShell(
    `Project report ${title}`,
    `
      <h1>Project report: ${escapeHtml(title)}</h1>
      <div class="meta">
        <p><strong>Total units:</strong> ${scopedUnits.length}</p>
        <p><strong>Total containers:</strong> ${scopedContainers.length}</p>
        <p><strong>Generated at:</strong> ${escapeHtml(fmtDate(new Date().toISOString()))}</p>
      </div>
      <h3>Units</h3>
      <table>
        <thead>
          <tr>
            <th>Client</th>
            <th>Project</th>
            <th>Unit</th>
            <th>Category</th>
            <th>Quality</th>
            <th>Installation</th>
            <th>Progress</th>
            <th>Checklist</th>
            <th>Issues</th>
          </tr>
        </thead>
        <tbody>${unitRows}</tbody>
      </table>
      <h3>Container schedule and balance</h3>
      <table>
        <thead>
          <tr>
            <th>Container</th>
            <th>Client</th>
            <th>Project</th>
            <th>Status</th>
            <th>ETA</th>
            <th>Planned Qty K/V/M/C</th>
            <th>Balance K/V/M/C</th>
          </tr>
        </thead>
        <tbody>${containerRows || '<tr><td colspan="7">No containers.</td></tr>'}</tbody>
      </table>
    `
  );

  const win = window.open("", "_blank");
  if (!win) return;
  win.document.open();
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 300);
}

setupForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(setupForm);

  const user = {
    id: uid(),
    name: data.get("name")?.toString().trim(),
    username: data.get("username")?.toString().trim().toLowerCase(),
    passwordHash: await hashPassword(data.get("password")?.toString() || ""),
    role: "developer",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  await put(USER_STORE, user);
  await loadAll();
  currentUser = user;
  hasAutoDispatchedCoiReminder = false;
  setSession(user.id);
  window.history.replaceState(null, "", "#home");
  setupForm.reset();
  renderAuth();
});

openSignupBtn?.addEventListener("click", () => {
  showSignupMode(true);
  toggleSubcontractorExtras("signup", signupEmploymentTypeSelect?.value || "tag");
  signupForm?.querySelector('input[name="firstName"]')?.focus();
});

backToLoginBtn?.addEventListener("click", () => {
  showSignupMode(false);
  loginForm?.querySelector('input[name="username"]')?.focus();
});

signupEmploymentTypeSelect?.addEventListener("change", () => {
  toggleSubcontractorExtras("signup", signupEmploymentTypeSelect.value);
});

userEditEmploymentTypeSelect?.addEventListener("change", () => {
  toggleSubcontractorExtras("edit", userEditEmploymentTypeSelect.value);
});

userAdminEmploymentTypeSelect?.addEventListener("change", () => {
  toggleSubcontractorExtras("admin", userAdminEmploymentTypeSelect.value);
});

signupForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(signupForm);
  const username = data.get("username")?.toString().trim().toLowerCase();
  const firstName = data.get("firstName")?.toString().trim() || "";
  const lastName = data.get("lastName")?.toString().trim() || "";
  const fullName = `${firstName} ${lastName}`.trim();
  if (!username) return;
  if (users.some((user) => user.username === username)) {
    alert(t("Usuario ja existe."));
    return;
  }
  const photoFile = signupForm.querySelector('input[name="photo"]')?.files?.[0];
  const photoDataUrl = photoFile ? await fileToDataUrl(photoFile) : "";
  const employmentType = data.get("employmentType")?.toString() || "";
  const contractorAreas = employmentType === "subcontractor" ? collectCheckedValues(signupForm, "contractorAreas") : [];
  const contractorCoiExpiry = employmentType === "subcontractor" ? data.get("contractorCoiExpiry")?.toString() || "" : "";
  const contractorCoiFileRaw = signupForm.querySelector('input[name="contractorCoiFile"]')?.files?.[0] || null;
  let contractorCoiFile = null;

  if (employmentType === "subcontractor") {
    if (!contractorCoiExpiry) {
      alert("Fill in the COI expiration date.");
      return;
    }
    const validation = validateCoiFile(contractorCoiFileRaw);
    if (!validation.ok) {
      alert(validation.message);
      return;
    }
    contractorCoiFile = {
      name: contractorCoiFileRaw.name || "coi",
      type: contractorCoiFileRaw.type || "",
      dataUrl: await fileToDataUrl(contractorCoiFileRaw),
      uploadedAt: new Date().toISOString(),
    };
  }

  const user = {
    id: uid(),
    name: fullName,
    firstName,
    lastName,
    companyName: data.get("companyName")?.toString().trim() || "",
    jobTitle: data.get("jobTitle")?.toString().trim() || "",
    employmentType,
    contractorCoi: employmentType === "subcontractor" ? data.get("contractorCoi")?.toString().trim() || "" : "",
    contractorCoiExpiry,
    contractorCoiFile,
    contractorW9: employmentType === "subcontractor" ? data.get("contractorW9")?.toString().trim() || "" : "",
    contractorAreas,
    contractorCoiReminderSentAt: "",
    contractorCoiReminderForExpiry: "",
    address: data.get("address")?.toString().trim() || "",
    phone: data.get("phone")?.toString().trim() || "",
    cellPhone: data.get("cellPhone")?.toString().trim() || "",
    email: data.get("email")?.toString().trim().toLowerCase() || "",
    photoDataUrl,
    username,
    passwordHash: await hashPassword(data.get("password")?.toString() || ""),
    role: "visitor",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
  await put(USER_STORE, normalizeUser(user));
  await loadAll();
  signupForm.reset();
  toggleSubcontractorExtras("signup", "tag");
  showSignupMode(false);
  alert("Registration submitted. An admin or developer can adjust your roles and access.");
  queueAutoSync();
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(loginForm);
  const username = data.get("username")?.toString().trim().toLowerCase();
  const passwordHash = await hashPassword(data.get("password")?.toString() || "");

  const user = users.find((entry) => entry.username === username && entry.passwordHash === passwordHash);
  if (!user) {
    alert(t("Usuario ou senha invalidos."));
    return;
  }

  currentUser = user;
  hasAutoDispatchedCoiReminder = false;
  setSession(user.id);
  pushAppAudit("Session login", "session", "auth");
  window.history.replaceState(null, "", "#home");
  loginForm.reset();
  renderAuth();
});

logoutBtn.addEventListener("click", () => {
  if (currentUser) pushAppAudit("Session logout", "session", "auth");
  currentUser = null;
  hasAutoDispatchedCoiReminder = false;
  editingUserId = "";
  userEditReturnView = "home";
  clearSession();
  window.history.replaceState(null, "", window.location.pathname + window.location.search);
  renderAuth();
});

userForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageUsers")) return;

  const data = new FormData(userForm);
  const firstName = data.get("firstName")?.toString().trim() || "";
  const lastName = data.get("lastName")?.toString().trim() || "";
  const formCompanyName = data.get("companyName")?.toString().trim() || "";
  const companyName = isAdmin() && !canViewAllEmployees() ? currentUser?.companyName || formCompanyName : formCompanyName;
  const jobTitle = data.get("jobTitle")?.toString().trim() || "";
  const employmentType = data.get("employmentType")?.toString() || "tag";
  const contractorCoi = employmentType === "subcontractor" ? data.get("contractorCoi")?.toString().trim() || "" : "";
  const contractorCoiExpiry = employmentType === "subcontractor" ? data.get("contractorCoiExpiry")?.toString() || "" : "";
  const contractorW9 = employmentType === "subcontractor" ? data.get("contractorW9")?.toString().trim() || "" : "";
  const contractorAreas = employmentType === "subcontractor" ? collectCheckedValues(userForm, "contractorAreas") : [];
  const address = data.get("address")?.toString().trim() || "";
  const phone = maskPhoneValue(data.get("phone")?.toString().trim() || "");
  const cellPhone = maskPhoneValue(data.get("cellPhone")?.toString().trim() || "");
  const email = data.get("email")?.toString().trim().toLowerCase() || "";
  const username = data.get("username")?.toString().trim().toLowerCase() || "";
  const plainPassword = data.get("password")?.toString() || "";
  const selectedRole = data.get("role")?.toString() || "visitor";

  if (!firstName || !lastName || !companyName || !jobTitle || !address || !cellPhone || !email || !username) {
    alert("Fill in all required fields.");
    return;
  }
  if (!isDeveloper() && selectedRole === "developer") {
    alert("Only developer can create or edit a developer account.");
    return;
  }
  if (employmentType === "subcontractor" && !contractorCoiExpiry) {
    alert("Fill in the COI expiration date.");
    return;
  }

  const targetUser = adminEditingUserId ? users.find((entry) => entry.id === adminEditingUserId) : null;
  if (targetUser && !canAccessEmployeeRecord(targetUser)) {
    alert("You can only manage users from your own company.");
    return;
  }
  if (isAdmin() && !canViewAllEmployees()) {
    const myCompany = normalizeCompanyName(currentUser?.companyName);
    if (myCompany && normalizeCompanyName(companyName) !== myCompany) {
      alert("You can only manage users from your own company.");
      return;
    }
  }
  if (!targetUser && !plainPassword) {
    alert("Password is required to create a new user.");
    return;
  }

  const usernameTaken = users.some(
    (user) => user.username === username && (!targetUser || user.id !== targetUser.id)
  );
  if (usernameTaken) {
    alert(t("Usuario ja existe."));
    return;
  }

  const photoFile = userForm.querySelector('input[name="photo"]')?.files?.[0];
  const contractorCoiFileRaw = userForm.querySelector('input[name="contractorCoiFile"]')?.files?.[0] || null;
  let contractorCoiFile = employmentType === "subcontractor" ? targetUser?.contractorCoiFile || null : null;
  if (employmentType === "subcontractor" && contractorCoiFileRaw) {
    const validation = validateCoiFile(contractorCoiFileRaw);
    if (!validation.ok) {
      alert(validation.message);
      return;
    }
    contractorCoiFile = {
      name: contractorCoiFileRaw.name || "coi",
      type: contractorCoiFileRaw.type || "",
      dataUrl: await fileToDataUrl(contractorCoiFileRaw),
      uploadedAt: new Date().toISOString(),
    };
  }
  if (employmentType === "subcontractor" && !contractorCoiFile?.dataUrl) {
    alert("Attach the COI file in PDF or JPG format.");
    return;
  }

  const photoDataUrl = photoFile ? await fileToDataUrl(photoFile) : targetUser?.photoDataUrl || "";
  const passwordHash = plainPassword ? await hashPassword(plainPassword) : targetUser?.passwordHash || "";
  const keepReminderState =
    Boolean(targetUser) &&
    employmentType === "subcontractor" &&
    targetUser.contractorCoiReminderForExpiry === contractorCoiExpiry &&
    Boolean(targetUser.contractorCoiReminderSentAt);

  const savedUser = normalizeUser({
    ...(targetUser || {}),
    id: targetUser?.id || uid(),
    name: `${firstName} ${lastName}`.trim(),
    firstName,
    lastName,
    companyName,
    jobTitle,
    employmentType,
    contractorCoi,
    contractorCoiExpiry,
    contractorCoiFile,
    contractorW9,
    contractorAreas,
    contractorCoiReminderSentAt: keepReminderState ? targetUser.contractorCoiReminderSentAt : "",
    contractorCoiReminderForExpiry: keepReminderState ? targetUser.contractorCoiReminderForExpiry : "",
    address,
    phone,
    cellPhone,
    email,
    photoDataUrl,
    username,
    passwordHash,
    role: selectedRole,
    createdAt: targetUser?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });

  await put(USER_STORE, savedUser);
  if (!targetUser) {
    pushEntityAudit(
      "Users",
      "created",
      `${savedUser.name} (${savedUser.username}) role "${savedUser.role}" company "${auditValue(savedUser.companyName)}"`,
      "users"
    );
  } else {
    const changes = collectAuditChanges(targetUser, savedUser, [
      { key: "firstName", label: "First name" },
      { key: "lastName", label: "Last name" },
      { key: "companyName", label: "Company name" },
      { key: "jobTitle", label: "Job title" },
      { key: "employmentType", label: "Employment type" },
      { key: "contractorCoi", label: "COI" },
      { key: "contractorCoiExpiry", label: "COI expiration" },
      { key: "contractorW9", label: "W9" },
      { key: "contractorAreas", label: "Sub contractor areas", map: (value) => ensureArray(value).sort().join(", ") },
      { key: "address", label: "Address" },
      { key: "phone", label: "Phone" },
      { key: "cellPhone", label: "Cell phone" },
      { key: "email", label: "Email" },
      { key: "username", label: "Username" },
      { key: "role", label: "Role" },
    ]);
    if (plainPassword) changes.push("Password: updated");
    if (photoFile) changes.push("Photo: updated");
    if (employmentType === "subcontractor" && contractorCoiFileRaw) changes.push("COI file: updated");
    pushEntityAudit(
      "Users",
      "updated",
      `${savedUser.name} (${savedUser.username})${changes.length ? ` | ${changes.join("; ")}` : " | no field changes"}`,
      "users"
    );
  }

  await loadAll();
  if (currentUser) currentUser = users.find((entry) => entry.id === currentUser.id) || currentUser;
  const refreshed = adminEditingUserId ? users.find((entry) => entry.id === adminEditingUserId) : null;
  if (refreshed) populateAdminUserForm(refreshed);
  else resetAdminUserForm();
  render();
  queueAutoSync();
});

userFormNewBtn?.addEventListener("click", () => {
  if (!can("manageUsers")) return;
  resetAdminUserForm();
});

userFormDeleteBtn?.addEventListener("click", async () => {
  if (!can("manageUsers")) return;
  if (!adminEditingUserId) return;
  if (adminEditingUserId === currentUser?.id) {
    alert("You cannot delete the currently logged-in user.");
    return;
  }
  const targetUser = users.find((entry) => entry.id === adminEditingUserId);
  if (!targetUser) return;
  if (!canAccessEmployeeRecord(targetUser)) {
    alert("You can only manage users from your own company.");
    return;
  }
  if (!isDeveloper() && (targetUser.role === "developer" || isPrimaryDeveloperUser(targetUser))) {
    alert("Developer account cannot be deleted by admin.");
    return;
  }
  pushEntityAudit("Users", "deleted", `${targetUser.name} (${targetUser.username}) role "${targetUser.role}"`, "users");
  await del(USER_STORE, targetUser.id);
  await loadAll();
  if (currentUser) currentUser = users.find((entry) => entry.id === currentUser.id) || currentUser;
  resetAdminUserForm();
  render();
  queueAutoSync();
});

cancelUserEditBtn?.addEventListener("click", () => {
  editingUserId = "";
  setView(canOpenView(userEditReturnView) ? userEditReturnView : "home");
});

userEditForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!currentUser) return;
  const targetUser = users.find((entry) => entry.id === editingUserId);
  if (!targetUser) {
    setView("home");
    return;
  }
  if (!isDeveloper() && !isAdmin() && targetUser.id !== currentUser.id) return;
  if (!isDeveloper() && targetUser.role === "developer" && targetUser.id !== currentUser.id) return;
  if (!canAccessEmployeeRecord(targetUser)) return;

  const data = new FormData(userEditForm);
  const firstName = data.get("firstName")?.toString().trim() || "";
  const lastName = data.get("lastName")?.toString().trim() || "";
  const companyName = data.get("companyName")?.toString().trim() || "";
  const jobTitle = data.get("jobTitle")?.toString().trim() || "";
  const employmentType = data.get("employmentType")?.toString() || "tag";
  const contractorCoi = employmentType === "subcontractor" ? data.get("contractorCoi")?.toString().trim() || "" : "";
  const contractorCoiExpiry = employmentType === "subcontractor" ? data.get("contractorCoiExpiry")?.toString() || "" : "";
  const contractorW9 = employmentType === "subcontractor" ? data.get("contractorW9")?.toString().trim() || "" : "";
  const contractorAreas = employmentType === "subcontractor" ? collectCheckedValues(userEditForm, "contractorAreas") : [];
  const address = data.get("address")?.toString().trim() || "";
  const phone = maskPhoneValue(data.get("phone")?.toString().trim() || "");
  const cellPhone = maskPhoneValue(data.get("cellPhone")?.toString().trim() || "");
  const email = data.get("email")?.toString().trim().toLowerCase() || "";

  if (isAdmin() && !canViewAllEmployees()) {
    const myCompany = normalizeCompanyName(currentUser?.companyName);
    if (myCompany && normalizeCompanyName(companyName) !== myCompany) {
      alert("You can only manage users from your own company.");
      return;
    }
  }

  if (!firstName || !lastName || !companyName || !jobTitle || !address || !cellPhone || !email) {
    alert("Fill in all required fields.");
    return;
  }
  if (employmentType === "subcontractor" && !contractorCoiExpiry) {
    alert("Fill in the COI expiration date.");
    return;
  }

  const photoFile = userEditForm.querySelector('input[name="photo"]')?.files?.[0];
  const contractorCoiFileRaw = userEditForm.querySelector('input[name="contractorCoiFile"]')?.files?.[0] || null;
  let contractorCoiFile = employmentType === "subcontractor" ? targetUser.contractorCoiFile || null : null;
  if (employmentType === "subcontractor" && contractorCoiFileRaw) {
    const validation = validateCoiFile(contractorCoiFileRaw);
    if (!validation.ok) {
      alert(validation.message);
      return;
    }
    contractorCoiFile = {
      name: contractorCoiFileRaw.name || "coi",
      type: contractorCoiFileRaw.type || "",
      dataUrl: await fileToDataUrl(contractorCoiFileRaw),
      uploadedAt: new Date().toISOString(),
    };
  }
  if (employmentType === "subcontractor" && !contractorCoiFile?.dataUrl) {
    alert("Attach the COI file in PDF or JPG format.");
    return;
  }
  const photoDataUrl = photoFile ? await fileToDataUrl(photoFile) : targetUser.photoDataUrl || "";
  const keepReminderState =
    employmentType === "subcontractor" &&
    targetUser.contractorCoiReminderForExpiry === contractorCoiExpiry &&
    Boolean(targetUser.contractorCoiReminderSentAt);

  const savedUser = normalizeUser({
    ...targetUser,
    name: `${firstName} ${lastName}`.trim(),
    firstName,
    lastName,
    companyName,
    jobTitle,
    employmentType,
    contractorCoi,
    contractorCoiExpiry,
    contractorCoiFile,
    contractorW9,
    contractorAreas,
    contractorCoiReminderSentAt: keepReminderState ? targetUser.contractorCoiReminderSentAt : "",
    contractorCoiReminderForExpiry: keepReminderState ? targetUser.contractorCoiReminderForExpiry : "",
    address,
    phone,
    cellPhone,
    email,
    photoDataUrl,
    updatedAt: new Date().toISOString(),
  });

  await put(USER_STORE, savedUser);
  const changes = collectAuditChanges(targetUser, savedUser, [
    { key: "firstName", label: "First name" },
    { key: "lastName", label: "Last name" },
    { key: "companyName", label: "Company name" },
    { key: "jobTitle", label: "Job title" },
    { key: "employmentType", label: "Employment type" },
    { key: "contractorCoi", label: "COI" },
    { key: "contractorCoiExpiry", label: "COI expiration" },
    { key: "contractorW9", label: "W9" },
    { key: "contractorAreas", label: "Sub contractor areas", map: (value) => ensureArray(value).sort().join(", ") },
    { key: "address", label: "Address" },
    { key: "phone", label: "Phone" },
    { key: "cellPhone", label: "Cell phone" },
    { key: "email", label: "Email" },
  ]);
  if (photoFile) changes.push("Photo: updated");
  if (employmentType === "subcontractor" && contractorCoiFileRaw) changes.push("COI file: updated");
  pushEntityAudit(
    "Users",
    "updated",
    `${savedUser.name} (${savedUser.username})${changes.length ? ` | ${changes.join("; ")}` : " | no field changes"}`,
    "users"
  );

  await loadAll();
  if (currentUser) currentUser = users.find((entry) => entry.id === currentUser.id) || currentUser;
  const refreshed = users.find((entry) => entry.id === editingUserId);
  if (refreshed) populateUserEditForm(refreshed);
  render();
  queueAutoSync();
  setView(canOpenView(userEditReturnView) ? userEditReturnView : "home");
});

clientForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(clientForm);
  const editingId = data.get("clientId")?.toString().trim() || "";
  const editingClient = editingId ? clients.find((client) => client.id === editingId) : null;
  const name = data.get("name")?.toString().trim();
  if (!name) return;

  if (clients.some((client) => client.name.toLowerCase() === name.toLowerCase() && client.id !== editingId)) {
    alert(t("Cliente ja cadastrado."));
    return;
  }

  const logoFile = clientForm.querySelector('input[name="logo"]')?.files?.[0];
  const logoDataUrl = logoFile ? await fileToDataUrl(logoFile) : editingClient?.logoDataUrl || "";
  const officePhone = data.get("officePhone")?.toString().trim() || "";
  const nowIso = new Date().toISOString();
  const nextId = editingClient?.id || uid();

  const savedClient = normalizeClient({
    ...(editingClient || {}),
    id: nextId,
    name,
    address: data.get("address")?.toString().trim(),
    officePhone,
    ownerName: data.get("ownerName")?.toString().trim(),
    contactOwner: data.get("contactOwner")?.toString().trim(),
    contactSeniorProjectManager: data.get("contactSeniorProjectManager")?.toString().trim(),
    seniorProjectManagerPhone: data.get("seniorProjectManagerPhone")?.toString().trim(),
    contactPerson: editingClient?.contactPerson || "",
    phone: officePhone,
    email: data.get("email")?.toString().trim(),
    offices: data.get("offices")?.toString().trim(),
    yearsInBusiness: data.get("yearsInBusiness")?.toString().trim(),
    logoDataUrl,
    createdAt: editingClient?.createdAt || nowIso,
    updatedAt: nowIso,
  });

  await put(CLIENT_STORE, savedClient);
  if (!editingClient) {
    pushEntityAudit("Clients", "created", `${savedClient.name} | office "${auditValue(savedClient.officePhone)}"`, "clients");
  } else {
    const changes = collectAuditChanges(editingClient, savedClient, [
      { key: "name", label: "Company name" },
      { key: "address", label: "Address" },
      { key: "officePhone", label: "Office phone" },
      { key: "email", label: "General email" },
      { key: "ownerName", label: "Owner" },
      { key: "contactOwner", label: "Owner contact" },
      { key: "contactSeniorProjectManager", label: "Senior PM" },
      { key: "seniorProjectManagerPhone", label: "Senior PM phone" },
      { key: "yearsInBusiness", label: "Years in business" },
      { key: "offices", label: "Other offices" },
    ]);
    if (logoFile) changes.push("Company logo: updated");
    pushEntityAudit(
      "Clients",
      "updated",
      `${savedClient.name}${changes.length ? ` | ${changes.join("; ")}` : " | no field changes"}`,
      "clients"
    );
  }

  selectedClientId = nextId;
  await loadAll();
  const refreshed = clients.find((client) => client.id === selectedClientId) || null;
  if (refreshed) populateClientForm(refreshed);
  render();
  queueAutoSync();
});

clientSearchInput?.addEventListener("input", () => {
  clientSearchQuery = (clientSearchInput.value || "").trim();
  renderClientsTable();
});

goProjectsFromClientBtn?.addEventListener("click", () => {
  openClientRegistrationShortcut("projects");
});

goPeopleFromClientBtn?.addEventListener("click", () => {
  openClientRegistrationShortcut("people");
});

clientFormNewBtn?.addEventListener("click", () => {
  if (!can("manageCatalog")) return;
  resetClientForm();
});

clientFormDeleteBtn?.addEventListener("click", async () => {
  if (!can("manageCatalog")) return;
  const targetId = clientForm.elements.clientId?.value || "";
  if (!targetId) return;
  if (projects.some((project) => project.clientId === targetId)) {
    alert(t("Nao e possivel excluir cliente com projetos vinculados."));
    return;
  }
  if (contacts.some((contact) => contact.clientId === targetId)) {
    alert(t("Nao e possivel excluir cliente com pessoas vinculadas."));
    return;
  }
  if (containers.some((container) => container.clientId === targetId)) {
    alert(t("Nao e possivel excluir cliente com containers vinculados."));
    return;
  }
  const targetClient = clients.find((entry) => entry.id === targetId);
  if (targetClient) pushEntityAudit("Clients", "deleted", `${targetClient.name}`, "clients");
  await del(CLIENT_STORE, targetId);
  selectedClientId = "";
  resetClientForm();
  await loadAll();
  render();
  queueAutoSync();
});

projectForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(projectForm);
  const clientId = data.get("clientId")?.toString();
  const name = data.get("name")?.toString().trim();

  if (!clientId || !name) {
    alert(t("Cliente e nome do projeto sao obrigatorios."));
    return;
  }

  if (projects.some((project) => project.clientId === clientId && project.name.toLowerCase() === name.toLowerCase())) {
    alert(t("Projeto ja cadastrado para este cliente."));
    return;
  }

  const createdProject = normalizeProject({
    id: uid(),
    clientId,
    name,
    code: data.get("code")?.toString().trim(),
    address: data.get("address")?.toString().trim(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });

  await put(PROJECT_STORE, createdProject);
  const linkedClient = clientById(createdProject.clientId);
  pushEntityAudit(
    "Projects",
    "created",
    `${createdProject.name} | client "${auditValue(linkedClient?.name || "-")}" | code "${auditValue(createdProject.code)}"`,
    "projects"
  );

  projectForm.reset();
  await loadAll();
  render();
  queueAutoSync();
});

contactForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(contactForm);
  const name = data.get("name")?.toString().trim();
  if (!name) return;

  const createdContact = normalizeContact({
    id: uid(),
    name,
    role: data.get("role")?.toString(),
    company: data.get("company")?.toString().trim(),
    clientId: data.get("clientId")?.toString() || "",
    projectId: data.get("projectId")?.toString() || "",
    phone: data.get("phone")?.toString().trim(),
    email: data.get("email")?.toString().trim(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });

  await put(CONTACT_STORE, createdContact);
  pushEntityAudit(
    "Contacts",
    "created",
    `${createdContact.name} | role "${auditValue(contactRoleLabel(createdContact.role))}" | company "${auditValue(createdContact.company)}"`,
    "contacts"
  );

  contactForm.reset();
  await loadAll();
  render();
  queueAutoSync();
});

materialForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(materialForm);
  const sku = data.get("sku")?.toString().trim();
  const description = data.get("description")?.toString().trim();
  if (!sku || !description) return;

  if (materials.some((material) => material.sku.toLowerCase() === sku.toLowerCase())) {
    alert("SKU ja cadastrado.");
    return;
  }

  const createdMaterial = normalizeMaterial({
    id: uid(),
    sku,
    description,
    category: data.get("category")?.toString() || "other",
    unit: data.get("unit")?.toString().trim() || "pcs",
    kitchenType: data.get("kitchenType")?.toString().trim() || "",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });

  await put(MATERIAL_STORE, createdMaterial);
  pushEntityAudit(
    "Materials",
    "created",
    `${createdMaterial.sku} | ${createdMaterial.description} | category "${auditValue(createdMaterial.category)}"`,
    "materials"
  );

  materialForm.reset();
  await loadAll();
  render();
  queueAutoSync();
});

containerClientSelect.addEventListener("change", syncContainerProjectSelect);
contactClientSelect.addEventListener("change", syncContactProjectSelect);
unitProjectSelect.addEventListener("change", syncUnitProjectInfo);

containerForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageContainers")) return;

  const data = new FormData(containerForm);
  const containerCode = data.get("containerCode")?.toString().trim();
  const clientId = data.get("clientId")?.toString();
  const projectId = data.get("projectId")?.toString();

  if (!containerCode || !clientId || !projectId) {
    alert(t("Container, cliente e projeto sao obrigatorios."));
    return;
  }

  if (containers.some((c) => c.containerCode.toLowerCase() === containerCode.toLowerCase())) {
    alert(t("Container ID ja cadastrado."));
    return;
  }

  const createdContainer = normalizeContainer({
    id: uid(),
    containerCode,
    supplier: data.get("supplier")?.toString().trim(),
    manufacturer: data.get("manufacturer")?.toString().trim(),
    clientId,
    projectId,
    etaDate: data.get("etaDate")?.toString(),
    departureAt: data.get("departureAt")?.toString() || "",
    arrivalStatus: data.get("arrivalStatus")?.toString(),
    qtyKitchens: toNumber(data.get("qtyKitchens")?.toString()),
    qtyVanities: toNumber(data.get("qtyVanities")?.toString()),
    qtyMedCabinets: toNumber(data.get("qtyMedCabinets")?.toString()),
    qtyCountertops: toNumber(data.get("qtyCountertops")?.toString()),
    notes: data.get("notes")?.toString().trim(),
    loosePartsNotes: data.get("loosePartsNotes")?.toString().trim(),
    materialItems: [],
    qrItems: [],
    movements: [],
    replacementQueue: [],
    factoryMissingQueue: [],
    auditLog: [],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });
  pushContainerAudit(createdContainer, `Container created: ${containerCode}`);
  await put(CONTAINER_STORE, createdContainer);

  containerForm.reset();
  syncContainerProjectSelect();
  await loadAll();
  render();
  queueAutoSync();
});

unitForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("createUnit")) return;

  const projectId = unitProjectSelect.value;
  if (!projectId) {
    alert(t("Selecione um projeto."));
    return;
  }

  const unit = toUnit(new FormData(unitForm));
  await put(UNIT_STORE, unit);
  pushAppAudit(`[Unit ${unit.unitCode || unit.id}] Unit created`, "unit-change", unit.projectName || "-");
  unitForm.reset();
  unitProjectSelect.value = projectId;
  syncUnitProjectInfo();
  await loadAll();
  render();
  queueAutoSync();
});

searchInput.addEventListener("input", renderUnits);
filterSelect.addEventListener("change", renderUnits);
qrLookupForm?.addEventListener("submit", (event) => {
  event.preventDefault();
  lastQrLookupCode = qrLookupInput?.value?.trim() || "";
  renderQrLookupPanel();
});
qrLookupInput?.addEventListener("input", () => {
  if (!qrLookupInput.value.trim()) {
    lastQrLookupCode = "";
    renderQrLookupPanel();
  }
});

projectReportBtn.addEventListener("click", () => {
  if (!can("report")) return;
  generateProjectReport(projectReportSelect.value);
});

exportBtn.addEventListener("click", async () => {
  if (!isDeveloper()) return;

  const payload = {
    exportedAt: new Date().toISOString(),
    units,
    photos,
    users,
    clients,
    projects,
    contacts,
    containers,
    materials,
    settings: [{ id: "syncConfig", ...syncConfig }],
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `cabinets-control-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
});

importInput.addEventListener("change", async () => {
  if (!isDeveloper()) return;
  const file = importInput.files?.[0];
  if (!file) return;

  try {
    const text = await file.text();
    const payload = JSON.parse(text);

    if (!Array.isArray(payload.units) || !Array.isArray(payload.photos)) {
      throw new Error("Formato invalido");
    }

    for (const unit of payload.units) await put(UNIT_STORE, normalizeUnit(unit));
    for (const photo of payload.photos) await put(PHOTO_STORE, photo);
    if (Array.isArray(payload.users)) for (const user of payload.users) await put(USER_STORE, normalizeUser(user));
    if (Array.isArray(payload.clients)) for (const client of payload.clients) await put(CLIENT_STORE, normalizeClient(client));
    if (Array.isArray(payload.projects)) for (const project of payload.projects) await put(PROJECT_STORE, normalizeProject(project));
    if (Array.isArray(payload.contacts)) for (const contact of payload.contacts) await put(CONTACT_STORE, normalizeContact(contact));
    if (Array.isArray(payload.containers)) for (const container of payload.containers) await put(CONTAINER_STORE, normalizeContainer(container));
    if (Array.isArray(payload.materials)) for (const material of payload.materials) await put(MATERIAL_STORE, normalizeMaterial(material));
    if (Array.isArray(payload.settings)) for (const setting of payload.settings) await put(SETTINGS_STORE, setting);

    await loadAll();
    if (currentUser) currentUser = users.find((user) => user.id === currentUser.id) || currentUser;
    render();
  } catch {
    alert(t("Falha ao importar backup. Verifique o arquivo."));
  } finally {
    importInput.value = "";
  }
});

syncConfigForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("sync")) return;

  const data = new FormData(syncConfigForm);
  syncConfig = {
    ...syncConfig,
    id: "syncConfig",
    supabaseUrl: data.get("supabaseUrl")?.toString().trim(),
    supabaseAnonKey: data.get("supabaseAnonKey")?.toString().trim(),
    tenant: data.get("tenant")?.toString().trim(),
    autoSync: data.get("autoSync")?.toString() === "true",
  };

  await saveSetting(syncConfig);
  updateSyncStatus(t("Configuracao salva."));
});

pushSyncBtn.addEventListener("click", () => pushCloud());
pullSyncBtn.addEventListener("click", () => pullCloud());

function updateConnectivity() {
  const online = navigator.onLine;
  connectionBadge.textContent = online ? t("Online") : t("Offline");
  connectionBadge.style.borderColor = online ? "#94c5a8" : "#e7b6b6";
  connectionBadge.style.background = online ? "#f4fcf7" : "#fff5f5";

  if (online) queueAutoSync();
}

window.addEventListener("online", updateConnectivity);
window.addEventListener("offline", updateConnectivity);

window.addEventListener("beforeinstallprompt", (event) => {
  event.preventDefault();
  deferredPrompt = event;
  installBtn.classList.remove("hidden");
});

installBtn.addEventListener("click", async () => {
  if (!deferredPrompt) return;
  deferredPrompt.prompt();
  await deferredPrompt.userChoice;
  deferredPrompt = null;
  installBtn.classList.add("hidden");
});

languageSelect?.addEventListener("change", () => {
  setLanguage(languageSelect.value);
});

editProfileBtn?.addEventListener("click", () => {
  if (!currentUser) return;
  const returnView = currentView === "userEdit" ? "home" : currentView;
  openUserEdit(currentUser.id, returnView);
});

openUserEditCoiFileBtn?.addEventListener("click", () => {
  const targetUser = users.find((entry) => entry.id === editingUserId);
  const dataUrl = targetUser?.contractorCoiFile?.dataUrl;
  if (!dataUrl) return;
  window.open(dataUrl, "_blank", "noopener");
});

openUserAdminCoiFileBtn?.addEventListener("click", () => {
  const targetUser = users.find((entry) => entry.id === adminEditingUserId);
  const dataUrl = targetUser?.contractorCoiFile?.dataUrl;
  if (!dataUrl) return;
  window.open(dataUrl, "_blank", "noopener");
});

[signupPhoneInput, signupCellPhoneInput, userEditPhoneInput, userEditCellPhoneInput, userAdminPhoneInput, userAdminCellPhoneInput].forEach((input) => {
  if (!input) return;
  input.addEventListener("input", () => {
    input.value = maskPhoneValue(input.value);
  });
  input.addEventListener("blur", () => {
    input.value = maskPhoneValue(input.value);
  });
});

toggleSubcontractorExtras("signup", signupEmploymentTypeSelect?.value || "tag");
toggleSubcontractorExtras("edit", userEditEmploymentTypeSelect?.value || "tag");
toggleSubcontractorExtras("admin", userAdminEmploymentTypeSelect?.value || "tag");

appMain?.addEventListener("click", (event) => {
  const sendCoiBtn = event.target.closest("[data-send-coi-reminder]");
  if (sendCoiBtn) {
    event.preventDefault();
    void dispatchCoiReminder(sendCoiBtn.dataset.sendCoiReminder || "");
    return;
  }

  const openCoiBtn = event.target.closest("[data-open-coi-file]");
  if (openCoiBtn) {
    event.preventDefault();
    const targetUser = users.find((entry) => entry.id === openCoiBtn.dataset.openCoiFile);
    const dataUrl = targetUser?.contractorCoiFile?.dataUrl;
    if (!dataUrl) return;
    window.open(dataUrl, "_blank", "noopener");
    return;
  }

  const actionTrigger = event.target.closest("[data-nav-action]");
  if (actionTrigger) {
    event.preventDefault();
    handleQuickNavAction(actionTrigger.dataset.navAction || "");
    return;
  }

  const trigger = event.target.closest("[data-nav-view]");
  if (!trigger) return;
  event.preventDefault();
  handleQuickNavAction(trigger.dataset.navView || "home");
});

projectSectorPanel?.addEventListener("click", (event) => {
  const trigger = event.target.closest("[data-project-sector]");
  if (!trigger) return;
  event.preventDefault();
  const sector = trigger.dataset.projectSector || "";
  projectsViewMode = "operations";
  if (currentView !== "projects") setView("projects");
  setProjectSector(sector);
});

window.addEventListener("hashchange", () => {
  if (!currentUser) return;
  const target = viewFromHash();
  if (target === currentView) return;
  setView(target, { updateHash: false });
});

async function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    try {
      await navigator.serviceWorker.register("sw.js", { updateViaCache: "none" });
    } catch (error) {
      console.error("Failed to register Service Worker", error);
    }
  }
}

(async function init() {
  try {
    setLanguage(DEFAULT_LANG, { rerender: false });
    db = await openDB();
    await loadAll();
    await ensureDeveloperRolePresence();
    await tryRestoreSession();
    updateConnectivity();
    renderAuth();
    await registerServiceWorker();
  } catch (error) {
    console.error("Failed to start app", error);
    showBootError(t("Nao foi possivel iniciar o banco local. Feche outras abas deste app e recarregue a pagina."));
  }
})();
