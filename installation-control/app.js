const DB_NAME = "cabinets-control-db";
const DB_VERSION = 8;
const UNIT_STORE = "units";
const PHOTO_STORE = "photos";
const USER_STORE = "users";
const SETTINGS_STORE = "settings";
const CLIENT_STORE = "clients";
const PROJECT_STORE = "projects";
const CONTACT_STORE = "contacts";
const CONTRACT_STORE = "contracts";
const CONTAINER_STORE = "containers";
const MATERIAL_STORE = "materials";
const DELIVERY_SKU_STORE = "deliverySkuItems";
const TRASH_STORE = "deletedRecords";
const SESSION_KEY = "cc-session-user-id";
const APP_AUDIT_KEY = "cc-app-audit-log";
const LOCAL_BACKUP_META_KEY = "cc-last-local-backup-at";
const APP_AUDIT_MAX = 3000;
const DELETE_RETENTION_MS = 48 * 60 * 60 * 1000;
const RESTORE_WINDOW_LABEL = "48 hours";

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
  "project-manager": "Project Manager",
  warehouse: "Warehouse",
  transport: "Transport",
  distribution: "Distribution",
  foreman: "Foreman",
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
  distribution: ["distribution"],
  foreman: ["distribution", "installation"],
  qa: ["quality"],
  installer: ["distribution", "installation"],
};

const DEFAULT_SYNC = {
  supabaseUrl: "",
  supabaseAnonKey: "",
  tenant: "",
  autoSync: false,
  lastSyncAt: null,
  lastPullCursor: null,
};

const PRESET_SYNC = (() => {
  const runtime = window.CABINETS_SYNC || {};
  return {
    supabaseUrl: String(runtime.supabaseUrl || "").trim(),
    supabaseAnonKey: String(runtime.supabaseAnonKey || "").trim(),
    tenant: String(runtime.tenant || "").trim(),
    autoSync: typeof runtime.autoSync === "boolean" ? runtime.autoSync : true,
  };
})();

function hasPresetSyncConfig() {
  return Boolean(PRESET_SYNC.supabaseUrl && PRESET_SYNC.supabaseAnonKey && PRESET_SYNC.tenant);
}

function applyPresetSyncConfig(config) {
  if (!hasPresetSyncConfig()) return config;
  return {
    ...config,
    supabaseUrl: PRESET_SYNC.supabaseUrl,
    supabaseAnonKey: PRESET_SYNC.supabaseAnonKey,
    tenant: PRESET_SYNC.tenant,
    autoSync: PRESET_SYNC.autoSync,
  };
}

const DEFAULT_LANG = "en";
const SUPPORTED_LANGS = ["en"];
const PRIMARY_DEVELOPER_NAME = "leandro baptista";
const PRIMARY_DEVELOPER_PASSWORD = "0000";
const PROJECT_SECTORS = ["fabrica", "warehouse", "delivery", "distribuicao", "instalacao", "punchlist"];
const PROJECT_SCOPE_LABELS = {
  kitchens: "Kitchens",
  vanities: "Vanities",
  "med-cabinets": "Med Cabinets",
  "wood-floor": "Wood Floor",
  millworking: "Millworking",
  tile: "Tile",
  painting: "Painting",
};
const QR_WORKFLOW_STAGES = ["warehouse", "delivery", "distribution", "installation", "punchlist"];
const QR_WORKFLOW_STAGE_LABELS = {
  warehouse: "Warehouse",
  delivery: "Delivery",
  distribution: "Distribution",
  installation: "Installation",
  punchlist: "Punch List",
};
const QR_WORKFLOW_STATUS_OPTIONS = ["generated", "in-progress", "completed", "hold", "issue"];
const UNIT_DEFAULT_QR_MATERIALS = [
  "Kitchen cabinet module",
  "Fridge panel",
  "Island panel",
  "DW panel",
  "Side panel",
  "Wall filler",
  "Base filler",
  "Handle",
  "Pin shelf",
  "Door bumper",
  "Med cabinet",
  "Vanity piece",
  "Wood floor pallet",
  "Tile box",
  "Wood trim",
  "Door",
  "Self-level pallet",
];
const DEFAULT_AVATAR_BY_GENDER = {
  male: "avatar-male.svg",
  female: "avatar-female.svg",
  unspecified: "",
};
const DELIVERY_SKU_SEED = Array.isArray(window.DELIVERY_SKU_SEED) ? window.DELIVERY_SKU_SEED : [];
const DELIVERY_IMPORT_FIELD_ALIASES = {
  sku: ["sku", "item", "itemcode", "materialcode", "code", "productcode"],
  description: ["description", "desc", "itemdescription", "material", "product", "name"],
  finish: ["finish", "color", "colour", "variant", "model", "style"],
  qty: ["qty", "quantity", "totalqty", "count", "pieces", "pcs", "units"],
  category: ["category", "type", "group", "family"],
  unit: ["unit", "uom", "measure"],
  kitchenType: ["kitchentype", "kitchen", "kitchentag", "kitchencategory"],
  clientName: ["client", "clientname", "customer"],
  projectName: ["project", "projectname", "job", "jobname"],
  unitCode: ["unitcode", "unit", "apartment", "apt", "locationunit"],
  source: ["source", "container", "list", "batch", "origin"],
};
const DELIVERY_IMPORT_SUPPORTED_HINT =
  "Supported auto-import: CSV, TSV, TXT, JSON, XLS, XLSX, PDF, JPG/PNG, DOCX. Legacy .DOC should be converted to DOCX.";
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
    "project-manager": "Project Manager",
    warehouse: "Warehouse",
    transport: "Transport",
    distribution: "Distribution",
    foreman: "Foreman",
    qa: "QA",
    installer: "Installer",
    visitor: "Visitor",
  },
  pt: {
    developer: "Desenvolvedor",
    admin: "Admin",
    "project-manager": "Project Manager",
    warehouse: "Warehouse",
    transport: "Transport",
    distribution: "Distribution",
    foreman: "Foreman",
    qa: "QA",
    installer: "Installer",
    visitor: "Visitante",
  },
  es: {
    developer: "Desarrollador",
    admin: "Admin",
    "project-manager": "Project Manager",
    warehouse: "Almacen",
    transport: "Transporte",
    distribution: "Distribution",
    foreman: "Foreman",
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
    Funcao: "Job Title",
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
    Perfil: "Access Profile",
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
      "Visitor access: read-only and reports. Access profile approval by admin/developer.",
    "Acesso visitante: apenas consulta e relatorios. Aprovacao de funcoes somente por admin.":
      "Visitor access: read-only and reports. Access profile approval is admin-only.",
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
  [/\\bPerfil\\b/g, "Access Profile"],
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
  [/\\bFuncao\\b/g, "Job Title"],
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

function userAccessProfile(user) {
  return String(user?.accessProfile || user?.role || "visitor").trim() || "visitor";
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
    return `User: ${lineUser[1]} | Access Profile: ${lineUser[2]}`;
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
let contracts = [];
let containers = [];
let materials = [];
let deliverySkuItems = [];
let trashRecords = [];
let deliveryImportDraft = null;
let ocrImportDraft = null;
let syncConfig = { ...DEFAULT_SYNC };
let photosByUnit = new Map();
let currentUser = null;
let deferredPrompt = null;
let autoSyncTimer = null;
let autoPullTimer = null;
let autoPullInFlight = false;
let autoPullBlockedUntil = 0;
let clientsNavHoverTimer = null;
const AUTO_PULL_INTERVAL_MS = 30 * 1000;
const AUTO_PULL_SUBMIT_GRACE_MS = 12 * 1000;
const AUTO_PULL_KINDS = ["unit", "user", "client", "project", "contact", "contract", "container", "material", "deliverySku", "trash"];
let currentView = "home";
let editingUserId = "";
let userEditReturnView = "home";
let adminEditingUserId = "";
let userAdminFormOpen = false;
let usersSubView = "directory";
let lastQrLookupCode = "";
let currentProjectSector = "fabrica";
let usersViewFilter = "all";
let hasAutoDispatchedCoiReminder = false;
let selectedClientId = "";
let selectedProjectId = "";
let selectedContractId = "";
let selectedClientDetailsProjectId = "";
let keepClientFormBlank = false;
let clientSearchQuery = "";
let projectsViewMode = "overview";
let clientsWorkspaceMode = "hub";
let manufactureSubView = "schedule";
let warehouseSubView = "operations";
let appAuditLog = loadAppAuditLog();
let projectScopeExtrasDraft = [];
let projectChecklistExtrasDraft = [];
let xlsxLoaderPromise = null;
let tesseractLoaderPromise = null;
let pdfjsLoaderPromise = null;
let mammothLoaderPromise = null;
let cameraScanStream = null;
let cameraScanFrameId = 0;
let cameraScanResolver = null;
let cameraScanCanvas = null;
let jsQrLoaderPromise = null;
let scanFeedbackAudioCtx = null;
let legacyUserRoleKeyDetected = false;

const authView = document.getElementById("authView");
const bootErrorPanel = document.getElementById("bootErrorPanel");
const bootErrorMessage = document.getElementById("bootErrorMessage");
const setupPanel = document.getElementById("setupPanel");
const loginPanel = document.getElementById("loginPanel");
const signupPanel = document.getElementById("signupPanel");
const setupForm = document.getElementById("setupForm");
const loginForm = document.getElementById("loginForm");
const signupForm = document.getElementById("signupForm");
const signupGenderSelect = document.getElementById("signupGender");
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
const clientsNavGroup = document.querySelector('.nav-group-clients[data-nav-group="clients"]');
const ocrImporterPanel = document.getElementById("ocrImporterPanel");

const masterDataPanel = document.getElementById("masterDataPanel");
const clientsSubpanel = document.getElementById("clientsSubpanel");
const projectsSubpanel = document.getElementById("projectsSubpanel");
const contactsSubpanel = document.getElementById("contactsSubpanel");
const contractsSubpanel = document.getElementById("contractsSubpanel");
const materialsSubpanel = document.getElementById("materialsSubpanel");
const manufactureScheduleView = document.getElementById("manufactureScheduleView");
const manufactureCatalogView = document.getElementById("manufactureCatalogView");
const manufactureSolicitationView = document.getElementById("manufactureSolicitationView");
const materialsReplacementSummaryTable = document.getElementById("materialsReplacementSummaryTable");
const materialsFactoryMissingSummaryTable = document.getElementById("materialsFactoryMissingSummaryTable");
const clientForm = document.getElementById("clientForm");
const projectForm = document.getElementById("projectForm");
const contactForm = document.getElementById("contactForm");
const contractForm = document.getElementById("contractForm");
const materialForm = document.getElementById("materialForm");
const clientsSummaryScroll = document.getElementById("clientsSummaryScroll");
const clientsCreateView = document.getElementById("clientsCreateView");
const clientDetailsView = document.getElementById("clientDetailsView");
const clientsHubView = document.getElementById("clientsHubView");
const clientsHubList = document.getElementById("clientsHubList");
const clientDetailsBackBtn = document.getElementById("clientDetailsBackBtn");
const clientDetailsPanel = document.getElementById("clientDetailsPanel");
const clientProjectsList = document.getElementById("clientProjectsList");
const clientProjectDetailsTable = document.getElementById("clientProjectDetailsTable");
const clientAdminTarget = document.getElementById("clientAdminTarget");
const clientLogoPreview = document.getElementById("clientLogoPreview");
const clientFormNewBtn = document.getElementById("clientFormNewBtn");
const clientFormDeleteBtn = document.getElementById("clientFormDeleteBtn");
const goProjectsFromClientBtn = document.getElementById("goProjectsFromClientBtn");
const goPeopleFromClientBtn = document.getElementById("goPeopleFromClientBtn");
const clientSearchInput = document.getElementById("clientSearchInput");
const projectsTable = document.getElementById("projectsTable");
const contactsTable = document.getElementById("contactsTable");
const contractsTable = document.getElementById("contractsTable");
const materialsTable = document.getElementById("materialsTable");
const projectClientSelect = document.getElementById("projectClientSelect");
const projectAdminTarget = document.getElementById("projectAdminTarget");
const projectFormNewBtn = document.getElementById("projectFormNewBtn");
const projectFormDeleteBtn = document.getElementById("projectFormDeleteBtn");
const projectScopeExtraInput = document.getElementById("projectScopeExtraInput");
const projectScopeExtraAddBtn = document.getElementById("projectScopeExtraAddBtn");
const projectScopeExtraList = document.getElementById("projectScopeExtraList");
const projectScopeExtrasJson = document.getElementById("projectScopeExtrasJson");
const projectChecklistExtraInput = document.getElementById("projectChecklistExtraInput");
const projectChecklistExtraAddBtn = document.getElementById("projectChecklistExtraAddBtn");
const projectChecklistExtraList = document.getElementById("projectChecklistExtraList");
const projectChecklistExtrasJson = document.getElementById("projectChecklistExtrasJson");
const contactClientSelect = document.getElementById("contactClientSelect");
const contactProjectSelect = document.getElementById("contactProjectSelect");
const contractClientSelect = document.getElementById("contractClientSelect");
const contractProjectSelect = document.getElementById("contractProjectSelect");
const contractAdminTarget = document.getElementById("contractAdminTarget");
const contractFormNewBtn = document.getElementById("contractFormNewBtn");
const contractFormDeleteBtn = document.getElementById("contractFormDeleteBtn");

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
const unitScopeWorkInput = unitForm?.querySelector('textarea[name="scopeWork"]');
const unitProjectHint = document.getElementById("unitProjectHint");
const warehouseProjectSummary = document.getElementById("warehouseProjectSummary");
const deliveryInventoryPanel = document.getElementById("deliveryInventoryPanel");
const deliveryInventoryCatalogBlock = document.getElementById("deliveryInventoryCatalogBlock");
const deliveryInventoryScanBlock = document.getElementById("deliveryInventoryScanBlock");
const deliveryInventorySeedBtn = document.getElementById("deliveryInventorySeedBtn");
const deliveryInventoryResetBtn = document.getElementById("deliveryInventoryResetBtn");
const deliveryInventoryTable = document.getElementById("deliveryInventoryTable");
const deliveryInventoryStatus = document.getElementById("deliveryInventoryStatus");
const deliveryAddForm = document.getElementById("deliveryAddForm");
const deliveryAddSkuInput = document.getElementById("deliveryAddSkuInput");
const deliveryAddDescriptionInput = document.getElementById("deliveryAddDescriptionInput");
const deliveryAddFinishInput = document.getElementById("deliveryAddFinishInput");
const deliveryAddQtyInput = document.getElementById("deliveryAddQtyInput");
const deliveryScanForm = document.getElementById("deliveryScanForm");
const deliveryScanQrInput = document.getElementById("deliveryScanQrInput");
const deliveryScanBtn = document.getElementById("deliveryScanBtn");
const deliveryScanQtyInput = document.getElementById("deliveryScanQtyInput");
const deliveryAssignClientSelect = document.getElementById("deliveryAssignClientSelect");
const deliveryAssignProjectSelect = document.getElementById("deliveryAssignProjectSelect");
const deliveryAssignUnitSelect = document.getElementById("deliveryAssignUnitSelect");
const deliveryDestinationNoteInput = document.getElementById("deliveryDestinationNoteInput");
const deliverySaveDestinationBtn = document.getElementById("deliverySaveDestinationBtn");
const deliveryImportForm = document.getElementById("deliveryImportForm");
const deliveryImportFileInput = document.getElementById("deliveryImportFileInput");
const deliveryImportTargetSelect = document.getElementById("deliveryImportTargetSelect");
const deliveryImportContainerWrap = document.getElementById("deliveryImportContainerWrap");
const deliveryImportContainerSelect = document.getElementById("deliveryImportContainerSelect");
const deliveryImportAnalyzeBtn = document.getElementById("deliveryImportAnalyzeBtn");
const deliveryImportApplyBtn = document.getElementById("deliveryImportApplyBtn");
const deliveryImportClearBtn = document.getElementById("deliveryImportClearBtn");
const deliveryImportStatus = document.getElementById("deliveryImportStatus");
const deliveryImportPreview = document.getElementById("deliveryImportPreview");
const ocrImportForm = document.getElementById("ocrImportForm");
const ocrImportFileInput = document.getElementById("ocrImportFileInput");
const ocrImportTargetSelect = document.getElementById("ocrImportTargetSelect");
const ocrImportContainerWrap = document.getElementById("ocrImportContainerWrap");
const ocrImportContainerSelect = document.getElementById("ocrImportContainerSelect");
const ocrImportAnalyzeBtn = document.getElementById("ocrImportAnalyzeBtn");
const ocrImportApplyBtn = document.getElementById("ocrImportApplyBtn");
const ocrImportClearBtn = document.getElementById("ocrImportClearBtn");
const ocrImportStatus = document.getElementById("ocrImportStatus");
const ocrImportTextPreview = document.getElementById("ocrImportTextPreview");
const ocrImportRowsPreview = document.getElementById("ocrImportRowsPreview");

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
const qrLookupScanBtn = document.getElementById("qrLookupScanBtn");
const qrFlowForm = document.getElementById("qrFlowForm");
const qrFlowCodeInput = document.getElementById("qrFlowCodeInput");
const qrFlowStageSelect = document.getElementById("qrFlowStageSelect");
const qrFlowStageCustomWrap = document.getElementById("qrFlowStageCustomWrap");
const qrFlowStageCustomInput = document.getElementById("qrFlowStageCustomInput");
const qrFlowStatusSelect = document.getElementById("qrFlowStatusSelect");
const qrFlowIssueTypeSelect = document.getElementById("qrFlowIssueTypeSelect");
const qrFlowUnitSelect = document.getElementById("qrFlowUnitSelect");
const qrFlowNoteInput = document.getElementById("qrFlowNoteInput");
const qrFlowScanBtn = document.getElementById("qrFlowScanBtn");
const qrFlowAccessHint = document.getElementById("qrFlowAccessHint");
const qrFlowResult = document.getElementById("qrFlowResult");

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
const usersRegistrationTabBtn = document.getElementById("usersRegistrationTabBtn");
const usersDirectoryTabBtn = document.getElementById("usersDirectoryTabBtn");
const usersDirectoryView = document.getElementById("usersDirectoryView");
const usersRegistrationView = document.getElementById("usersRegistrationView");
const userForm = document.getElementById("userForm");
const userFormPanel = document.getElementById("userFormPanel");
const usersNameRail = document.getElementById("usersNameRail");
const usersTable = document.getElementById("usersTable");
const userAdminTarget = document.getElementById("userAdminTarget");
const userIdCard = document.getElementById("userIdCard");
const userIdCardPhoto = document.getElementById("userIdCardPhoto");
const userIdCardFirstName = document.getElementById("userIdCardFirstName");
const userIdCardLastName = document.getElementById("userIdCardLastName");
const userIdCardRole = document.getElementById("userIdCardRole");
const userAdminPhotoPreview = document.getElementById("userAdminPhotoPreview");
const userAdminGenderSelect = document.getElementById("userAdminGender");
const userAdminEmploymentTypeSelect = document.getElementById("userAdminEmploymentType");
const userAdminPhoneInput = document.getElementById("userAdminPhone");
const userAdminCellPhoneInput = document.getElementById("userAdminCellPhone");
const userAdminCoiFileStatus = document.getElementById("userAdminCoiFileStatus");
const openUserAdminCoiFileBtn = document.getElementById("openUserAdminCoiFileBtn");
const userFormNewBtn = document.getElementById("userFormNewBtn");
const userDirectoryEditBtn = document.getElementById("userDirectoryEditBtn");
const userFormDeleteBtn = document.getElementById("userFormDeleteBtn");
const userFormCloseBtn = document.getElementById("userFormCloseBtn");
const userEditPanel = document.getElementById("userEditPanel");
const userEditForm = document.getElementById("userEditForm");
const userEditTarget = document.getElementById("userEditTarget");
const userEditPhotoPreview = document.getElementById("userEditPhotoPreview");
const userEditGenderSelect = document.getElementById("userEditGender");
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
const cameraScanModal = document.getElementById("cameraScanModal");
const cameraScanTitle = document.getElementById("cameraScanTitle");
const cameraScanVideo = document.getElementById("cameraScanVideo");
const cameraScanStatus = document.getElementById("cameraScanStatus");
const cameraScanCloseBtn = document.getElementById("cameraScanCloseBtn");
const cameraScanManualBtn = document.getElementById("cameraScanManualBtn");
const cameraScanPhotoBtn = document.getElementById("cameraScanPhotoBtn");
const cameraScanPhotoInput = document.getElementById("cameraScanPhotoInput");

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

function normalizeLookupText(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function formatBytes(bytes) {
  const value = Number(bytes || 0);
  if (!Number.isFinite(value) || value <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let size = value;
  let index = 0;
  while (size >= 1024 && index < units.length - 1) {
    size /= 1024;
    index += 1;
  }
  const precision = size >= 100 ? 0 : size >= 10 ? 1 : 2;
  return `${size.toFixed(precision)} ${units[index]}`;
}

function normalizeImportKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "");
}

function parseImportQty(value, fallback = 1) {
  const raw = String(value ?? "").trim();
  if (!raw) return fallback;
  const compact = raw.replace(/\s+/g, "");
  let normalized = compact.replace(/[^0-9,.\-]/g, "");
  if (normalized.includes(",") && normalized.includes(".")) {
    normalized = normalized.replace(/,/g, "");
  } else if (normalized.includes(",") && !normalized.includes(".")) {
    normalized = normalized.replace(",", ".");
  }
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(1, Math.trunc(parsed));
}

const DELIVERY_IMPORT_NORMALIZED_ALIASES = Object.fromEntries(
  Object.entries(DELIVERY_IMPORT_FIELD_ALIASES).map(([field, aliases]) => [field, aliases.map((entry) => normalizeImportKey(entry))])
);

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

function normalizeTextList(value) {
  if (Array.isArray(value)) {
    return value
      .map((entry) => String(entry || "").trim())
      .filter(Boolean);
  }
  return String(value || "")
    .split(/\r?\n|,/)
    .map((entry) => entry.trim())
    .filter(Boolean);
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
      if (!dbRef.objectStoreNames.contains(CONTRACT_STORE)) dbRef.createObjectStore(CONTRACT_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(CONTAINER_STORE)) dbRef.createObjectStore(CONTAINER_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(MATERIAL_STORE)) dbRef.createObjectStore(MATERIAL_STORE, { keyPath: "id" });
      if (!dbRef.objectStoreNames.contains(DELIVERY_SKU_STORE)) {
        dbRef.createObjectStore(DELIVERY_SKU_STORE, { keyPath: "id" });
      }
      if (!dbRef.objectStoreNames.contains(TRASH_STORE)) dbRef.createObjectStore(TRASH_STORE, { keyPath: "id" });
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
  setupPanel?.classList.add("hidden");
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

function deepClone(value) {
  if (value === undefined) return undefined;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

function trashExpiresAt(deletedAtIso) {
  const deletedAtTs = new Date(deletedAtIso || Date.now()).getTime();
  return new Date(deletedAtTs + DELETE_RETENTION_MS).toISOString();
}

function normalizeTrashRecord(record) {
  const deletedAt = record.deletedAt || new Date().toISOString();
  return {
    id: record.id || uid(),
    storeName: String(record.storeName || "").trim(),
    recordId: String(record.recordId || "").trim(),
    label: String(record.label || "").trim(),
    scope: String(record.scope || "-").trim() || "-",
    restoreType: String(record.restoreType || "record").trim() || "record",
    payload: deepClone(record.payload),
    context: deepClone(record.context) || {},
    relatedRecords: ensureArray(record.relatedRecords).map((entry) => ({
      storeName: String(entry?.storeName || "").trim(),
      payload: deepClone(entry?.payload),
    })),
    deletedByUserId: record.deletedByUserId || "",
    deletedByName: record.deletedByName || "",
    deletedAt,
    expiresAt: record.expiresAt || trashExpiresAt(deletedAt),
    updatedAt: record.updatedAt || deletedAt,
  };
}

function isTrashRecordExpired(record, nowTs = Date.now()) {
  const expiresAtTs = new Date(record?.expiresAt || 0).getTime();
  if (!Number.isFinite(expiresAtTs) || expiresAtTs <= 0) return true;
  return expiresAtTs <= nowTs;
}

function normalizeRecordForStore(storeName, payload) {
  if (!payload) return payload;
  if (storeName === UNIT_STORE) return normalizeUnit(payload);
  if (storeName === USER_STORE) return normalizeUser(payload);
  if (storeName === CLIENT_STORE) return normalizeClient(payload);
  if (storeName === PROJECT_STORE) return normalizeProject(payload);
  if (storeName === CONTACT_STORE) return normalizeContact(payload);
  if (storeName === CONTRACT_STORE) return normalizeContract(payload);
  if (storeName === CONTAINER_STORE) return normalizeContainer(payload);
  if (storeName === MATERIAL_STORE) return normalizeMaterial(payload);
  if (storeName === DELIVERY_SKU_STORE) return normalizeDeliverySkuItem(payload);
  if (storeName === TRASH_STORE) return normalizeTrashRecord(payload);
  return payload;
}

async function saveTrashRecord({
  storeName = "",
  recordId = "",
  label = "",
  scope = "-",
  payload = null,
  restoreType = "record",
  context = {},
  relatedRecords = [],
} = {}) {
  const record = normalizeTrashRecord({
    id: uid(),
    storeName,
    recordId,
    label,
    scope,
    restoreType,
    payload,
    context,
    relatedRecords,
    deletedByUserId: currentUser?.id || "",
    deletedByName: currentUser?.name || "System",
    deletedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });
  await put(TRASH_STORE, record);
  return record;
}

async function trashDeleteRecord(storeName, record, { label = "", scope = "-", relatedRecords = [] } = {}) {
  if (!record || !record.id) return false;
  await saveTrashRecord({
    storeName,
    recordId: record.id,
    label: label || `${storeName}:${record.id}`,
    scope,
    payload: record,
    restoreType: "record",
    relatedRecords,
  });
  await del(storeName, record.id);
  for (const related of relatedRecords) {
    const relatedStore = String(related?.storeName || "").trim();
    const relatedId = String(related?.payload?.id || "").trim();
    if (!relatedStore || !relatedId) continue;
    await del(relatedStore, relatedId);
  }
  return true;
}

function confirmDeleteAction(itemLabel = "this item", { restorable = true } = {}) {
  if (!restorable) {
    return window.confirm(`Are you sure you want to remove ${itemLabel}?`);
  }
  return window.confirm(`Are you sure you want to delete ${itemLabel}?\nYou can restore it from the hidden recovery area for up to ${RESTORE_WINDOW_LABEL}.`);
}

function normalizeUnit(unit) {
  const normalized = { ...unit };
  normalized.stages = normalized.stages || {};
  for (const stage of STAGES) {
    normalized.stages[stage.key] = normalized.stages[stage.key] || { done: false, at: null };
  }
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
  normalized.unitCode = normalized.unitCode || "";
  normalized.category = normalized.category || "";
  normalized.unitType = normalized.unitType || "";
  normalized.kitchenModel = normalized.kitchenModel || "";
  normalized.checkItems = ensureArray(normalized.checkItems).map((item, index) => normalizeCheckItem(item, normalized, index));
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
    floorsCount: String(project.floorsCount || "").trim(),
    apartmentsCount: String(project.apartmentsCount || "").trim(),
    scopeCategories: ensureArray(project.scopeCategories).filter((entry) => Object.prototype.hasOwnProperty.call(PROJECT_SCOPE_LABELS, entry)),
    scopeExtras: normalizeTextList(project.scopeExtras),
    unitChecklistTemplate: normalizeTextList(project.unitChecklistTemplate),
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

function normalizeContract(contract) {
  return {
    id: contract.id,
    clientId: contract.clientId || "",
    projectId: contract.projectId || "",
    title: String(contract.title || "").trim(),
    contractCode: String(contract.contractCode || "").trim(),
    status: String(contract.status || "draft").trim() || "draft",
    signedDate: normalizeDateField(contract.signedDate),
    startDate: normalizeDateField(contract.startDate),
    endDate: normalizeDateField(contract.endDate),
    amount: String(contract.amount || "").trim(),
    notes: String(contract.notes || "").trim(),
    createdAt: contract.createdAt || new Date().toISOString(),
    updatedAt: contract.updatedAt || contract.createdAt || new Date().toISOString(),
  };
}

function normalizeGender(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "male" || raw === "man") return "male";
  if (raw === "female" || raw === "woman") return "female";
  return "unspecified";
}

function normalizeDateField(value) {
  const raw = String(value || "").trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(raw) ? raw : "";
}

function defaultAvatarForGender(gender) {
  const normalized = normalizeGender(gender);
  return DEFAULT_AVATAR_BY_GENDER[normalized] || "";
}

function userAvatarSrc(user) {
  if (user?.photoDataUrl) return user.photoDataUrl;
  return defaultAvatarForGender(user?.gender);
}

function setAvatarPreview(previewEl, src) {
  if (!previewEl) return;
  if (src) {
    previewEl.src = src;
    previewEl.classList.remove("hidden");
    return;
  }
  previewEl.removeAttribute("src");
  previewEl.classList.add("hidden");
}

function adminPreviewFallbackSrc() {
  const targetUser = adminEditingUserId ? users.find((entry) => entry.id === adminEditingUserId) : null;
  if (targetUser?.photoDataUrl) return targetUser.photoDataUrl;
  return defaultAvatarForGender(userAdminGenderSelect?.value || targetUser?.gender);
}

function userEditPreviewFallbackSrc() {
  const targetUser = editingUserId ? users.find((entry) => entry.id === editingUserId) : null;
  if (targetUser?.photoDataUrl) return targetUser.photoDataUrl;
  return defaultAvatarForGender(userEditGenderSelect?.value || targetUser?.gender);
}

async function refreshAdminPhotoPreview() {
  const photoInput = userForm?.querySelector('input[name="photo"]');
  const pickedFile = photoInput?.files?.[0];
  if (pickedFile) {
    setAvatarPreview(userAdminPhotoPreview, await fileToDataUrl(pickedFile));
    renderUserIdCard(adminEditingUserId ? users.find((entry) => entry.id === adminEditingUserId) || null : null);
    return;
  }
  setAvatarPreview(userAdminPhotoPreview, adminPreviewFallbackSrc());
  renderUserIdCard(adminEditingUserId ? users.find((entry) => entry.id === adminEditingUserId) || null : null);
}

async function refreshUserEditPhotoPreview() {
  const photoInput = userEditForm?.querySelector('input[name="photo"]');
  const pickedFile = photoInput?.files?.[0];
  if (pickedFile) {
    setAvatarPreview(userEditPhotoPreview, await fileToDataUrl(pickedFile));
    return;
  }
  setAvatarPreview(userEditPhotoPreview, userEditPreviewFallbackSrc());
}

function splitNameParts(user = null) {
  const firstFromForm = String(userForm?.firstName?.value || "").trim();
  const lastFromForm = String(userForm?.lastName?.value || "").trim();

  let first = firstFromForm || String(user?.firstName || "").trim();
  let last = lastFromForm || String(user?.lastName || "").trim();
  if (first || last) return { first, last };

  const fullName = String(user?.name || "").trim();
  if (!fullName) return { first: "", last: "" };
  const parts = fullName.split(/\s+/);
  first = parts.shift() || "";
  last = parts.join(" ");
  return { first, last };
}

function formPhotoPreviewSrc() {
  const src = userAdminPhotoPreview?.getAttribute("src") || "";
  return String(src).trim();
}

function renderUserIdCard(user = null) {
  if (!userIdCard || !userIdCardPhoto || !userIdCardFirstName || !userIdCardLastName || !userIdCardRole) return;
  const { first, last } = splitNameParts(user);
  const jobTitle = String(userForm?.jobTitle?.value || user?.jobTitle || "").trim();
  const gender = normalizeGender(userForm?.gender?.value || user?.gender || "unspecified");
  const photoSrc = formPhotoPreviewSrc() || userAvatarSrc(user) || defaultAvatarForGender(gender) || "avatar-neutral.svg";

  userIdCardPhoto.src = photoSrc;
  userIdCardFirstName.textContent = first || "First Name";
  userIdCardLastName.textContent = last || "Last Name";
  userIdCardRole.textContent = jobTitle || "Job Title";
}

function normalizeUser(user) {
  const firstName = user.firstName || "";
  const lastName = user.lastName || "";
  const composedName = `${firstName} ${lastName}`.trim();
  const accessProfile = userAccessProfile(user);
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
    birthDate: normalizeDateField(user.birthDate),
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
    gender: normalizeGender(user.gender),
    photoDataUrl: user.photoDataUrl || "",
    username: (user.username || "").toLowerCase(),
    passwordHash: user.passwordHash || "",
    legacyPassword: user.legacyPassword || user.password || "",
    accessProfile,
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

function normalizeFinishValue(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ")
    .replace(/M4 STAINED/g, "M4-STAINED")
    .replace(/M5 WHITE/g, "M5-WHITE")
    .replace(/\s*-\s*/g, "-");
}

function normalizeDeliverySkuValue(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .replace(/^ВТР15$/i, "BTP15")
    .replace(/^W1542 Lights SIDE FINISHED$/i, "W1542-Lights SIDE FINISHED")
    .replace(/^VDB12-3\/ LOCKABLE MIDDLE D$/i, "VDB12-3 / LOCKABLE MIDDLE DRAWER")
    .replace(/^VDB12-3 - LOCKABLE MIDDLE DRAWER$/i, "VDB12-3 / LOCKABLE MIDDLE DRAWER");
}

function deliverySkuKey(sku, finish) {
  return `${normalizeDeliverySkuValue(sku)}||${normalizeFinishValue(finish) || "-"}`;
}

function makeDeliverySkuQrCode(item) {
  const skuToken = qrToken(item?.sku, 16) || "SKU";
  const finishToken = qrToken(item?.finish, 14) || "FINISH";
  const itemToken = qrToken(item?.id, 8) || qrToken(uid(), 8);
  return `DSKU|${skuToken}|${finishToken}|${itemToken}`;
}

function deliverySkuAvailableQty(item) {
  const totalQty = Math.max(0, Math.trunc(toNumber(item?.totalQty)));
  const scannedQty = Math.max(0, Math.trunc(toNumber(item?.scannedQty)));
  return Math.max(0, totalQty - scannedQty);
}

function normalizeDeliverySkuItem(item) {
  const normalizedSku = normalizeDeliverySkuValue(item.sku || item.code || "");
  const normalizedFinish = normalizeFinishValue(item.finish || item.variant || "");
  const totalQty = Math.max(0, Math.trunc(toNumber(item.totalQty || item.qty || 0)));
  const scannedQty = Math.min(totalQty, Math.max(0, Math.trunc(toNumber(item.scannedQty || 0))));
  const sourceRefs = ensureArray(item.sourceRefs).map((entry) => ({
    source: String(entry?.source || "").trim() || "Unknown source",
    qty: Math.max(0, Math.trunc(toNumber(entry?.qty || 0))),
  }));
  const scanLog = ensureArray(item.scanLog).map((entry) => ({
    id: entry.id || uid(),
    qty: Math.max(1, Math.trunc(toNumber(entry.qty || 1))),
    scannedAt: entry.scannedAt || entry.createdAt || new Date().toISOString(),
    scannedByUserId: entry.scannedByUserId || "",
    scannedByName: entry.scannedByName || "",
    clientId: entry.clientId || "",
    projectId: entry.projectId || "",
    unitId: entry.unitId || "",
    destinationNote: entry.destinationNote || "",
  }));
  const destination = item.destination || {};
  const normalized = {
    id: item.id || uid(),
    sku: normalizedSku,
    description: String(item.description || normalizedSku || "").trim(),
    finish: normalizedFinish || "UNSPECIFIED",
    totalQty,
    scannedQty,
    qrCode: String(item.qrCode || "").trim(),
    clientId: item.clientId || destination.clientId || "",
    projectId: item.projectId || destination.projectId || "",
    unitId: item.unitId || destination.unitId || "",
    destinationNote: String(item.destinationNote || destination.note || "").trim(),
    sourceRefs,
    scanLog,
    createdAt: item.createdAt || new Date().toISOString(),
    updatedAt: item.updatedAt || item.createdAt || new Date().toISOString(),
  };
  if (!normalized.qrCode) normalized.qrCode = makeDeliverySkuQrCode(normalized);
  return normalized;
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
    flowStage: normalizeWorkflowStage(item.flowStage || ""),
    flowStatus: normalizeWorkflowStatus(item.flowStatus || item.status || "generated"),
    flowUpdatedAt: item.flowUpdatedAt || item.updatedAt || item.createdAt || new Date().toISOString(),
    flowEvents: ensureArray(item.flowEvents).map((event) => ({
      id: event.id || uid(),
      at: event.at || event.createdAt || new Date().toISOString(),
      byUserId: event.byUserId || "",
      byName: event.byName || "",
      byAccessProfile: event.byAccessProfile || event.byRole || "",
      bySector: event.bySector || "",
      stage: normalizeWorkflowStage(event.stage || ""),
      status: normalizeWorkflowStatus(event.status || "in-progress"),
      issueType: event.issueType || "",
      note: event.note || "",
      unitId: event.unitId || "",
      unitLabel: event.unitLabel || "",
      source: event.source || "workflow",
    })),
    updatedAt: item.updatedAt || item.createdAt || new Date().toISOString(),
    createdAt: item.createdAt || new Date().toISOString(),
  }));
  qrItems.forEach((entry) => ensureQrItemFlowDefaults(entry));
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
  const [unitRows, photoRows, userRows, settingsRows, clientRows, projectRows, contactRows, contractRows, containerRows, materialRows, deliverySkuRows, trashRows] =
    await Promise.all([
      getAll(UNIT_STORE),
      getAll(PHOTO_STORE),
      getAll(USER_STORE),
      getAll(SETTINGS_STORE),
      getAll(CLIENT_STORE),
      getAll(PROJECT_STORE),
      getAll(CONTACT_STORE),
      getAll(CONTRACT_STORE),
      getAll(CONTAINER_STORE),
      getAll(MATERIAL_STORE),
      getAll(DELIVERY_SKU_STORE),
      getAll(TRASH_STORE),
    ]);

  legacyUserRoleKeyDetected = userRows.some((row) => Object.prototype.hasOwnProperty.call(row || {}, "role"));
  units = unitRows.map(normalizeUnit).sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
  photos = photoRows;
  users = userRows.map(normalizeUser).sort((a, b) => a.username.localeCompare(b.username));
  clients = clientRows.map(normalizeClient).sort((a, b) => a.name.localeCompare(b.name));
  projects = projectRows.map(normalizeProject).sort((a, b) => a.name.localeCompare(b.name));
  contacts = contactRows.map(normalizeContact).sort((a, b) => a.name.localeCompare(b.name));
  contracts = contractRows
    .map(normalizeContract)
    .sort((a, b) => a.title.localeCompare(b.title) || a.contractCode.localeCompare(b.contractCode));
  containers = containerRows.map(normalizeContainer).sort((a, b) => {
    const ad = a.etaDate ? new Date(a.etaDate).getTime() : 0;
    const bd = b.etaDate ? new Date(b.etaDate).getTime() : 0;
    return ad - bd;
  });
  materials = materialRows.map(normalizeMaterial).sort((a, b) => a.sku.localeCompare(b.sku));
  deliverySkuItems = deliverySkuRows
    .map(normalizeDeliverySkuItem)
    .sort((a, b) => a.sku.localeCompare(b.sku) || a.finish.localeCompare(b.finish));
  const normalizedTrash = trashRows.map(normalizeTrashRecord);
  const expiredTrashIds = normalizedTrash.filter((row) => isTrashRecordExpired(row)).map((row) => row.id);
  trashRecords = normalizedTrash
    .filter((row) => !expiredTrashIds.includes(row.id))
    .sort((a, b) => new Date(b.deletedAt || 0).getTime() - new Date(a.deletedAt || 0).getTime());
  if (expiredTrashIds.length) {
    await Promise.all(expiredTrashIds.map((id) => del(TRASH_STORE, id)));
  }

  const savedSync = settingsRows.find((row) => row.id === "syncConfig") || {};
  syncConfig = applyPresetSyncConfig({
    ...DEFAULT_SYNC,
    ...savedSync,
    autoSync: Boolean(savedSync.autoSync),
  });

  photosByUnit = new Map();
  for (const photo of photos) {
    if (!photosByUnit.has(photo.unitId)) photosByUnit.set(photo.unitId, []);
    photosByUnit.get(photo.unitId).push(photo);
  }
}

async function migrateLegacyUserRoleKey() {
  if (!legacyUserRoleKeyDetected || !users.length) return;
  const normalizedUsers = users.map((user) => normalizeUser(user));
  await Promise.all(normalizedUsers.map((user) => put(USER_STORE, user)));
  legacyUserRoleKeyDetected = false;
  await loadAll();
  if (syncEndpoint()) {
    await pushCloud({ silent: true, force: true, kinds: ["user"] });
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

  if (!preferredDeveloper) return;

  const now = new Date().toISOString();
  const developerPasswordHash = await hashPassword(PRIMARY_DEVELOPER_PASSWORD);
  const usersToUpdate = [];
  let preferredDeveloperDraft = { ...preferredDeveloper };
  let preferredDeveloperChanged = false;

  if (userAccessProfile(preferredDeveloper) !== "developer") {
    preferredDeveloperDraft.accessProfile = "developer";
    preferredDeveloperChanged = true;
  }

  if (!userPasswordMatches(preferredDeveloper, PRIMARY_DEVELOPER_PASSWORD, developerPasswordHash)) {
    preferredDeveloperDraft.passwordHash = developerPasswordHash;
    preferredDeveloperDraft.legacyPassword = "";
    preferredDeveloperDraft.password = "";
    preferredDeveloperChanged = true;
  }

  if (preferredDeveloperChanged) {
    preferredDeveloperDraft.updatedAt = now;
    usersToUpdate.push(normalizeUser(preferredDeveloperDraft));
  }

  users
    .filter((user) => user.id !== preferredDeveloper.id && userAccessProfile(user) === "developer")
    .forEach((user) => {
      usersToUpdate.push(
        normalizeUser({
          ...user,
          accessProfile: "admin",
          updatedAt: now,
        })
      );
    });

  if (!usersToUpdate.length) return;
  await Promise.all(usersToUpdate.map((user) => put(USER_STORE, user)));
  await loadAll();
  await pushCloud({ silent: true, force: true, kinds: ["user"] });
}

function clientById(id) {
  return clients.find((client) => client.id === id);
}

function projectById(id) {
  return projects.find((project) => project.id === id);
}

function contractById(id) {
  return contracts.find((contract) => contract.id === id);
}

function materialById(id) {
  return materials.find((material) => material.id === id);
}

function deliverySkuById(id) {
  return deliverySkuItems.find((entry) => entry.id === id);
}

function deliverySkuByQrCode(qrCode) {
  const normalized = String(qrCode || "").trim();
  if (!normalized) return null;
  return deliverySkuItems.find((entry) => entry.qrCode === normalized) || null;
}

function trashStoreLabel(storeName) {
  if (storeName === UNIT_STORE) return "Units";
  if (storeName === PHOTO_STORE) return "Photos";
  if (storeName === USER_STORE) return "Users";
  if (storeName === CLIENT_STORE) return "Clients";
  if (storeName === PROJECT_STORE) return "Projects";
  if (storeName === CONTACT_STORE) return "People in Project";
  if (storeName === CONTRACT_STORE) return "Contracts";
  if (storeName === CONTAINER_STORE) return "Containers";
  if (storeName === MATERIAL_STORE) return "Materials";
  if (storeName === DELIVERY_SKU_STORE) return "Delivery Inventory";
  return storeName || "Record";
}

function formatTrashTimeLeft(entry) {
  const expiresAtTs = new Date(entry?.expiresAt || 0).getTime();
  if (!Number.isFinite(expiresAtTs) || expiresAtTs <= 0) return "Expired";
  const remainingMs = expiresAtTs - Date.now();
  if (remainingMs <= 0) return "Expired";
  const totalMinutes = Math.ceil(remainingMs / (60 * 1000));
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  if (hours <= 0) return `${minutes}m`;
  if (hours < 48) return `${hours}h ${minutes}m`;
  const days = Math.floor(hours / 24);
  const remHours = hours % 24;
  return `${days}d ${remHours}h`;
}

async function purgeExpiredTrashRecords() {
  const expiredIds = trashRecords.filter((entry) => isTrashRecordExpired(entry)).map((entry) => entry.id);
  if (!expiredIds.length) return 0;
  await Promise.all(expiredIds.map((id) => del(TRASH_STORE, id)));
  trashRecords = trashRecords.filter((entry) => !expiredIds.includes(entry.id));
  return expiredIds.length;
}

async function permanentlyDeleteTrashRecord(trashId, options = {}) {
  const target = trashRecords.find((entry) => entry.id === trashId);
  if (!target) return false;
  await del(TRASH_STORE, target.id);
  trashRecords = trashRecords.filter((entry) => entry.id !== target.id);
  if (!options.silentAudit) {
    const reasonSuffix = options.reason ? ` (${options.reason})` : "";
    pushEntityAudit(
      "Deleted Records",
      "purged",
      `${target.label || target.recordId || target.id}${reasonSuffix}`,
      target.scope || "trash"
    );
  }
  return true;
}

async function restoreTrashRecord(trashId) {
  const target = trashRecords.find((entry) => entry.id === trashId);
  if (!target) {
    alert("Deleted record not found.");
    return false;
  }
  if (isTrashRecordExpired(target)) {
    await permanentlyDeleteTrashRecord(target.id, { reason: "expired" });
    await loadAll();
    render();
    queueAutoSync();
    alert("This deleted record has expired and can no longer be restored.");
    return false;
  }

  const now = new Date().toISOString();
  const restoreType = String(target.restoreType || "record");

  if (restoreType === "record") {
    if (!target.storeName || !target.payload?.id) {
      alert("Restore data is invalid for this record.");
      return false;
    }
    await put(target.storeName, normalizeRecordForStore(target.storeName, target.payload));
    for (const related of ensureArray(target.relatedRecords)) {
      const relatedStore = String(related?.storeName || "").trim();
      const relatedPayload = related?.payload;
      if (!relatedStore || !relatedPayload?.id) continue;
      await put(relatedStore, normalizeRecordForStore(relatedStore, relatedPayload));
    }
  } else if (restoreType === "container-manifest-item") {
    const containerId = target.context?.containerId || "";
    const container = containers.find((entry) => entry.id === containerId);
    if (!container) {
      alert("Container no longer exists. Restore this container first.");
      return false;
    }
    if (target.payload?.id && !container.materialItems.some((item) => item.id === target.payload.id)) {
      container.materialItems.push(target.payload);
    }
    const qrItems = ensureArray(target.context?.qrItems);
    qrItems.forEach((qrItem) => {
      if (!qrItem?.id) return;
      if (!container.qrItems.some((entry) => entry.id === qrItem.id)) container.qrItems.push(qrItem);
    });
    container.updatedAt = now;
    await put(CONTAINER_STORE, normalizeContainer(container));
  } else if (restoreType === "unit-check-item") {
    const unitId = target.context?.unitId || "";
    const unit = units.find((entry) => entry.id === unitId);
    if (!unit) {
      alert("Unit no longer exists. Restore this unit first.");
      return false;
    }
    if (target.payload?.id && !unit.checkItems.some((entry) => entry.id === target.payload.id)) {
      unit.checkItems.push(target.payload);
    }
    unit.updatedAt = now;
    await put(UNIT_STORE, normalizeUnit(unit));
  } else if (restoreType === "unit-dispatch-task") {
    const unitId = target.context?.unitId || "";
    const unit = units.find((entry) => entry.id === unitId);
    if (!unit) {
      alert("Unit no longer exists. Restore this unit first.");
      return false;
    }
    if (target.payload?.id && !unit.dispatchTasks.some((entry) => entry.id === target.payload.id)) {
      unit.dispatchTasks.push(target.payload);
    }
    unit.updatedAt = now;
    await put(UNIT_STORE, normalizeUnit(unit));
  } else if (restoreType === "unit-site-receipt") {
    const unitId = target.context?.unitId || "";
    const unit = units.find((entry) => entry.id === unitId);
    if (!unit) {
      alert("Unit no longer exists. Restore this unit first.");
      return false;
    }
    if (target.payload?.id && !unit.siteReceipts.some((entry) => entry.id === target.payload.id)) {
      unit.siteReceipts.push(target.payload);
    }
    unit.updatedAt = now;
    await put(UNIT_STORE, normalizeUnit(unit));
  } else if (restoreType === "delivery-inventory-reset") {
    await writeDeliverySkuItems(ensureArray(target.payload), { replace: true });
  } else {
    alert("Restore type is not supported for this record.");
    return false;
  }

  await permanentlyDeleteTrashRecord(target.id, { silentAudit: true });
  pushEntityAudit("Deleted Records", "restored", target.label || target.recordId || target.id, target.scope || "trash");
  await loadAll();
  if (currentUser) currentUser = users.find((entry) => entry.id === currentUser.id) || currentUser;
  render();
  queueAutoSync();
  return true;
}

function deliverySkuMergeRows(...lists) {
  const map = new Map();
  lists.flat().forEach((raw) => {
    const normalized = normalizeDeliverySkuItem(raw);
    if (!normalized.sku) return;
    const key = deliverySkuKey(normalized.sku, normalized.finish);
    if (!map.has(key)) {
      map.set(key, {
        ...normalized,
        sourceRefs: [],
        scanLog: [],
      });
    }
    const existing = map.get(key);
    existing.totalQty += normalized.totalQty;
    existing.scannedQty = Math.min(existing.totalQty, existing.scannedQty + normalized.scannedQty);
    if (!existing.description && normalized.description) existing.description = normalized.description;
    if (!existing.clientId && normalized.clientId) existing.clientId = normalized.clientId;
    if (!existing.projectId && normalized.projectId) existing.projectId = normalized.projectId;
    if (!existing.unitId && normalized.unitId) existing.unitId = normalized.unitId;
    if (!existing.destinationNote && normalized.destinationNote) existing.destinationNote = normalized.destinationNote;
    if (!existing.qrCode && normalized.qrCode) existing.qrCode = normalized.qrCode;

    const sourceMap = new Map(existing.sourceRefs.map((entry) => [entry.source, entry.qty]));
    ensureArray(normalized.sourceRefs).forEach((entry) => {
      const source = String(entry.source || "").trim() || "Unknown source";
      sourceMap.set(source, (sourceMap.get(source) || 0) + Math.max(0, Math.trunc(toNumber(entry.qty || 0))));
    });
    existing.sourceRefs = Array.from(sourceMap.entries()).map(([source, qty]) => ({ source, qty }));
    existing.scanLog = [...existing.scanLog, ...ensureArray(normalized.scanLog)];
    existing.updatedAt = normalized.updatedAt || existing.updatedAt;
  });
  return Array.from(map.values())
    .map((entry) => {
      const normalized = normalizeDeliverySkuItem(entry);
      if (!normalized.qrCode) normalized.qrCode = makeDeliverySkuQrCode(normalized);
      return normalized;
    })
    .sort((a, b) => a.sku.localeCompare(b.sku) || a.finish.localeCompare(b.finish));
}

async function writeDeliverySkuItems(items, { replace = false } = {}) {
  const transaction = db.transaction([DELIVERY_SKU_STORE], "readwrite");
  const store = transaction.objectStore(DELIVERY_SKU_STORE);
  if (replace) store.clear();
  items.forEach((item) => {
    store.put(normalizeDeliverySkuItem(item));
  });
  await transactionDonePromise(transaction);
  await loadAll();
  render();
  queueAutoSync();
}

async function saveDeliverySkuItem(item) {
  await put(DELIVERY_SKU_STORE, normalizeDeliverySkuItem(item));
  await loadAll();
  render();
  queueAutoSync();
}

function deliverySkuSeedRows() {
  return DELIVERY_SKU_SEED.map((entry) =>
    normalizeDeliverySkuItem({
      id: uid(),
      sku: entry.sku,
      description: entry.description || entry.sku,
      finish: entry.finish,
      totalQty: Math.max(1, Math.trunc(toNumber(entry.totalQty || 1))),
      scannedQty: 0,
      sourceRefs: ensureArray(entry.sourceRefs),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    })
  );
}

function deliverySelectionProjects(clientId) {
  if (!clientId) return projects;
  return projects.filter((project) => project.clientId === clientId);
}

function deliverySelectionUnits(projectId) {
  if (!projectId) return units;
  return units.filter((unit) => unit.projectId === projectId);
}

function setDeliveryInventoryStatus(message = "", isError = false) {
  if (!deliveryInventoryStatus) return;
  deliveryInventoryStatus.textContent = message;
  deliveryInventoryStatus.style.color = isError ? "#b83232" : "var(--ink-soft)";
}

function setDeliveryImportStatus(message = "", isError = false) {
  if (!deliveryImportStatus) return;
  deliveryImportStatus.textContent = message;
  deliveryImportStatus.style.color = isError ? "#b83232" : "var(--ink-soft)";
}

function syncDeliveryImportApplyState() {
  if (!deliveryImportApplyBtn) return;
  const hasRows = Boolean(deliveryImportDraft?.supported && ensureArray(deliveryImportDraft?.rows).length > 0);
  const target = deliveryImportTargetSelect?.value || "deliverySku";
  deliveryImportApplyBtn.disabled = !hasRows || !canImportRowsByTarget(target);
}

function syncDeliveryImportContainerSelect() {
  if (!deliveryImportContainerSelect) return;
  const current = deliveryImportContainerSelect.value;
  const rows = containers.map((container) => {
    const clientName = clientById(container.clientId)?.name || "No client";
    const projectName = projectById(container.projectId)?.name || "No project";
    return `<option value="${escapeHtml(container.id)}">${escapeHtml(container.containerCode)} | ${escapeHtml(clientName)} | ${escapeHtml(projectName)}</option>`;
  });
  deliveryImportContainerSelect.innerHTML = `<option value="">Select a container</option>${rows.join("")}`;
  if (containers.some((entry) => entry.id === current)) deliveryImportContainerSelect.value = current;
}

function refreshDeliveryImportTargetUi() {
  const target = deliveryImportTargetSelect?.value || "deliverySku";
  const needsContainer = target === "containerManifest";
  deliveryImportContainerWrap?.classList.toggle("hidden", !needsContainer);
  if (deliveryImportApplyBtn) {
    if (target === "deliverySku") deliveryImportApplyBtn.textContent = "Import to Delivery inventory";
    else if (target === "containerManifest") deliveryImportApplyBtn.textContent = "Import to Container manifest";
    else deliveryImportApplyBtn.textContent = "Import to Material catalog";
  }
}

function clearDeliveryImportDraft({ clearFile = true, clearStatus = false } = {}) {
  deliveryImportDraft = null;
  if (clearFile && deliveryImportFileInput) deliveryImportFileInput.value = "";
  if (deliveryImportPreview) deliveryImportPreview.innerHTML = "";
  if (clearStatus) setDeliveryImportStatus("");
  syncDeliveryImportApplyState();
}

function detectImportFileKind(file) {
  const name = String(file?.name || "").toLowerCase();
  const type = String(file?.type || "").toLowerCase();
  if (name.endsWith(".csv") || type.includes("csv")) return { key: "csv", label: "CSV" };
  if (name.endsWith(".tsv")) return { key: "tsv", label: "TSV" };
  if (name.endsWith(".txt") || type.startsWith("text/")) return { key: "text", label: "Text" };
  if (name.endsWith(".json") || type.includes("json")) return { key: "json", label: "JSON" };
  if (
    name.endsWith(".xlsx") ||
    name.endsWith(".xls") ||
    type.includes("spreadsheet") ||
    type.includes("excel")
  ) {
    return { key: "spreadsheet", label: "Spreadsheet (XLS/XLSX)" };
  }
  if (name.endsWith(".pdf") || type.includes("pdf")) return { key: "pdf", label: "PDF" };
  if (type.startsWith("image/") || /\.(jpg|jpeg|png|gif|webp|bmp|heic)$/i.test(name)) return { key: "image", label: "Image" };
  if (name.endsWith(".doc") || name.endsWith(".docx") || type.includes("word")) return { key: "word", label: "Word document" };
  return { key: "other", label: "Unknown format" };
}

async function ensureXlsxLibrary() {
  if (window.XLSX) return window.XLSX;
  if (xlsxLoaderPromise) return xlsxLoaderPromise;

  xlsxLoaderPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    script.onload = () => {
      if (window.XLSX) resolve(window.XLSX);
      else reject(new Error("XLSX parser loaded but unavailable."));
    };
    script.onerror = () => reject(new Error("Could not load XLSX parser library."));
    document.head.appendChild(script);
  });

  return xlsxLoaderPromise;
}

async function ensureTesseractLibrary() {
  if (window.Tesseract?.recognize) return window.Tesseract;
  if (tesseractLoaderPromise) return tesseractLoaderPromise;

  tesseractLoaderPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
    script.onload = () => {
      if (window.Tesseract?.recognize) resolve(window.Tesseract);
      else reject(new Error("OCR library loaded but unavailable."));
    };
    script.onerror = () => reject(new Error("Could not load OCR library."));
    document.head.appendChild(script);
  });

  return tesseractLoaderPromise;
}

async function ensurePdfJsLibrary() {
  if (window.pdfjsLib?.getDocument) return window.pdfjsLib;
  if (pdfjsLoaderPromise) return pdfjsLoaderPromise;

  pdfjsLoaderPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
    script.onload = () => {
      if (!window.pdfjsLib?.getDocument) {
        reject(new Error("PDF parser loaded but unavailable."));
        return;
      }
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
      resolve(window.pdfjsLib);
    };
    script.onerror = () => reject(new Error("Could not load PDF parser library."));
    document.head.appendChild(script);
  });

  return pdfjsLoaderPromise;
}

async function ensureMammothLibrary() {
  if (window.mammoth?.extractRawText) return window.mammoth;
  if (mammothLoaderPromise) return mammothLoaderPromise;

  mammothLoaderPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js";
    script.onload = () => {
      if (window.mammoth?.extractRawText) resolve(window.mammoth);
      else reject(new Error("DOCX parser loaded but unavailable."));
    };
    script.onerror = () => reject(new Error("Could not load DOCX parser library."));
    document.head.appendChild(script);
  });

  return mammothLoaderPromise;
}

async function extractTextFromImageFile(file, progressCb = null) {
  const Tesseract = await ensureTesseractLibrary();
  const result = await Tesseract.recognize(file, "eng", {
    logger: (entry) => {
      if (!progressCb) return;
      if (entry?.status === "recognizing text" && Number.isFinite(entry.progress)) {
        progressCb(`OCR image in progress: ${Math.round(entry.progress * 100)}%`);
      }
    },
  });
  return String(result?.data?.text || "");
}

async function extractTextFromPdfFile(file, progressCb = null) {
  const pdfjsLib = await ensurePdfJsLibrary();
  const data = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;
  const pageLimit = Math.min(pdf.numPages, 12);
  const chunks = [];

  for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
    if (progressCb) progressCb(`Reading PDF page ${pageNumber}/${pageLimit}...`);
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const line = textContent.items.map((item) => item.str).join(" ");
    chunks.push(line);
  }

  return chunks.join("\n");
}

async function extractTextFromDocxFile(file, progressCb = null) {
  if (progressCb) progressCb("Reading DOCX content...");
  const mammoth = await ensureMammothLibrary();
  const result = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
  return String(result?.value || "");
}

function parseImportRowsFromPlainText(text) {
  const rows = [];
  const lines = String(text || "")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((entry) => entry.trim())
    .filter(Boolean);

  const pushParsedRow = (skuRaw, descriptionRaw, finishRaw, qtyRaw) => {
    const sku = normalizeDeliverySkuValue(skuRaw);
    if (!sku) return;
    rows.push({
      sku,
      description: String(descriptionRaw || sku).trim() || sku,
      finish: normalizeFinishValue(finishRaw || "UNSPECIFIED") || "UNSPECIFIED",
      qty: parseImportQty(qtyRaw, 1),
      category: "other",
      unit: "pcs",
      kitchenType: "",
      clientName: "",
      projectName: "",
      unitCode: "",
      source: "OCR text",
    });
  };

  for (const line of lines) {
    if (/^(sku|item|code|description|qty|quantity|finish|color)\b/i.test(line)) continue;

    let columns = [];
    if (line.includes("\t")) columns = line.split("\t");
    else if (line.includes(";")) columns = line.split(";");
    else if (line.includes("|")) columns = line.split("|");
    else if ((line.match(/,/g) || []).length >= 2) columns = parseDelimitedTextRows(line, ",")[0] || [];
    columns = columns.map((entry) => String(entry || "").trim()).filter(Boolean);

    if (columns.length >= 4) {
      pushParsedRow(columns[0], columns[1], columns[2], columns[3]);
      continue;
    }

    if (columns.length === 3) {
      pushParsedRow(columns[0], columns[1], "UNSPECIFIED", columns[2]);
      continue;
    }

    const basic = line.match(/^([A-Z0-9][A-Z0-9._\-\/]{1,})\s+(.+?)\s+(\d{1,6})$/i);
    if (basic) {
      pushParsedRow(basic[1], basic[2], "UNSPECIFIED", basic[3]);
      continue;
    }

    const extended = line.match(/^([A-Z0-9][A-Z0-9._\-\/]{1,})\s+(.+?)\s+(M\d[-\w]+|[A-Z][A-Z0-9_-]{1,})\s+(\d{1,6})$/i);
    if (extended) {
      pushParsedRow(extended[1], extended[2], extended[3], extended[4]);
    }
  }

  return rows;
}

function detectDelimiterInText(text) {
  const sampleLine = String(text || "")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .find((line) => line.trim());
  if (!sampleLine) return ",";
  const candidates = [",", ";", "\t", "|"];
  let best = ",";
  let bestCount = -1;
  candidates.forEach((delimiter) => {
    const count = sampleLine.split(delimiter).length - 1;
    if (count > bestCount) {
      bestCount = count;
      best = delimiter;
    }
  });
  return best;
}

function parseDelimitedTextRows(text, delimiter) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const ch = text[index];

    if (ch === '"') {
      if (inQuotes && text[index + 1] === '"') {
        cell += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && ch === delimiter) {
      row.push(cell);
      cell = "";
      continue;
    }

    if (!inQuotes && (ch === "\n" || ch === "\r")) {
      if (ch === "\r" && text[index + 1] === "\n") index += 1;
      row.push(cell);
      if (row.some((entry) => String(entry || "").trim())) rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += ch;
  }

  row.push(cell);
  if (row.some((entry) => String(entry || "").trim())) rows.push(row);
  return rows;
}

function importObjectsFromMatrix(matrix) {
  const rows = ensureArray(matrix).map((line) => ensureArray(line).map((value) => String(value ?? "").trim()));
  if (!rows.length) return [];

  const first = rows[0];
  const knownHeaders = first.filter((header) => {
    const normalized = normalizeImportKey(header);
    return Object.values(DELIVERY_IMPORT_NORMALIZED_ALIASES).some((aliases) => aliases.includes(normalized));
  }).length;
  const hasHeader = knownHeaders > 0;
  const headers = hasHeader ? first : first.map((_, index) => `col${index + 1}`);
  const startIndex = hasHeader ? 1 : 0;
  const objects = [];

  for (let rowIndex = startIndex; rowIndex < rows.length; rowIndex += 1) {
    const current = rows[rowIndex];
    if (!current.some((entry) => String(entry || "").trim())) continue;
    const obj = {};
    const size = Math.max(headers.length, current.length);
    for (let col = 0; col < size; col += 1) {
      const key = headers[col] || `col${col + 1}`;
      obj[key] = current[col] ?? "";
    }
    objects.push(obj);
  }

  return objects;
}

function importObjectsFromJsonPayload(payload) {
  if (Array.isArray(payload)) return payload;
  if (!payload || typeof payload !== "object") return [];
  if (Array.isArray(payload.rows)) return payload.rows;
  if (Array.isArray(payload.items)) return payload.items;
  if (Array.isArray(payload.data)) return payload.data;
  return [payload];
}

function pickImportValue(map, field) {
  const aliases = DELIVERY_IMPORT_NORMALIZED_ALIASES[field] || [];
  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(map, alias)) return map[alias];
  }
  return "";
}

function parseImportRowObject(rawObject) {
  if (!rawObject || typeof rawObject !== "object") return null;
  const normalizedMap = {};
  Object.entries(rawObject).forEach(([key, value]) => {
    const normalizedKey = normalizeImportKey(key);
    if (!normalizedKey || Object.prototype.hasOwnProperty.call(normalizedMap, normalizedKey)) return;
    normalizedMap[normalizedKey] = value;
  });

  const descriptionRaw = String(pickImportValue(normalizedMap, "description") || "").trim();
  let sku = normalizeDeliverySkuValue(pickImportValue(normalizedMap, "sku"));
  if (!sku && descriptionRaw) sku = normalizeDeliverySkuValue(descriptionRaw);
  if (!sku) return null;

  const finish = normalizeFinishValue(pickImportValue(normalizedMap, "finish") || "UNSPECIFIED") || "UNSPECIFIED";
  const qty = parseImportQty(pickImportValue(normalizedMap, "qty"), 1);
  const category = String(pickImportValue(normalizedMap, "category") || "other").trim().toLowerCase() || "other";
  const unit = String(pickImportValue(normalizedMap, "unit") || "pcs").trim() || "pcs";

  return {
    sku,
    description: descriptionRaw || sku,
    finish,
    qty,
    category,
    unit,
    kitchenType: String(pickImportValue(normalizedMap, "kitchenType") || "").trim(),
    clientName: String(pickImportValue(normalizedMap, "clientName") || "").trim(),
    projectName: String(pickImportValue(normalizedMap, "projectName") || "").trim(),
    unitCode: String(pickImportValue(normalizedMap, "unitCode") || "").trim(),
    source: String(pickImportValue(normalizedMap, "source") || "").trim(),
  };
}

function parseImportRowsFromObjects(objects) {
  const rows = [];
  const matrixRows = [];

  ensureArray(objects).forEach((obj) => {
    if (Array.isArray(obj)) {
      matrixRows.push(obj);
      return;
    }
    if (!obj || typeof obj !== "object") return;
    const parsed = parseImportRowObject(obj);
    if (parsed) rows.push(parsed);
  });

  if (matrixRows.length) {
    importObjectsFromMatrix(matrixRows).forEach((rowObj) => {
      const parsed = parseImportRowObject(rowObj);
      if (parsed) rows.push(parsed);
    });
  }

  return rows;
}

async function parseRowsFromImportFile(file, { progressCb = null } = {}) {
  const info = detectImportFileKind(file);
  const base = {
    fileName: file.name || "file",
    fileTypeLabel: info.label,
    fileSizeLabel: formatBytes(file.size),
    supported: true,
    warnings: [],
    rows: [],
    rawRowsCount: 0,
    extractedText: "",
  };

  if (info.key === "other") {
    return {
      ...base,
      supported: false,
      warnings: [`Detected: ${info.label}. ${DELIVERY_IMPORT_SUPPORTED_HINT}`],
    };
  }

  if (info.key === "word" && !String(file.name || "").toLowerCase().endsWith(".docx")) {
    return {
      ...base,
      supported: false,
      warnings: ["Legacy .DOC is not parseable in-browser. Convert the file to DOCX and retry."],
    };
  }

  if (["pdf", "image", "word"].includes(info.key)) {
    if (progressCb) progressCb(`Starting ${info.label} extraction...`);
    let extractedText = "";

    if (info.key === "image") extractedText = await extractTextFromImageFile(file, progressCb);
    if (info.key === "pdf") extractedText = await extractTextFromPdfFile(file, progressCb);
    if (info.key === "word") extractedText = await extractTextFromDocxFile(file, progressCb);

    const rows = parseImportRowsFromPlainText(extractedText);
    const warnings = [];
    if (!extractedText.trim()) warnings.push(`No readable text found in ${info.label}.`);
    if (!rows.length) warnings.push("Text was extracted but no SKU/quantity rows were recognized.");

    return {
      ...base,
      rows,
      extractedText,
      warnings,
      rawRowsCount: rows.length,
    };
  }

  let rowObjects = [];
  if (info.key === "csv" || info.key === "tsv" || info.key === "text") {
    const text = await file.text();
    const delimiter = info.key === "tsv" ? "\t" : detectDelimiterInText(text);
    rowObjects = importObjectsFromMatrix(parseDelimitedTextRows(text, delimiter));
    if (!rowObjects.length && text.trim()) {
      const fallbackRows = parseImportRowsFromPlainText(text);
      return {
        ...base,
        rows: fallbackRows,
        extractedText: text,
        rawRowsCount: fallbackRows.length,
        warnings: fallbackRows.length ? ["Used text-pattern parser fallback."] : ["No recognized rows in text file."],
      };
    }
  } else if (info.key === "json") {
    const payload = JSON.parse(await file.text());
    const rows = importObjectsFromJsonPayload(payload);
    rowObjects = Array.isArray(rows[0]) ? importObjectsFromMatrix(rows) : rows;
  } else if (info.key === "spreadsheet") {
    const XLSX = await ensureXlsxLibrary();
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
      return {
        ...base,
        supported: false,
        warnings: ["Spreadsheet has no readable sheets."],
      };
    }
    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
    rowObjects = importObjectsFromMatrix(matrix);
  }

  const parsedRows = parseImportRowsFromObjects(rowObjects);
  return {
    ...base,
    rows: parsedRows,
    rawRowsCount: rowObjects.length,
  };
}

function renderDeliveryImportPreview() {
  if (!deliveryImportPreview) return;
  if (!deliveryImportDraft) {
    deliveryImportPreview.innerHTML = "";
    return;
  }

  if (!deliveryImportDraft.supported) {
    deliveryImportPreview.innerHTML = "";
    return;
  }

  const previewRows = ensureArray(deliveryImportDraft.rows).slice(0, 12);
  if (!previewRows.length) {
    deliveryImportPreview.innerHTML = `<table class="data-table"><tbody><tr><td>No recognized material rows found in this file.</td></tr></tbody></table>`;
    return;
  }

  const htmlRows = previewRows
    .map(
      (row) => `<tr>
      <td>${escapeHtml(row.sku)}</td>
      <td>${escapeHtml(row.description || "-")}</td>
      <td>${escapeHtml(row.finish || "-")}</td>
      <td>${row.qty}</td>
      <td>${escapeHtml(row.clientName || "-")}</td>
      <td>${escapeHtml(row.projectName || "-")}</td>
      <td>${escapeHtml(row.unitCode || "-")}</td>
    </tr>`
    )
    .join("");

  const totalRows = ensureArray(deliveryImportDraft.rows).length;
  const hasMore = totalRows > previewRows.length;
  deliveryImportPreview.innerHTML = `<table class="data-table"><thead><tr><th>SKU</th><th>Description</th><th>Finish</th><th>Qty</th><th>Client</th><th>Project</th><th>Unit</th></tr></thead><tbody>${htmlRows}</tbody></table>${
    hasMore ? `<p class="hint">Showing ${previewRows.length} of ${totalRows} rows.</p>` : ""
  }`;
}

function findImportMatchByName(list, nameGetter, lookup) {
  const normalizedLookup = normalizeLookupText(lookup);
  if (!normalizedLookup) return null;
  const exact = list.find((entry) => normalizeLookupText(nameGetter(entry)) === normalizedLookup);
  if (exact) return exact;
  return (
    list.find((entry) => {
      const normalizedName = normalizeLookupText(nameGetter(entry));
      return normalizedName && (normalizedName.includes(normalizedLookup) || normalizedLookup.includes(normalizedName));
    }) || null
  );
}

function resolveImportDestinationIds(row) {
  const client = findImportMatchByName(clients, (entry) => entry.name, row.clientName);
  const clientId = client?.id || "";

  let projectPool = projects;
  if (clientId) projectPool = projects.filter((entry) => entry.clientId === clientId);
  const project = findImportMatchByName(projectPool, (entry) => entry.name, row.projectName);
  const projectId = project?.id || "";

  let unitPool = units;
  if (projectId) unitPool = units.filter((entry) => entry.projectId === projectId);
  const unit = findImportMatchByName(unitPool, (entry) => entry.unitCode, row.unitCode);
  const unitId = unit?.id || "";

  return { clientId, projectId, unitId };
}

function mergeSourceRefs(sourceRefs, source, qty) {
  const map = new Map(
    ensureArray(sourceRefs).map((entry) => [
      String(entry?.source || "").trim() || "Unknown source",
      Math.max(0, Math.trunc(toNumber(entry?.qty || 0))),
    ])
  );
  const normalizedSource = String(source || "").trim() || "File import";
  map.set(normalizedSource, (map.get(normalizedSource) || 0) + Math.max(0, Math.trunc(toNumber(qty || 0))));
  return Array.from(map.entries()).map(([refSource, refQty]) => ({ source: refSource, qty: refQty }));
}

function aggregateImportRows(rows) {
  const map = new Map();
  ensureArray(rows).forEach((row) => {
    const key = deliverySkuKey(row.sku, row.finish);
    if (!map.has(key)) {
      map.set(key, {
        ...row,
        qty: 0,
      });
    }
    const target = map.get(key);
    target.qty += Math.max(1, Math.trunc(toNumber(row.qty || 1)));
    if (!target.description && row.description) target.description = row.description;
    if (!target.category && row.category) target.category = row.category;
    if (!target.unit && row.unit) target.unit = row.unit;
    if (!target.kitchenType && row.kitchenType) target.kitchenType = row.kitchenType;
    if (!target.clientName && row.clientName) target.clientName = row.clientName;
    if (!target.projectName && row.projectName) target.projectName = row.projectName;
    if (!target.unitCode && row.unitCode) target.unitCode = row.unitCode;
    if (!target.source && row.source) target.source = row.source;
  });
  return Array.from(map.values());
}

async function importRowsToDeliveryInventory(rows, fileName) {
  const aggregated = aggregateImportRows(rows);
  if (!aggregated.length) return { created: 0, updated: 0, importedRows: 0 };

  const output = deliverySkuItems.map((entry) => normalizeDeliverySkuItem(entry));
  const byKey = new Map(output.map((entry) => [deliverySkuKey(entry.sku, entry.finish), entry]));
  const sourceFallback = `File import: ${fileName}`;
  let created = 0;
  let updated = 0;
  const now = new Date().toISOString();

  aggregated.forEach((row) => {
    const key = deliverySkuKey(row.sku, row.finish);
    const destinationIds = resolveImportDestinationIds(row);
    const sourceLabel = row.source || sourceFallback;
    const existing = byKey.get(key);
    if (existing) {
      existing.totalQty += row.qty;
      if (!existing.description && row.description) existing.description = row.description;
      if (!existing.clientId && destinationIds.clientId) existing.clientId = destinationIds.clientId;
      if (!existing.projectId && destinationIds.projectId) existing.projectId = destinationIds.projectId;
      if (!existing.unitId && destinationIds.unitId) existing.unitId = destinationIds.unitId;
      existing.sourceRefs = mergeSourceRefs(existing.sourceRefs, sourceLabel, row.qty);
      existing.updatedAt = now;
      updated += 1;
      return;
    }

    const createdItem = normalizeDeliverySkuItem({
      id: uid(),
      sku: row.sku,
      description: row.description || row.sku,
      finish: row.finish || "UNSPECIFIED",
      totalQty: row.qty,
      scannedQty: 0,
      clientId: destinationIds.clientId,
      projectId: destinationIds.projectId,
      unitId: destinationIds.unitId,
      sourceRefs: [{ source: sourceLabel, qty: row.qty }],
      scanLog: [],
      createdAt: now,
      updatedAt: now,
    });
    output.push(createdItem);
    byKey.set(key, createdItem);
    created += 1;
  });

  await writeDeliverySkuItems(output, { replace: true });
  pushEntityAudit(
    "Delivery Inventory",
    "imported",
    `Imported ${aggregated.length} row(s) from ${fileName}. Created: ${created}. Updated: ${updated}.`,
    "delivery"
  );
  return { created, updated, importedRows: aggregated.length };
}

async function importRowsToMaterialCatalog(rows, fileName) {
  const aggregated = aggregateImportRows(rows);
  if (!aggregated.length) return { created: 0, updated: 0, importedRows: 0 };

  const existingBySku = new Map(materials.map((entry) => [normalizeDeliverySkuValue(entry.sku), entry]));
  const now = new Date().toISOString();
  let created = 0;
  let updated = 0;
  const transaction = db.transaction([MATERIAL_STORE], "readwrite");
  const store = transaction.objectStore(MATERIAL_STORE);

  aggregated.forEach((row) => {
    const lookupKey = normalizeDeliverySkuValue(row.sku);
    const existing = existingBySku.get(lookupKey);
    if (existing) {
      store.put(
        normalizeMaterial({
          ...existing,
          description: row.description || existing.description,
          category: row.category || existing.category || "other",
          unit: row.unit || existing.unit || "pcs",
          kitchenType: row.kitchenType || existing.kitchenType || "",
          updatedAt: now,
        })
      );
      updated += 1;
      return;
    }

    store.put(
      normalizeMaterial({
        id: uid(),
        sku: row.sku,
        description: row.description || row.sku,
        category: row.category || "other",
        unit: row.unit || "pcs",
        kitchenType: row.kitchenType || "",
        createdAt: now,
        updatedAt: now,
      })
    );
    created += 1;
  });

  await transactionDonePromise(transaction);
  await loadAll();
  render();
  queueAutoSync();
  pushEntityAudit(
    "Materials",
    "imported",
    `Imported ${aggregated.length} row(s) from ${fileName}. Created: ${created}. Updated: ${updated}.`,
    "materials"
  );
  return { created, updated, importedRows: aggregated.length };
}

async function importRowsToContainerManifest(rows, fileName, containerId) {
  const container = containers.find((entry) => entry.id === containerId);
  if (!container) throw new Error("Select a valid container before importing.");

  const aggregated = aggregateImportRows(rows);
  if (!aggregated.length) return { created: 0, qrCreated: 0, importedRows: 0, containerCode: container.containerCode };

  const now = new Date().toISOString();
  let created = 0;
  let qrCreated = 0;

  aggregated.forEach((row) => {
    const matchedMaterial =
      materials.find((entry) => normalizeDeliverySkuValue(entry.sku) === normalizeDeliverySkuValue(row.sku)) || null;
    const materialId = matchedMaterial?.id || "";
    const category = row.category || matchedMaterial?.category || "other";
    const unit = row.unit || matchedMaterial?.unit || "pcs";
    const description = row.description || matchedMaterial?.description || row.sku;
    const code = row.sku || `ITEM-${String(container.materialItems.length + 1).padStart(3, "0")}`;
    const qty = Math.max(1, Math.trunc(toNumber(row.qty || 1)));

    const manifestItemId = uid();
    container.materialItems.push({
      id: manifestItemId,
      materialId,
      code,
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
    created += 1;

    for (let index = 0; index < qty; index += 1) {
      container.qrItems.push({
        id: uid(),
        manifestItemId,
        qrCode: makeQrCode(container, code || "ITEM"),
        materialId,
        code,
        description,
        category,
        unit,
        status: "in-warehouse",
        clientId: "",
        projectId: "",
        unitId: "",
        unitLabel: "",
        kitchenType: row.kitchenType || matchedMaterial?.kitchenType || "",
        assignedAt: "",
        assignedBy: "",
        deliveredAt: "",
        issueType: "",
        issueNote: "",
        flowStage: "warehouse",
        flowStatus: "generated",
        flowUpdatedAt: now,
        flowEvents: [
          {
            id: uid(),
            at: now,
            byUserId: currentUser?.id || "",
            byName: currentUser?.name || "System",
            stage: "warehouse",
            status: "generated",
            issueType: "",
            note: `QR generated from imported file: ${fileName}`,
            unitId: "",
            unitLabel: "",
            source: "manifest",
          },
        ],
        createdAt: now,
        updatedAt: now,
      });
      qrCreated += 1;
    }
  });

  pushContainerAudit(container, `Imported ${aggregated.length} row(s) from file: ${fileName}`);
  pushContainerAudit(container, `Generated ${qrCreated} QR item(s) from imported rows.`);
  await saveContainer(container);
  pushEntityAudit(
    "Containers",
    "manifest-imported",
    `${container.containerCode}: imported ${aggregated.length} row(s), created ${qrCreated} QR item(s) from ${fileName}.`,
    "containers"
  );
  return { created, qrCreated, importedRows: aggregated.length, containerCode: container.containerCode };
}

function canImportRowsByTarget(target) {
  const canDeliveryImport = can("manageCatalog") || can("manageContainerRelease") || can("siteReceive") || isAdmin() || isDeveloper();
  const canMaterialImport = can("manageCatalog") || isAdmin() || isDeveloper();
  const canManifestImport = can("manageContainerManifest") || isAdmin() || isDeveloper();

  if (target === "deliverySku") return canDeliveryImport;
  if (target === "materialCatalog") return canMaterialImport;
  if (target === "containerManifest") return canManifestImport;
  return false;
}

async function applyImportRowsByTarget({ rows, target, fileName, containerId }) {
  if (!ensureArray(rows).length) throw new Error("No recognized rows to import.");
  if (!canImportRowsByTarget(target)) throw new Error("Your access profile cannot import to this target.");

  if (target === "deliverySku") {
    const result = await importRowsToDeliveryInventory(rows, fileName);
    return `Delivery import complete. Created: ${result.created}. Updated: ${result.updated}. Imported rows: ${result.importedRows}.`;
  }

  if (target === "materialCatalog") {
    const result = await importRowsToMaterialCatalog(rows, fileName);
    return `Material catalog import complete. Created: ${result.created}. Updated: ${result.updated}. Imported rows: ${result.importedRows}.`;
  }

  if (!containerId) throw new Error("Select a container before importing to manifest.");
  const result = await importRowsToContainerManifest(rows, fileName, containerId);
  return `Container import complete (${result.containerCode}). Manifest rows: ${result.created}. QR generated: ${result.qrCreated}.`;
}

function setOcrImportStatus(message = "", isError = false) {
  if (!ocrImportStatus) return;
  ocrImportStatus.textContent = message;
  ocrImportStatus.style.color = isError ? "#b83232" : "var(--ink-soft)";
}

function syncOcrImportContainerSelect() {
  if (!ocrImportContainerSelect) return;
  const current = ocrImportContainerSelect.value;
  const rows = containers.map((container) => {
    const clientName = clientById(container.clientId)?.name || "No client";
    const projectName = projectById(container.projectId)?.name || "No project";
    return `<option value="${escapeHtml(container.id)}">${escapeHtml(container.containerCode)} | ${escapeHtml(clientName)} | ${escapeHtml(projectName)}</option>`;
  });
  ocrImportContainerSelect.innerHTML = `<option value="">Select a container</option>${rows.join("")}`;
  if (containers.some((entry) => entry.id === current)) ocrImportContainerSelect.value = current;
}

function refreshOcrImportTargetUi() {
  const target = ocrImportTargetSelect?.value || "deliverySku";
  const needsContainer = target === "containerManifest";
  ocrImportContainerWrap?.classList.toggle("hidden", !needsContainer);
  if (ocrImportApplyBtn) {
    if (target === "deliverySku") ocrImportApplyBtn.textContent = "Import to Delivery inventory";
    else if (target === "containerManifest") ocrImportApplyBtn.textContent = "Import to Container manifest";
    else ocrImportApplyBtn.textContent = "Import to Material catalog";
  }
}

function syncOcrImportApplyState() {
  if (!ocrImportApplyBtn) return;
  const target = ocrImportTargetSelect?.value || "deliverySku";
  const hasRows = Boolean(ocrImportDraft?.supported && ensureArray(ocrImportDraft?.rows).length > 0);
  ocrImportApplyBtn.disabled = !hasRows || !canImportRowsByTarget(target);
}

function clearOcrImportDraft({ clearFile = true, clearStatus = false } = {}) {
  ocrImportDraft = null;
  if (clearFile && ocrImportFileInput) ocrImportFileInput.value = "";
  if (ocrImportTextPreview) ocrImportTextPreview.value = "";
  if (ocrImportRowsPreview) ocrImportRowsPreview.innerHTML = "";
  if (clearStatus) setOcrImportStatus("");
  syncOcrImportApplyState();
}

function renderOcrImportRowsPreview() {
  if (!ocrImportRowsPreview) return;
  if (!ocrImportDraft?.supported) {
    ocrImportRowsPreview.innerHTML = "";
    return;
  }
  const previewRows = ensureArray(ocrImportDraft.rows).slice(0, 12);
  if (!previewRows.length) {
    ocrImportRowsPreview.innerHTML = `<table class="data-table"><tbody><tr><td>No recognized material rows found in extracted text.</td></tr></tbody></table>`;
    return;
  }

  const body = previewRows
    .map(
      (row) => `<tr>
      <td>${escapeHtml(row.sku)}</td>
      <td>${escapeHtml(row.description || "-")}</td>
      <td>${escapeHtml(row.finish || "-")}</td>
      <td>${row.qty}</td>
    </tr>`
    )
    .join("");

  const totalRows = ensureArray(ocrImportDraft.rows).length;
  const moreText = totalRows > previewRows.length ? `<p class="hint">Showing ${previewRows.length} of ${totalRows} rows.</p>` : "";
  ocrImportRowsPreview.innerHTML = `<table class="data-table"><thead><tr><th>SKU</th><th>Description</th><th>Finish</th><th>Qty</th></tr></thead><tbody>${body}</tbody></table>${moreText}`;
}

async function analyzeOcrImportFile() {
  const file = ocrImportFileInput?.files?.[0];
  if (!file) {
    setOcrImportStatus("Choose a file first.", true);
    clearOcrImportDraft({ clearFile: false });
    return;
  }

  setOcrImportStatus(`Analyzing ${file.name}...`);
  try {
    const analysis = await parseRowsFromImportFile(file, {
      progressCb: (message) => setOcrImportStatus(message),
    });
    ocrImportDraft = analysis;
    if (ocrImportTextPreview) ocrImportTextPreview.value = analysis.extractedText || "";
    renderOcrImportRowsPreview();
    syncOcrImportApplyState();

    if (!analysis.supported) {
      setOcrImportStatus(`${analysis.warnings.join(" ")}`, true);
      return;
    }

    const warningText = analysis.warnings.length ? ` ${analysis.warnings.join(" ")}` : "";
    setOcrImportStatus(
      `Detected ${analysis.fileTypeLabel} (${analysis.fileSizeLabel}). Parsed rows: ${analysis.rawRowsCount}. Recognized rows: ${analysis.rows.length}.${warningText}`
    );
  } catch (error) {
    clearOcrImportDraft({ clearFile: false });
    setOcrImportStatus(`OCR analysis failed: ${error?.message || "unknown error"}`, true);
  }
}

async function applyOcrImportDraft() {
  if (!ocrImportDraft?.supported || !ensureArray(ocrImportDraft.rows).length) {
    setOcrImportStatus("Run OCR/Analyze first and confirm recognized rows.", true);
    return;
  }
  const target = ocrImportTargetSelect?.value || "deliverySku";
  const fileName = ocrImportDraft.fileName || "ocr-file";
  const containerId = ocrImportContainerSelect?.value || "";
  try {
    const message = await applyImportRowsByTarget({
      rows: ocrImportDraft.rows,
      target,
      fileName,
      containerId,
    });
    setOcrImportStatus(message);
  } catch (error) {
    setOcrImportStatus(`Import failed: ${error?.message || "unknown error"}`, true);
  }
}

async function analyzeDeliveryImportFile() {
  const file = deliveryImportFileInput?.files?.[0];
  if (!file) {
    setDeliveryImportStatus("Choose a file first.", true);
    clearDeliveryImportDraft({ clearFile: false });
    return;
  }

  setDeliveryImportStatus(`Analyzing ${file.name}...`);
  try {
    const analysis = await parseRowsFromImportFile(file, {
      progressCb: (message) => setDeliveryImportStatus(message),
    });
    deliveryImportDraft = analysis;
    renderDeliveryImportPreview();
    syncDeliveryImportApplyState();

    if (!analysis.supported) {
      setDeliveryImportStatus(`${analysis.warnings.join(" ")}`, true);
      return;
    }

    const warningText = analysis.warnings.length ? ` ${analysis.warnings.join(" ")}` : "";
    setDeliveryImportStatus(
      `Detected ${analysis.fileTypeLabel} (${analysis.fileSizeLabel}). Parsed rows: ${analysis.rawRowsCount}. Recognized material rows: ${analysis.rows.length}.${warningText}`
    );
  } catch (error) {
    clearDeliveryImportDraft({ clearFile: false });
    setDeliveryImportStatus(`Could not analyze file: ${error?.message || "unknown error"}`, true);
  }
}

async function applyDeliveryImportDraft() {
  if (!deliveryImportDraft?.supported || !ensureArray(deliveryImportDraft.rows).length) {
    setDeliveryImportStatus("Analyze a supported file with recognized rows before importing.", true);
    return;
  }

  const target = deliveryImportTargetSelect?.value || "deliverySku";
  const fileName = deliveryImportDraft.fileName || "import-file";
  const containerId = deliveryImportContainerSelect?.value || "";

  try {
    const message = await applyImportRowsByTarget({
      rows: deliveryImportDraft.rows,
      target,
      fileName,
      containerId,
    });
    setDeliveryImportStatus(message);
  } catch (error) {
    setDeliveryImportStatus(`Import failed: ${error?.message || "unknown error"}`, true);
  }
}

function syncDeliveryAssignProjectSelect() {
  if (!deliveryAssignProjectSelect) return;
  const current = deliveryAssignProjectSelect.value;
  const clientId = deliveryAssignClientSelect?.value || "";
  const list = deliverySelectionProjects(clientId);
  deliveryAssignProjectSelect.innerHTML = `<option value="">Not assigned</option>${list
    .map((project) => `<option value="${escapeHtml(project.id)}">${escapeHtml(project.name)}</option>`)
    .join("")}`;
  if (list.some((entry) => entry.id === current)) deliveryAssignProjectSelect.value = current;
}

function syncDeliveryAssignUnitSelect() {
  if (!deliveryAssignUnitSelect) return;
  const current = deliveryAssignUnitSelect.value;
  const projectId = deliveryAssignProjectSelect?.value || "";
  const list = deliverySelectionUnits(projectId);
  deliveryAssignUnitSelect.innerHTML = `<option value="">Not assigned</option>${list
    .map((unit) => `<option value="${escapeHtml(unit.id)}">${escapeHtml(unit.unitCode || "-")} - ${escapeHtml(unit.projectName || "-")}</option>`)
    .join("")}`;
  if (list.some((entry) => entry.id === current)) deliveryAssignUnitSelect.value = current;
}

function resetDeliveryScanDestination() {
  if (deliveryAssignClientSelect) deliveryAssignClientSelect.value = "";
  syncDeliveryAssignProjectSelect();
  if (deliveryAssignProjectSelect) deliveryAssignProjectSelect.value = "";
  syncDeliveryAssignUnitSelect();
  if (deliveryAssignUnitSelect) deliveryAssignUnitSelect.value = "";
  if (deliveryDestinationNoteInput) deliveryDestinationNoteInput.value = "";
}

function deliveryDestinationLabel(item) {
  const client = clientById(item.clientId);
  const project = projectById(item.projectId);
  const unit = units.find((entry) => entry.id === item.unitId);
  const parts = [client?.name, project?.name, unit?.unitCode].filter(Boolean);
  if (item.destinationNote) parts.push(item.destinationNote);
  return parts.join(" | ") || "-";
}

function renderDeliveryInventoryTable() {
  if (!deliveryInventoryTable) return;
  const rows = deliverySkuItems
    .map((item) => {
      const available = deliverySkuAvailableQty(item);
      const sourceRefs = ensureArray(item.sourceRefs)
        .filter((entry) => entry.qty > 0)
        .map((entry) => `${entry.source}: ${entry.qty}`)
        .join(", ");
      return `<tr>
        <td><span class="qr-inline-code">${escapeHtml(item.qrCode)}</span></td>
        <td>${escapeHtml(item.sku)}</td>
        <td>${escapeHtml(item.finish || "-")}</td>
        <td>${escapeHtml(item.description || "-")}</td>
        <td>${item.totalQty}</td>
        <td>${item.scannedQty}</td>
        <td>${available}</td>
        <td>${escapeHtml(deliveryDestinationLabel(item))}</td>
        <td>${escapeHtml(sourceRefs || "-")}</td>
        <td>${escapeHtml(fmtDate(item.updatedAt))}</td>
        <td>
          <div class="actions-inline">
            <button class="secondary xs-btn" type="button" data-delivery-copy-qr="${item.id}">Copy QR</button>
            <button class="secondary xs-btn" type="button" data-delivery-select="${item.id}">Select</button>
          </div>
        </td>
      </tr>`;
    })
    .join("");

  deliveryInventoryTable.innerHTML = `<table class="data-table"><thead><tr><th>QR</th><th>SKU</th><th>Finish</th><th>Description</th><th>Total</th><th>Scanned</th><th>Available</th><th>Destination</th><th>Source</th><th>Updated</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="11">No delivery inventory items yet.</td></tr>'
  }</tbody></table>`;

  deliveryInventoryTable.querySelectorAll("[data-delivery-copy-qr]").forEach((button) => {
    button.addEventListener("click", async () => {
      const item = deliverySkuById(button.dataset.deliveryCopyQr);
      if (!item) return;
      try {
        await navigator.clipboard.writeText(item.qrCode);
      } catch {
        window.prompt("Copy QR code:", item.qrCode);
      }
    });
  });

  deliveryInventoryTable.querySelectorAll("[data-delivery-select]").forEach((button) => {
    button.addEventListener("click", () => {
      const item = deliverySkuById(button.dataset.deliverySelect);
      if (!item) return;
      if (deliveryScanQrInput) deliveryScanQrInput.value = item.qrCode;
      if (deliveryAssignClientSelect) deliveryAssignClientSelect.value = item.clientId || "";
      syncDeliveryAssignProjectSelect();
      if (deliveryAssignProjectSelect) deliveryAssignProjectSelect.value = item.projectId || "";
      syncDeliveryAssignUnitSelect();
      if (deliveryAssignUnitSelect) deliveryAssignUnitSelect.value = item.unitId || "";
      if (deliveryDestinationNoteInput) deliveryDestinationNoteInput.value = item.destinationNote || "";
      setDeliveryInventoryStatus(`Selected ${item.sku} (${item.finish}).`);
      deliveryScanForm?.scrollIntoView({ behavior: "smooth", block: "center" });
    });
  });
}

function renderDeliveryInventoryPanel() {
  if (!deliveryInventoryPanel) return;
  const visible = currentView === "projects" && currentProjectSector === "delivery";
  deliveryInventoryPanel.classList.toggle("hidden-view", !visible);
  if (!visible) return;

  const canEdit = can("manageCatalog") || can("manageContainerRelease") || can("siteReceive") || isAdmin() || isDeveloper();
  const canImportAny = can("manageCatalog") || can("manageContainerManifest") || can("manageContainerRelease") || can("siteReceive") || isAdmin() || isDeveloper();
  deliveryInventorySeedBtn.disabled = !canEdit;
  deliveryInventoryResetBtn.disabled = !canEdit;
  deliveryAddForm?.querySelectorAll("input,button").forEach((field) => (field.disabled = !canEdit));
  deliveryScanForm?.querySelectorAll("input,select,button").forEach((field) => (field.disabled = !canEdit));
  deliveryImportForm?.querySelectorAll("input,select,button").forEach((field) => {
    if (field === deliveryImportApplyBtn) return;
    field.disabled = !canImportAny;
  });
  if (deliveryImportApplyBtn && !canImportAny) deliveryImportApplyBtn.disabled = true;

  const currentClient = deliveryAssignClientSelect?.value || "";
  deliveryAssignClientSelect.innerHTML = `<option value="">Not assigned</option>${clients
    .map((client) => `<option value="${escapeHtml(client.id)}">${escapeHtml(client.name)}</option>`)
    .join("")}`;
  if (clients.some((client) => client.id === currentClient)) deliveryAssignClientSelect.value = currentClient;
  syncDeliveryAssignProjectSelect();
  syncDeliveryAssignUnitSelect();
  syncDeliveryImportContainerSelect();
  refreshDeliveryImportTargetUi();
  syncDeliveryImportApplyState();
  renderDeliveryImportPreview();
  renderDeliveryInventoryTable();
}

function renderOcrImporterPanel() {
  if (!ocrImporterPanel) return;
  const visible = currentView === "ocrImporter";
  ocrImporterPanel.classList.toggle("hidden-view", !visible);
  if (!visible) return;

  const canUse = can("manageCatalog") || can("manageContainerManifest") || can("manageContainerRelease") || can("siteReceive") || isAdmin() || isDeveloper();
  ocrImportForm?.querySelectorAll("input,select,button").forEach((field) => {
    if (field === ocrImportApplyBtn) return;
    field.disabled = !canUse;
  });

  syncOcrImportContainerSelect();
  refreshOcrImportTargetUi();
  syncOcrImportApplyState();
  renderOcrImportRowsPreview();
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

function qrToken(value, size = 10) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "")
    .slice(0, size);
}

function makeUnitChecklistQrCode(unit, item, index = 1) {
  const projectToken = qrToken(unit?.projectCode || unit?.projectId || unit?.projectName, 10) || "PROJECT";
  const unitToken = qrToken(unit?.unitCode || unit?.id, 10) || "UNIT";
  const itemToken = qrToken(item?.code || item?.description || item?.id, 12) || "ITEM";
  const itemIdToken = qrToken(item?.id, 8) || "00000000";
  const seq = String(index).padStart(3, "0");
  return `UQR|${projectToken}|${unitToken}|${itemToken}|${seq}|${itemIdToken}`;
}

function normalizeChecklistQrItems(item, unit, expectedQty, now) {
  const list = ensureArray(item.qrItems)
    .map((entry, idx) => ({
      id: entry.id || uid(),
      qrCode: entry.qrCode || makeUnitChecklistQrCode(unit, item, idx + 1),
      createdAt: entry.createdAt || now,
      updatedAt: entry.updatedAt || entry.createdAt || now,
    }))
    .filter((entry) => entry.qrCode);

  let index = list.length;
  while (list.length < expectedQty) {
    index += 1;
    list.push({
      id: uid(),
      qrCode: makeUnitChecklistQrCode(unit, item, index),
      createdAt: now,
      updatedAt: now,
    });
  }
  return list;
}

function normalizeCheckItem(item, unit = null, fallbackIndex = 0) {
  const now = new Date().toISOString();
  const expectedQty = Math.max(1, Math.trunc(toNumber(item.expectedQty) || 1));
  const checkedQty = Math.max(0, Math.trunc(toNumber(item.checkedQty)));
  const normalized = {
    id: item.id || uid(),
    code: String(item.code || "").trim(),
    description: String(item.description || "").trim(),
    expectedQty,
    checkedQty,
    status: item.status || "ok",
    note: item.note || "",
    createdAt: item.createdAt || now,
    updatedAt: item.updatedAt || item.createdAt || now,
  };
  if (!normalized.code) normalized.code = `ITEM-${String(fallbackIndex + 1).padStart(3, "0")}`;
  if (!normalized.description) normalized.description = normalized.code;
  normalized.qrItems = normalizeChecklistQrItems(normalized, unit, expectedQty, now);
  return normalized;
}

function createCheckItem(payload, unit = null, fallbackIndex = 0) {
  return normalizeCheckItem(
    {
      id: uid(),
      code: payload.code || "",
      description: payload.description || "",
      expectedQty: payload.expectedQty,
      checkedQty: payload.checkedQty || 0,
      status: payload.status || "ok",
      note: payload.note || "",
      qrItems: [],
      createdAt: payload.createdAt || new Date().toISOString(),
      updatedAt: payload.updatedAt || payload.createdAt || new Date().toISOString(),
    },
    unit,
    fallbackIndex
  );
}

function addDefaultChecklistItemsToUnit(unit) {
  const existing = new Set(ensureArray(unit.checkItems).map((entry) => String(entry.description || "").trim().toLowerCase()).filter(Boolean));
  let added = 0;
  for (const description of UNIT_DEFAULT_QR_MATERIALS) {
    const key = description.trim().toLowerCase();
    if (!key || existing.has(key)) continue;
    unit.checkItems.push(createCheckItem({ code: "", description, expectedQty: 1, checkedQty: 0, status: "ok" }, unit, unit.checkItems.length));
    existing.add(key);
    added += 1;
  }
  return added;
}

function unitChecklistQrCodesText(item) {
  return ensureArray(item?.qrItems).map((entry) => entry.qrCode).filter(Boolean);
}

function openUnitChecklistQrLabels(unit, item, { autoPrint = false } = {}) {
  const labels = unitChecklistQrCodesText(item);
  if (!labels.length) return;
  const win = window.open("", "_blank");
  if (!win) return;
  const html = `<!doctype html>
  <html lang="en">
    <head>
      <meta charset="utf-8" />
      <title>Unit Item QR Labels</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 18px; color: #1f2937; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 10px; }
        .label { border: 1px solid #d1d5db; border-radius: 12px; padding: 10px; }
        h1 { margin: 0 0 10px; font-size: 14px; }
        h2 { margin: 0 0 6px; font-size: 13px; }
        img { width: 170px; height: 170px; border: 1px solid #e5e7eb; display: block; margin: 8px auto; }
        p { margin: 4px 0; font-size: 12px; }
        .qr-code { font-family: "Courier New", monospace; word-break: break-all; font-size: 11px; }
        .actions { margin-bottom: 10px; }
        button { border: 1px solid #cbd5e1; border-radius: 999px; background: #fff; padding: 6px 10px; cursor: pointer; }
        @media print { .actions { display: none; } body { margin: 0; } .label { page-break-inside: avoid; } }
      </style>
    </head>
    <body>
      <div class="actions"><button onclick="window.print()">Print / Save PDF</button></div>
      <div class="grid">
        ${labels
          .map(
            (qrCode, idx) => `<div class="label">
            <h1>${escapeHtml(item.code || "-")} | ${escapeHtml(item.description || "-")}</h1>
            <h2>Piece ${idx + 1} / ${labels.length}</h2>
            <img src="${qrImageUrl(qrCode, 280)}" alt="QR ${escapeHtml(qrCode)}" />
            <p><strong>Unit:</strong> ${escapeHtml(unit?.unitCode || "-")}</p>
            <p><strong>Project:</strong> ${escapeHtml(unit?.projectName || "-")}</p>
            <p><strong>QR:</strong> <span class="qr-code">${escapeHtml(qrCode)}</span></p>
          </div>`
          )
          .join("")}
      </div>
    </body>
  </html>`;
  win.document.open();
  win.document.write(html);
  win.document.close();
  if (autoPrint) setTimeout(() => win.print(), 450);
}

async function downloadUnitChecklistQrPng(item) {
  const firstCode = unitChecklistQrCodesText(item)[0];
  if (!firstCode) return;
  const url = qrImageUrl(firstCode, 512);
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error("QR image download failed");
    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = objectUrl;
    const base = String(firstCode || "unit-item-qr").replace(/[^a-zA-Z0-9_-]+/g, "_");
    a.download = `${base}.png`;
    a.click();
    URL.revokeObjectURL(objectUrl);
  } catch {
    window.open(url, "_blank");
  }
}

function workflowStageLabel(stage) {
  if (!stage) return "-";
  return QR_WORKFLOW_STAGE_LABELS[stage] || stage;
}

function workflowStatusLabel(status) {
  const raw = String(status || "").trim().toLowerCase();
  if (raw === "generated") return "Generated";
  if (raw === "in-progress") return "In progress";
  if (raw === "completed") return "Completed";
  if (raw === "hold") return "On hold";
  if (raw === "issue") return "Issue";
  return raw || "-";
}

function issueTypeLabel(issueType) {
  const raw = String(issueType || "").trim().toLowerCase();
  if (!raw) return "";
  if (raw === "nao-chegou") return "Not received";
  if (raw === "faltando-pecas") return "Missing parts";
  if (raw === "quebrado") return "Damaged";
  if (raw === "medida-diferente") return "Wrong size";
  if (raw === "cor-errada") return "Wrong color";
  return raw;
}

function qrFlowSourceLabel(source) {
  const raw = String(source || "").trim().toLowerCase();
  if (!raw) return "Workflow";
  if (raw === "manifest") return "Manufacturer Manifest";
  if (raw === "dispatch") return "Warehouse Dispatch";
  if (raw === "site-receive") return "Site Receive";
  if (raw === "workflow") return "Workflow Update";
  return raw
    .split("-")
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function normalizeWorkflowStage(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "warehouse") return "warehouse";
  if (raw === "delivery" || raw === "site delivery" || raw === "site-delivery") return "delivery";
  if (raw === "distribution") return "distribution";
  if (raw === "installation") return "installation";
  if (raw === "punch list" || raw === "punchlist") return "punchlist";
  return raw;
}

function normalizeWorkflowStatus(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (QR_WORKFLOW_STATUS_OPTIONS.includes(raw)) return raw;
  if (raw === "received") return "completed";
  if (raw === "dispatched" || raw === "assigned") return "in-progress";
  return "in-progress";
}

function canUpdateQrWorkflow() {
  if (!currentUser) return false;
  if (isDeveloper() || isAdmin() || isProjectManager()) return true;
  return can("manageContainerRelease") || can("siteReceive") || can("checklist") || can("manageContainerIssues");
}

function stageFromProjectSector(sector) {
  if (sector === "warehouse") return "warehouse";
  if (sector === "delivery") return "delivery";
  if (sector === "distribuicao") return "distribution";
  if (sector === "instalacao") return "installation";
  if (sector === "punchlist") return "punchlist";
  return "";
}

function qrLockedStageForCurrentUser() {
  if (!currentUser) return "";
  if (isDeveloper() || isAdmin() || isProjectManager()) return "";
  const accessProfile = userAccessProfile(currentUser);
  if (accessProfile === "warehouse") return "warehouse";
  if (accessProfile === "transport") return "delivery";
  if (accessProfile === "distribution") return "distribution";
  if (accessProfile === "qa") return "punchlist";
  if (accessProfile === "installer" || accessProfile === "foreman") {
    const bySector = stageFromProjectSector(currentProjectSector);
    if (bySector === "distribution" || bySector === "installation") return bySector;
    return accessProfile === "foreman" ? "installation" : "distribution";
  }
  return "";
}

function qrAllowedStagesForCurrentUser() {
  if (!currentUser) return [];
  if (isDeveloper() || isAdmin() || isProjectManager()) return [...QR_WORKFLOW_STAGES, "custom"];
  const locked = qrLockedStageForCurrentUser();
  if (locked) return [locked];
  return [];
}

function qrDefaultStageForCurrentUser() {
  const locked = qrLockedStageForCurrentUser();
  if (locked) return locked;
  const sectorStage = stageFromProjectSector(currentProjectSector);
  if (sectorStage) return sectorStage;
  return "warehouse";
}

function syncQrFlowStageOptions({ preserveSelection = true } = {}) {
  if (!qrFlowStageSelect) return;
  const options = qrAllowedStagesForCurrentUser();
  if (!options.length) return;
  const previous = preserveSelection ? String(qrFlowStageSelect.value || "").trim() : "";
  qrFlowStageSelect.innerHTML = options
    .map((stage) => `<option value="${stage}">${stage === "custom" ? "Custom stage" : workflowStageLabel(stage)}</option>`)
    .join("");
  const preferred = options.includes(previous) ? previous : qrDefaultStageForCurrentUser();
  qrFlowStageSelect.value = options.includes(preferred) ? preferred : options[0];
  qrFlowStageSelect.disabled = !canUpdateQrWorkflow() || Boolean(qrLockedStageForCurrentUser());
  updateQrFlowStageCustomVisibility();
}

function renderQrFlowAccessHint() {
  if (!qrFlowAccessHint) return;
  if (!currentUser) {
    qrFlowAccessHint.textContent = "";
    return;
  }
  if (!canUpdateQrWorkflow()) {
    qrFlowAccessHint.textContent = "Read-only: your access profile cannot update QR workflow stages.";
    return;
  }
  if (isDeveloper() || isAdmin() || isProjectManager()) {
    qrFlowAccessHint.textContent = "Admin/Project Manager/Developer mode: you can route this QR to any workflow stage.";
    return;
  }
  const stage = qrLockedStageForCurrentUser() || qrDefaultStageForCurrentUser();
  if (!stage) {
    qrFlowAccessHint.textContent = "Your access profile has no QR workflow stage assigned.";
    return;
  }
  qrFlowAccessHint.textContent = `Access profile routing active: scans from ${roleLabel(userAccessProfile(currentUser))} go to ${workflowStageLabel(stage)}.`;
}

function resolveQrStageForSubmit() {
  const lockedStage = qrLockedStageForCurrentUser();
  if (lockedStage) return { stage: lockedStage, stagePick: lockedStage, stageCustom: "" };
  const stagePick = qrFlowStageSelect?.value || qrDefaultStageForCurrentUser();
  const stageCustom = qrFlowStageCustomInput?.value?.trim() || "";
  const stage = stagePick === "custom" ? stageCustom : stagePick;
  return { stage, stagePick, stageCustom };
}

function ensureQrItemFlowDefaults(item) {
  if (!item) return;
  item.flowEvents = ensureArray(item.flowEvents);
  if (!item.flowStatus) item.flowStatus = normalizeWorkflowStatus(item.status || "generated");
  if (!item.flowStage) {
    if (item.status === "received" || item.status === "issue") item.flowStage = "delivery";
    else item.flowStage = "warehouse";
  }
  if (!item.flowUpdatedAt) item.flowUpdatedAt = item.updatedAt || item.createdAt || new Date().toISOString();
  if (!item.flowEvents.length && item.createdAt) {
    item.flowEvents.push({
      id: uid(),
      at: item.createdAt,
      byUserId: "",
      byName: "System",
      stage: "warehouse",
      status: "generated",
      issueType: "",
      note: "QR generated from manifest.",
      unitId: item.unitId || "",
      unitLabel: item.unitLabel || "",
      source: "manifest",
    });
  }
}

function appendQrFlowEvent(container, qrItem, { stage, status, issueType = "", note = "", source = "workflow", unit = null, at = "" }) {
  if (!qrItem) return;
  ensureQrItemFlowDefaults(qrItem);
  const now = at || new Date().toISOString();
  const normalizedStage = normalizeWorkflowStage(stage) || "warehouse";
  const normalizedStatus = normalizeWorkflowStatus(status);
  const normalizedIssueType = String(issueType || "").trim();
  const normalizedNote = String(note || "").trim();
  const targetUnit = unit || units.find((entry) => entry.id === qrItem.unitId) || null;

  const event = {
    id: uid(),
    at: now,
    byUserId: currentUser?.id || "",
    byName: currentUser?.name || "System",
    byAccessProfile: userAccessProfile(currentUser),
    bySector: stageFromProjectSector(currentProjectSector) || "",
    stage: normalizedStage,
    status: normalizedStatus,
    issueType: normalizedIssueType,
    note: normalizedNote,
    unitId: targetUnit?.id || qrItem.unitId || "",
    unitLabel: targetUnit?.unitCode || qrItem.unitLabel || "",
    source,
  };

  qrItem.flowEvents.push(event);
  qrItem.flowStage = normalizedStage;
  qrItem.flowStatus = normalizedStatus;
  qrItem.flowUpdatedAt = now;
  qrItem.updatedAt = now;
  if (targetUnit) {
    qrItem.unitId = targetUnit.id;
    qrItem.unitLabel = targetUnit.unitCode;
  }

  if (normalizedStage === "warehouse" && normalizedStatus === "in-progress") qrItem.status = "dispatched";
  if (normalizedStage === "warehouse" && normalizedStatus === "generated") qrItem.status = "in-warehouse";
  if (normalizedStage === "delivery" && normalizedStatus === "completed") qrItem.status = "received";
  if (normalizedStatus === "issue") {
    qrItem.status = "issue";
    qrItem.issueType = normalizedIssueType;
    qrItem.issueNote = normalizedNote;
  }

  const actionText = `${workflowStageLabel(normalizedStage)}: ${normalizedStatus}${normalizedIssueType ? ` (${normalizedIssueType})` : ""}`;
  const detailText = `${qrItem.qrCode}${targetUnit ? ` -> ${targetUnit.unitCode}` : ""}${normalizedNote ? ` | ${normalizedNote}` : ""}`;
  pushContainerAudit(container, `${actionText} | ${detailText}`);
}

function applyStageToUnitFromQr(unit, stage, status, note = "", at = "") {
  if (!unit) return;
  const when = at || new Date().toISOString();
  const normalizedStage = normalizeWorkflowStage(stage);
  const normalizedStatus = normalizeWorkflowStatus(status);
  const done = normalizedStatus === "completed";

  if (normalizedStage === "warehouse") unit.stages.warehouse = { done, at: done ? when : null };
  if (normalizedStage === "delivery") {
    unit.stages.transportation = { done, at: done ? when : null };
    unit.stages.siteDelivery = { done, at: done ? when : null };
    if (normalizedStatus === "issue") unit.deliveryQuality = "rejected";
  }
  if (normalizedStage === "distribution") unit.stages.distribution = { done, at: done ? when : null };
  if (normalizedStage === "installation") {
    unit.stages.installation = { done, at: done ? when : null };
    if (normalizedStatus === "in-progress") unit.installationStatus = "in-progress";
    if (normalizedStatus === "completed") unit.installationStatus = "completed";
    if (normalizedStatus === "hold") unit.installationStatus = "blocked";
  }
  if (normalizedStage === "punchlist") unit.stages.quality = { done, at: done ? when : null };
  if (note) {
    unit.deliveryNotes = unit.deliveryNotes ? `${unit.deliveryNotes}\n${note}` : note;
  }
}

function findQrRecord(qrCode, projectId = "") {
  if (!qrCode) return null;
  const container = findQrItemContainer(qrCode, projectId);
  if (!container) return null;
  const item = ensureArray(container.qrItems).find((entry) => entry.qrCode === qrCode);
  if (!item) return null;
  return { container, item };
}

function populateQrFlowUnitSelect() {
  if (!qrFlowUnitSelect) return;
  const currentProjectUnits =
    currentView === "projects" && selectedProjectId
      ? units.filter((entry) => entry.projectId === selectedProjectId)
      : units;
  const uniqueUnits = currentProjectUnits
    .slice()
    .sort((a, b) => `${a.projectName || ""}-${a.unitCode || ""}`.localeCompare(`${b.projectName || ""}-${b.unitCode || ""}`));
  const previous = qrFlowUnitSelect.value;
  qrFlowUnitSelect.innerHTML = `<option value="">Unit (optional)</option>${uniqueUnits
    .map(
      (entry) =>
        `<option value="${escapeHtml(entry.id)}">${escapeHtml(entry.unitCode || "-")} - ${escapeHtml(entry.projectName || "-")}</option>`
    )
    .join("")}`;
  if (uniqueUnits.some((entry) => entry.id === previous)) qrFlowUnitSelect.value = previous;
}

function updateQrFlowStageCustomVisibility() {
  if (!qrFlowStageCustomWrap || !qrFlowStageCustomInput) return;
  const customMode = qrFlowStageSelect?.value === "custom";
  qrFlowStageCustomWrap.classList.toggle("hidden", !customMode);
  qrFlowStageCustomInput.required = Boolean(customMode);
  if (!customMode) qrFlowStageCustomInput.value = "";
}

function updateQrFlowIssueVisibility() {
  if (!qrFlowIssueTypeSelect) return;
  const issueMode = qrFlowStatusSelect?.value === "issue";
  qrFlowIssueTypeSelect.disabled = !issueMode;
  qrFlowIssueTypeSelect.required = issueMode;
  if (!issueMode) qrFlowIssueTypeSelect.value = "";
}

function setQrFlowResult(message = "", isError = false) {
  if (!qrFlowResult) return;
  qrFlowResult.textContent = message;
  qrFlowResult.style.color = isError ? "#b83232" : "var(--ink-soft)";
}

async function submitQrFlowUpdate({ fromScan = false } = {}) {
  if (!canUpdateQrWorkflow()) return false;

  const qrCode = qrFlowCodeInput?.value?.trim() || "";
  const { stage, stagePick, stageCustom } = resolveQrStageForSubmit();
  const status = normalizeWorkflowStatus(qrFlowStatusSelect?.value || "in-progress");
  const issueType = status === "issue" ? String(qrFlowIssueTypeSelect?.value || "").trim() : "";
  const note = qrFlowNoteInput?.value?.trim() || "";
  const unitId = qrFlowUnitSelect?.value || "";
  const normalizedStage = normalizeWorkflowStage(stage);
  const lockedStage = qrLockedStageForCurrentUser();
  const allowedStages = qrAllowedStagesForCurrentUser()
    .map((entry) => normalizeWorkflowStage(entry))
    .filter(Boolean);

  if (!qrCode) {
    setQrFlowResult("Enter or scan a QR code before updating.", true);
    return false;
  }
  if (!stage) {
    setQrFlowResult("Select a valid stage.", true);
    return false;
  }
  if (status === "issue" && !issueType) {
    setQrFlowResult("Select an issue type when status is Issue.", true);
    return false;
  }
  if (!isDeveloper() && !isAdmin() && !isProjectManager()) {
    if (lockedStage && normalizedStage !== normalizeWorkflowStage(lockedStage)) {
      setQrFlowResult(`Your access profile can only update ${workflowStageLabel(lockedStage)} stage.`, true);
      return false;
    }
    if (allowedStages.length && normalizedStage && !allowedStages.includes(normalizedStage)) {
      setQrFlowResult(`This stage is not available for your access profile (${roleLabel(userAccessProfile(currentUser))}).`, true);
      return false;
    }
    if (stagePick === "custom" || stageCustom) {
      setQrFlowResult("Custom stage is only available to Admin/Developer.", true);
      return false;
    }
  }

  const scopedProjectId = currentView === "projects" ? selectedProjectId : "";
  const snapshot = findQrRecord(qrCode, scopedProjectId);
  if (!snapshot) {
    setQrFlowResult(`QR not found: ${qrCode}`, true);
    return false;
  }

  const { container, item } = snapshot;
  const now = new Date().toISOString();
  const targetUnit = unitId
    ? units.find((entry) => entry.id === unitId) || null
    : units.find((entry) => entry.id === item.unitId) || null;

  if (targetUnit) {
    item.clientId = targetUnit.clientId || container.clientId;
    item.projectId = targetUnit.projectId || container.projectId;
    item.unitId = targetUnit.id;
    item.unitLabel = targetUnit.unitCode;
    if (!item.assignedAt) item.assignedAt = now;
    if (!item.assignedBy) item.assignedBy = currentUser?.name || "";
  }

  if (normalizeWorkflowStage(stage) === "delivery" && status === "completed") item.deliveredAt = now;
  if (status === "issue") {
    item.issueType = issueType;
    item.issueNote = note;
  } else if (status === "completed") {
    item.issueType = "";
    item.issueNote = "";
  }

  appendQrFlowEvent(container, item, {
    stage,
    status,
    issueType,
    note,
    source: "workflow",
    unit: targetUnit,
    at: now,
  });
  if (targetUnit) {
    applyStageToUnitFromQr(targetUnit, stage, status, note, now);
    targetUnit.updatedAt = now;
  }
  if (status === "issue") {
    pushQrIssueQueue(container, item, targetUnit, issueType, note, { source: "workflow", at: now });
  }

  if (targetUnit) {
    await saveUnitAndContainer(targetUnit, container, `QR workflow updated: ${item.qrCode} -> ${workflowStageLabel(stage)} (${workflowStatusLabel(status)})`);
  } else {
    await saveContainer(container);
  }

  lastQrLookupCode = qrCode;
  if (qrLookupInput) qrLookupInput.value = qrCode;
  const actionPrefix = fromScan ? "Scanned and updated" : "Workflow updated";
  setQrFlowResult(`${actionPrefix} for ${qrCode}: ${workflowStageLabel(stage)} / ${workflowStatusLabel(status)}.`);
  renderQrLookupPanel();
  return true;
}

function pushQrIssueQueue(container, qrItem, unit, issueType, note, { source = "workflow", at = "" } = {}) {
  if (!container || !qrItem || !issueType) return;
  const now = at || new Date().toISOString();
  const queueEntry = {
    id: uid(),
    createdAt: now,
    qrCode: qrItem.qrCode,
    code: qrItem.code,
    description: qrItem.description || qrItem.code || qrItem.qrCode,
    unitId: unit?.id || qrItem.unitId || "",
    unitLabel: unit?.unitCode || qrItem.unitLabel || "",
    issueType,
    note: String(note || "").trim(),
    kitchenType: qrItem.kitchenType || "",
    clientId: container.clientId || qrItem.clientId || "",
    projectId: container.projectId || qrItem.projectId || "",
    byUserId: currentUser?.id || "",
    byName: currentUser?.name || "System",
    source,
  };

  if (qrIssueRoutesToFactory(issueType)) {
    container.factoryMissingQueue.push(queueEntry);
  } else {
    container.replacementQueue.push(queueEntry);
  }
}

function resetQrFlowForm() {
  qrFlowForm?.reset();
  syncQrFlowStageOptions({ preserveSelection: false });
  if (qrFlowStatusSelect) qrFlowStatusSelect.value = "in-progress";
  updateQrFlowStageCustomVisibility();
  updateQrFlowIssueVisibility();
  renderQrFlowAccessHint();
  setQrFlowResult("");
}

function stopCameraScanStream() {
  if (cameraScanFrameId) cancelAnimationFrame(cameraScanFrameId);
  cameraScanFrameId = 0;
  if (cameraScanStream) {
    cameraScanStream.getTracks().forEach((track) => track.stop());
    cameraScanStream = null;
  }
  if (cameraScanVideo) {
    cameraScanVideo.pause();
    cameraScanVideo.srcObject = null;
  }
}

function closeCameraScanModal(result = null) {
  stopCameraScanStream();
  if (cameraScanModal) cameraScanModal.classList.add("hidden");
  if (cameraScanResolver) {
    const resolver = cameraScanResolver;
    cameraScanResolver = null;
    resolver(result);
  }
}

function setCameraScanStatus(message) {
  if (cameraScanStatus) cameraScanStatus.textContent = message;
}

function cameraAccessErrorMessage(error) {
  const errorName = String(error?.name || "");
  if (errorName === "NotAllowedError" || errorName === "SecurityError") {
    if (!window.isSecureContext) return "Camera requires HTTPS. Open this app from a secure URL.";
    return "Camera permission denied. Allow camera access in browser settings and retry.";
  }
  if (errorName === "NotFoundError" || errorName === "OverconstrainedError") {
    return "No compatible camera found on this device.";
  }
  if (errorName === "NotReadableError" || errorName === "AbortError") {
    return "Camera is busy. Close other camera apps and retry.";
  }
  return "Unable to access camera. Use photo scan or manual entry.";
}

function loadJsQrLibrary() {
  if (window.jsQR) return Promise.resolve(window.jsQR);
  if (jsQrLoaderPromise) return jsQrLoaderPromise;
  jsQrLoaderPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js";
    script.async = true;
    script.onload = () => {
      if (window.jsQR) resolve(window.jsQR);
      else reject(new Error("jsQR loaded but unavailable"));
    };
    script.onerror = () => reject(new Error("Unable to load jsQR library"));
    document.head.appendChild(script);
  });
  return jsQrLoaderPromise;
}

function scanFeedbackAudioContext() {
  const AudioCtor = window.AudioContext || window.webkitAudioContext;
  if (!AudioCtor) return null;
  if (!scanFeedbackAudioCtx) scanFeedbackAudioCtx = new AudioCtor();
  return scanFeedbackAudioCtx;
}

async function primeScanFeedbackAudio() {
  const context = scanFeedbackAudioContext();
  if (!context) return;
  if (context.state === "suspended") {
    try {
      await context.resume();
    } catch {
      // Ignore; browser may block resume until another user gesture.
    }
  }
}

function playScanFeedbackSound() {
  const context = scanFeedbackAudioContext();
  if (context) {
    const playTone = () => {
      const now = context.currentTime;
      const gain = context.createGain();
      const osc = context.createOscillator();
      osc.type = "square";
      osc.frequency.setValueAtTime(1580, now);
      gain.gain.setValueAtTime(0.0001, now);
      gain.gain.exponentialRampToValueAtTime(0.24, now + 0.006);
      gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.06);
      osc.connect(gain);
      gain.connect(context.destination);
      osc.start(now);
      osc.stop(now + 0.065);
    };
    if (context.state === "suspended") {
      context.resume().then(playTone).catch(() => {});
    } else {
      playTone();
    }
  }
  if (navigator.vibrate) navigator.vibrate(20);
}

function startBarcodeDetectorLoop(DetectorCtor) {
  const detector = new DetectorCtor({ formats: ["qr_code"] });
  const scan = async () => {
    if (!cameraScanVideo?.srcObject || !cameraScanResolver) return;
    try {
      const barcodes = await detector.detect(cameraScanVideo);
      const hit = ensureArray(barcodes).find((entry) => entry.rawValue);
      if (hit?.rawValue) {
        playScanFeedbackSound();
        closeCameraScanModal(String(hit.rawValue).trim());
        return;
      }
      setCameraScanStatus("Point the camera at a QR code.");
    } catch {
      setCameraScanStatus("Scanning...");
    }
    cameraScanFrameId = requestAnimationFrame(scan);
  };
  cameraScanFrameId = requestAnimationFrame(scan);
}

function startJsQrDetectionLoop(jsQR) {
  if (!cameraScanVideo) return;
  if (!cameraScanCanvas) cameraScanCanvas = document.createElement("canvas");
  const context = cameraScanCanvas.getContext("2d", { willReadFrequently: true });
  if (!context) {
    setCameraScanStatus("QR decoder unavailable. Use manual entry.");
    return;
  }

  let lastScanAt = 0;
  const scan = (timestamp = 0) => {
    if (!cameraScanVideo.srcObject || !cameraScanResolver) return;
    const width = cameraScanVideo.videoWidth || 0;
    const height = cameraScanVideo.videoHeight || 0;
    if (!width || !height) {
      cameraScanFrameId = requestAnimationFrame(scan);
      return;
    }
    if (timestamp - lastScanAt < 120) {
      cameraScanFrameId = requestAnimationFrame(scan);
      return;
    }
    lastScanAt = timestamp;
    if (cameraScanCanvas.width !== width || cameraScanCanvas.height !== height) {
      cameraScanCanvas.width = width;
      cameraScanCanvas.height = height;
    }

    try {
      context.drawImage(cameraScanVideo, 0, 0, width, height);
      const frame = context.getImageData(0, 0, width, height);
      const hit = jsQR(frame.data, width, height, { inversionAttempts: "attemptBoth" });
      if (hit?.data) {
        playScanFeedbackSound();
        closeCameraScanModal(String(hit.data).trim());
        return;
      }
      setCameraScanStatus("Point the camera at a QR code.");
    } catch {
      setCameraScanStatus("Scanning...");
    }
    cameraScanFrameId = requestAnimationFrame(scan);
  };

  cameraScanFrameId = requestAnimationFrame(scan);
}

async function startBarcodeDetectionLoop() {
  if (!cameraScanVideo) return;
  const Detector = window.BarcodeDetector;
  if (Detector) {
    startBarcodeDetectorLoop(Detector);
    return;
  }
  setCameraScanStatus("Loading QR decoder for this browser...");
  try {
    const jsQR = await loadJsQrLibrary();
    if (!cameraScanResolver || !cameraScanVideo?.srcObject) return;
    startJsQrDetectionLoop(jsQR);
  } catch {
    setCameraScanStatus("Camera QR scan is not available here. Use photo scan or manual entry.");
  }
}

async function requestCameraStream() {
  const attempts = [
    { video: { facingMode: { exact: "environment" } }, audio: false },
    { video: { facingMode: { ideal: "environment" } }, audio: false },
    { video: true, audio: false },
  ];
  let lastError = null;
  for (const constraints of attempts) {
    try {
      return await navigator.mediaDevices.getUserMedia(constraints);
    } catch (error) {
      lastError = error;
      if (error?.name === "NotAllowedError" || error?.name === "SecurityError") break;
    }
  }
  throw lastError || new Error("Unable to access camera stream");
}

function imageDataFromFile(file) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const blobUrl = URL.createObjectURL(file);
    img.onload = () => {
      const width = img.naturalWidth || img.width;
      const height = img.naturalHeight || img.height;
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d", { willReadFrequently: true });
      if (!context) {
        URL.revokeObjectURL(blobUrl);
        reject(new Error("Canvas context unavailable"));
        return;
      }
      context.drawImage(img, 0, 0, width, height);
      const frame = context.getImageData(0, 0, width, height);
      URL.revokeObjectURL(blobUrl);
      resolve(frame);
    };
    img.onerror = () => {
      URL.revokeObjectURL(blobUrl);
      reject(new Error("Unable to read selected image"));
    };
    img.src = blobUrl;
  });
}

async function decodeQrFromImageFile(file) {
  if (!file) return "";
  const jsQR = await loadJsQrLibrary();
  const frame = await imageDataFromFile(file);
  const hit = jsQR(frame.data, frame.width, frame.height, { inversionAttempts: "attemptBoth" });
  return hit?.data ? String(hit.data).trim() : "";
}

async function openCameraScanner({ title = "Scan QR code" } = {}) {
  if (!navigator.mediaDevices?.getUserMedia) {
    const manual = window.prompt("Camera not available on this device. Paste/enter QR code manually:");
    return manual ? String(manual).trim() : "";
  }

  if (!cameraScanModal || !cameraScanVideo) {
    const manual = window.prompt("Scanner UI not available. Paste/enter QR code manually:");
    return manual ? String(manual).trim() : "";
  }

  return new Promise(async (resolve) => {
    cameraScanResolver = resolve;
    cameraScanTitle.textContent = title;
    setCameraScanStatus("Starting camera...");
    cameraScanModal.classList.remove("hidden");

    try {
      cameraScanStream = await requestCameraStream();
      cameraScanVideo.srcObject = cameraScanStream;
      await cameraScanVideo.play();
      await startBarcodeDetectionLoop();
    } catch (error) {
      setCameraScanStatus(cameraAccessErrorMessage(error));
    }
  });
}

async function scanQrIntoInput(targetInput, title) {
  await primeScanFeedbackAudio();
  const scanned = await openCameraScanner({ title });
  if (!scanned) return "";
  targetInput.value = scanned;
  targetInput.dispatchEvent(new Event("input", { bubbles: true }));
  return scanned;
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
  const record = findQrRecord(qrCode);
  if (!record) return null;
  const { container, item } = record;

  const client = clientById(container.clientId);
  const project = projectById(container.projectId);
  const linkedUnit = units.find((entry) => entry.id === item.unitId) || null;
  const receiptUnits = units.filter((entry) => entry.siteReceipts.some((receipt) => receipt.qrCode === qrCode));
  const unit = linkedUnit || receiptUnits[0] || null;
  const events = [];

  const flowEvents = ensureArray(item.flowEvents);
  flowEvents.forEach((entry) => {
    const issueLabel = issueTypeLabel(entry.issueType);
    const detailParts = [];
    if (entry.unitLabel) detailParts.push(`Unit ${entry.unitLabel}`);
    if (issueLabel) detailParts.push(`Issue: ${issueLabel}`);
    if (entry.note) detailParts.push(entry.note);
    if (entry.byName) {
      const actorAccessProfile = entry.byAccessProfile ? roleLabel(entry.byAccessProfile) : "";
      detailParts.push(`By ${entry.byName}${actorAccessProfile ? ` (${actorAccessProfile})` : ""}`);
    }
    events.push({
      at: entry.at,
      source: qrFlowSourceLabel(entry.source || entry.stage || "workflow"),
      action: `${workflowStageLabel(entry.stage)} • ${workflowStatusLabel(entry.status)}`,
      detail: detailParts.join(" | ") || `${item.code || "-"} | ${item.description || "-"}`,
    });
  });

  if (!flowEvents.length) {
    events.push({
      at: item.createdAt,
      source: "Manufacturer Manifest",
      action: "Warehouse • Generated",
      detail: `${item.code || "-"} | ${item.description || "-"}`,
    });
    if (item.assignedAt) {
      events.push({
        at: item.assignedAt,
        source: "Warehouse Dispatch",
        action: "Delivery • In progress",
        detail: `${item.unitLabel || "-"}${item.kitchenType ? ` | ${item.kitchenType}` : ""}`,
      });
    }
    if (item.deliveredAt) {
      const legacyIssueLabel = issueTypeLabel(item.issueType);
      events.push({
        at: item.deliveredAt,
        source: "Site Receive",
        action: item.issueType ? "Delivery • Issue" : "Delivery • Completed",
        detail: legacyIssueLabel ? `${legacyIssueLabel}${item.issueNote ? ` | ${item.issueNote}` : ""}` : "OK",
      });
    }
  }

  ensureArray(container.replacementQueue)
    .filter((entry) => entry.qrCode === qrCode)
    .forEach((entry) => {
      const issueLabel = issueTypeLabel(entry.issueType) || entry.issueType;
      events.push({
        at: entry.createdAt,
        source: "Warehouse",
        action: "Replacement queue",
        detail: `${issueLabel}${entry.note ? ` | ${entry.note}` : ""}`,
      });
    });

  ensureArray(container.factoryMissingQueue)
    .filter((entry) => entry.qrCode === qrCode)
    .forEach((entry) => {
      const issueLabel = issueTypeLabel(entry.issueType) || entry.issueType;
      events.push({
        at: entry.createdAt,
        source: "Factory Claim",
        action: "Factory missing queue",
        detail: `${issueLabel}${entry.note ? ` | ${entry.note}` : ""}`,
      });
    });

  receiptUnits.forEach((receiptUnit) => {
    receiptUnit.siteReceipts
      .filter((entry) => entry.qrCode === qrCode)
      .forEach((entry) => {
        const issueLabel = issueTypeLabel(entry.issue) || entry.issue;
        events.push({
          at: entry.createdAt,
          source: "Site Receive",
          action: "Receipt record",
          detail: `${entry.item} (${entry.qty}) - ${issueLabel}${entry.note ? ` | ${entry.note}` : ""}`,
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
          action: "Unit audit log",
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
      <td>${escapeHtml(workflowStageLabel(item.flowStage || "warehouse"))}</td>
      <td>${escapeHtml(workflowStatusLabel(item.flowStatus || item.status))}</td>
      <td>${escapeHtml(item.unitLabel || "-")}</td>
      <td>${escapeHtml(item.kitchenType || "-")}</td>
      <td>${escapeHtml(fmtDate(item.updatedAt || item.createdAt))}</td>
      <td><div class="actions-inline"><button class="secondary xs-btn" data-qr-print="${item.id}">PDF/Print</button><button class="secondary xs-btn" data-qr-png="${item.id}">PNG</button></div></td>
    </tr>`
    )
    .join("");
  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>QR</th><th>SKU</th><th>Item</th><th>Stage</th><th>Status</th><th>Unit</th><th>Kitchen Type</th><th>Updated</th><th>Label</th></tr></thead><tbody>${
    rows || '<tr><td colspan="9">No QR items generated.</td></tr>'
  }</tbody></table>`;
}

function renderQueueTable(wrapper, list) {
  if (!wrapper) return;
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
      <td>${escapeHtml(issueTypeLabel(entry.issueType) || entry.issueType || "-")}</td>
      <td>${escapeHtml(entry.note || "-")}</td>
    </tr>`
    )
    .join("");
  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>Date</th><th>QR</th><th>SKU</th><th>Description</th><th>Unit</th><th>Issue</th><th>Detail</th></tr></thead><tbody>${
    rows || '<tr><td colspan="7">No records.</td></tr>'
  }</tbody></table>`;
}

function materialQueueSummaryRows(listKey) {
  const rows = [];
  containers.forEach((container) => {
    const fallbackClient = clientById(container.clientId)?.name || "-";
    const fallbackProject = projectById(container.projectId)?.name || "-";
    ensureArray(container[listKey]).forEach((entry) => {
      rows.push({
        createdAt: entry.createdAt || "",
        containerCode: container.containerCode || "-",
        clientName: clientById(entry.clientId || container.clientId)?.name || fallbackClient,
        projectName: projectById(entry.projectId || container.projectId)?.name || fallbackProject,
        qrCode: entry.qrCode || "-",
        code: entry.code || "-",
        description: entry.description || "-",
        unitLabel: entry.unitLabel || "-",
        issueType: issueTypeLabel(entry.issueType) || entry.issueType || "-",
        note: entry.note || "-",
      });
    });
  });
  rows.sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());
  return rows;
}

function renderMaterialQueueSummaryTable(wrapper, rows) {
  if (!wrapper) return;
  const body = rows
    .map(
      (entry) => `<tr>
      <td>${escapeHtml(fmtDate(entry.createdAt))}</td>
      <td>${escapeHtml(entry.containerCode)}</td>
      <td>${escapeHtml(entry.clientName)}</td>
      <td>${escapeHtml(entry.projectName)}</td>
      <td>${escapeHtml(entry.qrCode)}</td>
      <td>${escapeHtml(entry.code)}</td>
      <td>${escapeHtml(entry.description)}</td>
      <td>${escapeHtml(entry.unitLabel)}</td>
      <td>${escapeHtml(entry.issueType)}</td>
      <td>${escapeHtml(entry.note)}</td>
    </tr>`
    )
    .join("");
  wrapper.innerHTML = `<table class="data-table"><thead><tr><th>Date</th><th>Container</th><th>Client</th><th>Project</th><th>QR</th><th>SKU</th><th>Description</th><th>Unit</th><th>Issue</th><th>Detail</th></tr></thead><tbody>${
    body || '<tr><td colspan="10">No records.</td></tr>'
  }</tbody></table>`;
}

function renderManufactureCatalogSummaries() {
  renderMaterialQueueSummaryTable(materialsReplacementSummaryTable, materialQueueSummaryRows("replacementQueue"));
  renderMaterialQueueSummaryTable(materialsFactoryMissingSummaryTable, materialQueueSummaryRows("factoryMissingQueue"));
}

function findQrItemContainer(qrCode, projectId = "") {
  return containers.find((container) => {
    if (projectId && container.projectId !== projectId) return false;
    return ensureArray(container.qrItems).some((item) => item.qrCode === qrCode);
  });
}

function qrIssueRoutesToFactory(issueType) {
  const normalized = String(issueType || "").trim().toLowerCase();
  return ["nao-chegou", "faltando-pecas", "not-received", "missing-parts"].includes(normalized);
}

function isAdmin() {
  return userAccessProfile(currentUser) === "admin";
}

function isProjectManager() {
  return userAccessProfile(currentUser) === "project-manager";
}

function isDeveloper() {
  return userAccessProfile(currentUser) === "developer";
}

function normalizeCompanyName(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function isTagAdminUser(user = currentUser) {
  if (!user || userAccessProfile(user) !== "admin") return false;
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
  if (userAccessProfile(targetUser) === "developer" || isPrimaryDeveloperUser(targetUser)) return false;
  if (!isAdmin()) return targetUser.id === currentUser.id;
  if (isTagAdminUser(currentUser)) return true;

  const myCompany = normalizeCompanyName(currentUser.companyName);
  if (!myCompany) return targetUser.id === currentUser.id;
  return normalizeCompanyName(targetUser.companyName) === myCompany || targetUser.id === currentUser.id;
}

function can(action, stageKey = "") {
  if (!currentUser) return false;
  const accessProfile = userAccessProfile(currentUser);
  if (isDeveloper()) return true;
  if (isProjectManager()) {
    if (action === "sync") return false;
    if (action === "manageUsers") return false;
    if (action === "report") return true;
    if (action === "toggleStage") return true;
    return [
      "manageCatalog",
      "manageContainers",
      "manageContainerManifest",
      "manageContainerRelease",
      "manageContainerIssues",
      "dispatchTasks",
      "siteReceive",
      "createUnit",
      "quality",
      "install",
      "notes",
      "issues",
      "photos",
      "checklist",
    ].includes(action);
  }
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
  if (accessProfile === "visitor") return action === "report";

  if (action === "manageUsers") return false;
  if (action === "sync") return false;
  if (action === "manageCatalog") return false;
  if (action === "manageContainers") return ["warehouse", "transport"].includes(accessProfile);
  if (action === "manageContainerManifest") return ["warehouse", "transport"].includes(accessProfile);
  if (action === "manageContainerRelease") return accessProfile === "warehouse";
  if (action === "manageContainerIssues") return ["warehouse", "transport", "distribution", "foreman"].includes(accessProfile);
  if (action === "dispatchTasks") return accessProfile === "warehouse";
  if (action === "siteReceive") return ["transport", "distribution", "foreman", "installer"].includes(accessProfile);
  if (action === "deleteContainer") return false;
  if (action === "deleteUnit") return false;
  if (action === "createUnit") return accessProfile === "warehouse";
  if (action === "toggleStage") {
    if (stageKey === "quality") return false;
    return STAGE_ROLE_ACCESS[accessProfile]?.includes(stageKey) || false;
  }
  // Quality control and punch list decisions are restricted to Admin / Project Manager / Developer.
  if (action === "quality") return false;
  if (action === "install") return ["foreman", "installer"].includes(accessProfile);
  if (action === "notes") {
    if (currentProjectSector === "punchlist") return false;
    return ["distribution", "foreman", "installer"].includes(accessProfile);
  }
  if (action === "issues") {
    if (currentProjectSector === "punchlist") return false;
    return ["distribution", "foreman", "installer"].includes(accessProfile);
  }
  if (action === "photos") {
    if (currentProjectSector === "punchlist") return false;
    return ["distribution", "foreman", "installer"].includes(accessProfile);
  }
  if (action === "checklist") return ["warehouse", "transport", "distribution", "foreman"].includes(accessProfile);
  if (action === "report") return true;

  return false;
}

function rolePermissionsSummary(accessProfile) {
  if (accessProfile === "developer") return "Developer access: full control, including system structure and settings.";
  if (accessProfile === "admin") return "Admin access: can modify all project phases and quality control approvals/rejections.";
  if (accessProfile === "project-manager") return "Project Manager access: can modify all project phases and quality control approvals/rejections.";
  if (accessProfile === "warehouse") return "Warehouse access: unit creation, manifest, release/hold, and warehouse dispatch tasks.";
  if (accessProfile === "transport") return "Transport access: delivery and site receive updates.";
  if (accessProfile === "distribution") return "Distribution access: distribution-stage updates and field coordination.";
  if (accessProfile === "foreman") return "Foreman access: distribution/installation updates and execution supervision.";
  if (accessProfile === "qa") return "QA access: read/track quality records. Final QC decisions are by Admin or Project Manager.";
  if (accessProfile === "installer") return "Installer access: installation execution updates.";
  if (accessProfile === "visitor") return "Visitor access: read-only and reports. Access profiles are assigned by Admin/Developer.";
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

function isSha256Hash(value) {
  return /^[a-f0-9]{64}$/i.test(String(value || "").trim());
}

function userLegacyPassword(user) {
  const legacy = String(user?.legacyPassword || "").trim();
  if (legacy) return legacy;
  const oldPassword = String(user?.password || "").trim();
  if (oldPassword) return oldPassword;
  return "";
}

function userPasswordMatches(user, plainPassword, sha256Hash) {
  const stored = String(user?.passwordHash || "").trim();
  const inputHash = String(sha256Hash || "").trim().toLowerCase();

  if (stored) {
    if (isSha256Hash(stored)) return stored.toLowerCase() === inputHash;
    return stored === plainPassword || stored.toLowerCase() === inputHash;
  }

  const legacy = userLegacyPassword(user);
  if (!legacy) return false;
  return legacy === plainPassword || legacy.toLowerCase() === inputHash;
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
  if (currentUser) {
    startAutoPullLoop();
    authView.classList.add("hidden");
    appMain.classList.remove("hidden");
    userBadge.classList.remove("hidden");
    editProfileBtn?.classList.remove("hidden");
    logoutBtn.classList.remove("hidden");
    userBadge.textContent = `${currentUser.name} (${roleLabel(userAccessProfile(currentUser))})`;
    currentView = viewFromHash();
    render();
    void maybeAutoDispatchCoiReminder();
    return;
  }

  stopAutoPullLoop();
  appMain.classList.add("hidden");
  authView.classList.remove("hidden");
  userBadge.classList.add("hidden");
  editProfileBtn?.classList.add("hidden");
  logoutBtn.classList.add("hidden");
  setupPanel?.classList.add("hidden");
  showSignupMode(false);
  applyLanguageToUi();
}

function setFormEnabled(form, enabled) {
  if (!form) return;
  form.querySelectorAll("input, select, textarea, button").forEach((field) => {
    field.disabled = !enabled;
  });
}

function viewFromHash() {
  const raw = window.location.hash.replace(/^#/, "").trim();
  if (!raw) return "home";
  const allowed = new Set(["home", "sync", "users", "clients", "projects", "manufacture", "userEdit", "ocrImporter"]);
  return allowed.has(raw) ? raw : "home";
}

function canOpenManufactureWorkspace() {
  return (
    can("manageContainers") ||
    can("manageCatalog") ||
    can("manageContainerManifest") ||
    can("manageContainerRelease") ||
    can("manageContainerIssues")
  );
}

function canOpenView(view) {
  if (!currentUser) return false;
  if (view === "home") return true;
  if (view === "sync") return can("sync");
  if (view === "users") return can("manageUsers");
  if (view === "userEdit") return true;
  if (view === "clients") return can("manageCatalog");
  if (view === "projects") return true;
  if (view === "manufacture") return canOpenManufactureWorkspace();
  if (view === "ocrImporter") return can("manageCatalog") || can("manageContainerManifest") || isAdmin() || isDeveloper();
  return false;
}

function allowedProjectSectorsForAccessProfile(accessProfile) {
  if (accessProfile === "developer" || accessProfile === "admin" || accessProfile === "project-manager") return PROJECT_SECTORS;
  if (accessProfile === "visitor") return ["delivery", "distribuicao"];
  if (accessProfile === "warehouse") return ["warehouse"];
  if (accessProfile === "transport") return ["delivery"];
  if (accessProfile === "distribution") return ["distribuicao"];
  if (accessProfile === "foreman") return ["distribuicao", "instalacao"];
  if (accessProfile === "installer") return ["distribuicao", "instalacao"];
  if (accessProfile === "qa") return ["punchlist"];
  return ["delivery"];
}

function allowedProjectSectors() {
  return allowedProjectSectorsForAccessProfile(userAccessProfile(currentUser));
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

function nodeVisibleForSector(node, sectorKey = currentProjectSector) {
  const list = String(node?.dataset?.sectors || "")
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
  if (!list.length) return true;
  return list.includes(sectorKey);
}

function applySectorVisibility(root, sectorKey = currentProjectSector) {
  if (!root) return;
  root.querySelectorAll("[data-sectors]").forEach((node) => {
    node.classList.toggle("hidden-view", !nodeVisibleForSector(node, sectorKey));
  });
}

function setManufactureSubView(view = "schedule") {
  const normalized = ["schedule", "catalog", "solicitation"].includes(view) ? view : "schedule";
  const previous = manufactureSubView;
  manufactureSubView = normalized;
  if (previous !== manufactureSubView && currentUser) {
    pushAppAudit(`Navigation manufacture folder: ${previous} -> ${manufactureSubView}`, "navigation", "manufacture");
  }
}

function applyManufactureSubviewVisibility() {
  const inManufacture = currentView === "manufacture";
  const active = ["schedule", "catalog", "solicitation"].includes(manufactureSubView) ? manufactureSubView : "schedule";
  manufactureScheduleView?.classList.toggle("hidden-view", inManufacture ? active !== "schedule" : false);
  manufactureCatalogView?.classList.toggle("hidden-view", inManufacture ? active !== "catalog" : true);
  manufactureSolicitationView?.classList.toggle("hidden-view", inManufacture ? active !== "solicitation" : true);
}

function normalizeWarehouseSubView(view = "") {
  return ["schedule", "inventory", "qr", "operations"].includes(view) ? view : "operations";
}

function setWarehouseSubView(view = "operations") {
  const normalized = normalizeWarehouseSubView(view);
  const previous = warehouseSubView;
  warehouseSubView = normalized;
  if (previous !== warehouseSubView && currentUser) {
    pushAppAudit(`Navigation warehouse folder: ${previous} -> ${warehouseSubView}`, "navigation", "warehouse");
  }
}

function applyMasterSubpanelMode() {
  if (!clientsSubpanel || !projectsSubpanel || !contactsSubpanel) return;
  const showClients = currentView === "clients";
  const showProjects = currentView === "projects";
  masterDataPanel?.classList.toggle("clients-mode", showClients);

  if (showClients) {
    const showHub = clientsWorkspaceMode === "hub";
    const showNewClients = clientsWorkspaceMode === "newClients";
    const showProjectsCatalog = clientsWorkspaceMode === "projects";
    const showContracts = clientsWorkspaceMode === "contracts";
    clientsSubpanel.classList.toggle("hidden-view", !(showHub || showNewClients));
    projectsSubpanel.classList.toggle("hidden-view", !showProjectsCatalog);
    contactsSubpanel.classList.toggle("hidden-view", !showProjectsCatalog);
    contractsSubpanel?.classList.toggle("hidden-view", !showContracts);
    return;
  }

  if (showProjects) {
    clientsSubpanel.classList.add("hidden-view");
    contractsSubpanel?.classList.add("hidden-view");
    const showProjectCatalog = projectsViewMode === "overview";
    projectsSubpanel.classList.toggle("hidden-view", !showProjectCatalog);
    contactsSubpanel.classList.toggle("hidden-view", !showProjectCatalog);
    return;
  }

  clientsSubpanel.classList.remove("hidden-view");
  projectsSubpanel.classList.remove("hidden-view");
  contactsSubpanel.classList.remove("hidden-view");
  contractsSubpanel?.classList.remove("hidden-view");
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
  const inManufacture = currentView === "manufacture";
  const inWarehouseWorkspace = inProjects && ["warehouse", "delivery"].includes(currentProjectSector);
  const activeWarehouseSubView = normalizeWarehouseSubView(warehouseSubView);
  projectSectorPanel?.classList.toggle("hidden-view", !inProjects);
  const containerVisible = inManufacture || (inProjects && ["fabrica", "warehouse"].includes(currentProjectSector));
  let deliveryInventoryVisible = inProjects && ["delivery"].includes(currentProjectSector);
  let unitEntryVisible = inProjects && ["warehouse"].includes(currentProjectSector);
  let unitsVisible =
    inProjects && ["warehouse", "delivery", "distribuicao", "instalacao", "punchlist"].includes(currentProjectSector);
  let qrLookupVisible = unitsVisible;

  if (inWarehouseWorkspace) {
    if (activeWarehouseSubView === "inventory") {
      deliveryInventoryVisible = currentProjectSector === "delivery";
      unitEntryVisible = false;
      unitsVisible = false;
      qrLookupVisible = false;
    } else if (activeWarehouseSubView === "qr") {
      deliveryInventoryVisible = currentProjectSector === "delivery";
      unitEntryVisible = false;
      unitsVisible = false;
      qrLookupVisible = true;
    } else if (activeWarehouseSubView === "schedule") {
      deliveryInventoryVisible = false;
      unitEntryVisible = false;
      unitsVisible = true;
      qrLookupVisible = false;
    } else {
      deliveryInventoryVisible = false;
      unitEntryVisible = currentProjectSector === "warehouse";
      unitsVisible = true;
      qrLookupVisible = false;
    }
  }

  containerPanel.classList.toggle("hidden-view", !containerVisible);
  deliveryInventoryPanel?.classList.toggle("hidden-view", !deliveryInventoryVisible);
  const showInventoryCatalog = deliveryInventoryVisible && activeWarehouseSubView !== "qr";
  const showInventoryScan = deliveryInventoryVisible && activeWarehouseSubView === "qr";
  deliveryInventoryCatalogBlock?.classList.toggle("hidden-view", !showInventoryCatalog);
  deliveryInventoryScanBlock?.classList.toggle("hidden-view", !showInventoryScan);
  deliveryInventoryTable?.classList.toggle("hidden-view", !showInventoryCatalog);
  unitEntryPanel.classList.toggle("hidden-view", !unitEntryVisible);
  unitsToolbarPanel.classList.toggle("hidden-view", !unitsVisible);
  qrLookupPanel.classList.toggle("hidden-view", !qrLookupVisible);
  unitsContainer.classList.toggle("hidden-view", !unitsVisible);
  if (containerVisible) {
    applySectorVisibility(containerPanel, inManufacture ? "fabrica" : currentProjectSector);
  }
  applyManufactureSubviewVisibility();
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
    manufacture: [containerPanel],
    ocrImporter: [ocrImporterPanel],
    projects: [
      masterDataPanel,
      projectSectorPanel,
      containerPanel,
      deliveryInventoryPanel,
      unitEntryPanel,
      unitsToolbarPanel,
      qrLookupPanel,
      unitsContainer,
    ],
  };

  const allPanels = [
    homePanel,
    syncPanel,
    usersPanel,
    userEditPanel,
    ocrImporterPanel,
    masterDataPanel,
    projectSectorPanel,
    containerPanel,
    deliveryInventoryPanel,
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
  if (currentView === "clients") {
    if (clientsWorkspaceMode === "newClients") {
      selectedClientDetailsProjectId = "";
      keepClientFormBlank = true;
      setClientsSectionMode("create");
    } else if (clientsWorkspaceMode === "hub") {
      setClientsSectionMode("hub");
    }
  }
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

function partnerCategoryFromUser(user) {
  const text = `${user?.companyName || ""} ${user?.jobTitle || ""}`.toLowerCase();
  if (/(factory|manufactur|fabrica)/.test(text)) return "manufacturers";
  if (/(import|importadora)/.test(text)) return "importers";
  if (/(truck|delivery|trucking)/.test(text)) return "truck-deliveries";
  if (/(transport|carrier|logistic|freight|shipping|transportadora)/.test(text)) return "carriers";
  if (/(material|construction|supply|lumber|tile|wood|cabinet)/.test(text)) return "construction-materials";
  return "uncategorized";
}

function usersFilterMatches(user) {
  if (usersViewFilter === "subcontractor") return user.employmentType === "subcontractor";
  if (usersViewFilter === "partner") return user.employmentType === "supplier";
  if (usersViewFilter === "partner-manufacturers")
    return user.employmentType === "supplier" && partnerCategoryFromUser(user) === "manufacturers";
  if (usersViewFilter === "partner-construction-materials")
    return user.employmentType === "supplier" && partnerCategoryFromUser(user) === "construction-materials";
  if (usersViewFilter === "partner-importers")
    return user.employmentType === "supplier" && partnerCategoryFromUser(user) === "importers";
  if (usersViewFilter === "partner-carriers")
    return user.employmentType === "supplier" && partnerCategoryFromUser(user) === "carriers";
  if (usersViewFilter === "partner-truck-deliveries")
    return user.employmentType === "supplier" && partnerCategoryFromUser(user) === "truck-deliveries";
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

function openWarehouseWorkspace(subView = "operations", preferredSector = "warehouse") {
  if (!canOpenView("projects")) {
    setView("home");
    return;
  }
  const allowed = allowedProjectSectors();
  let targetSector = preferredSector;
  if (!allowed.includes(targetSector)) {
    if (allowed.includes("delivery")) targetSector = "delivery";
    else if (allowed.includes("warehouse")) targetSector = "warehouse";
    else targetSector = allowed[0] || "delivery";
  }
  projectsViewMode = "operations";
  setUsersViewFilter("all");
  setWarehouseSubView(subView);
  setView("projects");
  setProjectSector(targetSector);
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

function openProjectWarehouse(projectId) {
  const project = projects.find((entry) => entry.id === projectId);
  if (!project || !canOpenView("projects")) {
    setView("home");
    return;
  }

  selectedProjectId = project.id;
  selectedClientId = project.clientId || selectedClientId;
  projectsViewMode = "operations";
  setWarehouseSubView("operations");
  setUsersViewFilter("all");
  setView("projects");
  setProjectSector("warehouse");

  if (projectReportSelect) projectReportSelect.value = project.id;
  if (unitProjectSelect) {
    unitProjectSelect.value = project.id;
    syncUnitProjectInfo();
  }
  if (containerClientSelect) {
    containerClientSelect.value = project.clientId || "";
    syncContainerProjectSelect();
    if (containerProjectSelect) containerProjectSelect.value = project.id;
  }
  if (contactClientSelect) {
    contactClientSelect.value = project.clientId || "";
    syncContactProjectSelect();
    if (contactProjectSelect) contactProjectSelect.value = project.id;
  }

  searchInput.value = project.name || "";
  renderUnits();
  pushAppAudit(`Navigation to warehouse from project link: ${project.name}`, "navigation", project.name || "-");
  unitEntryPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function openProjectCatalog(projectId) {
  const project = projects.find((entry) => entry.id === projectId);
  if (!project || !canOpenView("projects")) {
    setView("home");
    return;
  }

  selectedProjectId = project.id;
  selectedClientId = project.clientId || selectedClientId;
  projectsViewMode = "overview";
  setUsersViewFilter("all");
  setView("projects");

  if (projectClientSelect) projectClientSelect.value = project.clientId || "";
  if (contactClientSelect) {
    contactClientSelect.value = project.clientId || "";
    syncContactProjectSelect();
    if (contactProjectSelect) contactProjectSelect.value = project.id;
  }
  if (contractClientSelect) {
    contractClientSelect.value = project.clientId || "";
    syncContractProjectSelect();
    if (contractProjectSelect) contractProjectSelect.value = project.id;
  }
  populateProjectForm(project);
  renderProjectsTable();
  renderContactsTable();
  renderContractsTable();
  projectsSubpanel?.scrollIntoView({ behavior: "smooth", block: "center" });
  pushAppAudit(`Navigation to projects catalog: ${project.name}`, "navigation", project.name || "-");
}

function openClientsHub() {
  if (!can("manageCatalog")) {
    setView("home");
    return;
  }
  setUsersViewFilter("all");
  setClientsWorkspaceMode("hub");
  setView("clients");
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
    setUsersSubView("directory");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "usersRegistration") {
    if (!can("manageUsers")) {
      setView("home");
      return;
    }
    setUsersViewFilter("all");
    openUsersRegistrationClean();
    setView("users");
    return;
  }

  if (action === "usersPeople") {
    if (!can("manageUsers")) {
      setView("home");
      return;
    }
    setUsersViewFilter("all");
    setUsersSubView("directory");
    setView("users");
    return;
  }

  if (action === "clients") {
    openClientsHub();
    return;
  }

  if (action === "clientsNew") {
    if (!can("manageCatalog")) {
      setView("home");
      return;
    }
    setUsersViewFilter("all");
    setClientsWorkspaceMode("newClients");
    setView("clients");
    return;
  }

  if (action === "clientsProjects") {
    if (!can("manageCatalog")) {
      setView("home");
      return;
    }
    setUsersViewFilter("all");
    setClientsWorkspaceMode("projects");
    setView("clients");
    projectsSubpanel?.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  if (action === "clientsContracts") {
    if (!can("manageCatalog")) {
      setView("home");
      return;
    }
    setUsersViewFilter("all");
    setClientsWorkspaceMode("contracts");
    setView("clients");
    contractsSubpanel?.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  if (action === "projects") {
    projectsViewMode = "overview";
    setUsersViewFilter("all");
    setView(canOpenView("projects") ? "projects" : "home");
    return;
  }

  if (action === "ocrImporter") {
    setUsersViewFilter("all");
    setView(canOpenView("ocrImporter") ? "ocrImporter" : "home");
    return;
  }

  if (action === "subcontractors") {
    setUsersViewFilter("subcontractor");
    setUsersSubView("directory");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "partners") {
    setUsersViewFilter("partner");
    setUsersSubView("directory");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "partnersManufacturers") {
    setUsersViewFilter("partner-manufacturers");
    setUsersSubView("directory");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "partnersConstructionMaterials") {
    setUsersViewFilter("partner-construction-materials");
    setUsersSubView("directory");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "partnersImporters") {
    setUsersViewFilter("partner-importers");
    setUsersSubView("directory");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "partnersCarriers") {
    setUsersViewFilter("partner-carriers");
    setUsersSubView("directory");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "partnersTruckDeliveries") {
    setUsersViewFilter("partner-truck-deliveries");
    setUsersSubView("directory");
    setView(can("manageUsers") ? "users" : "home");
    return;
  }

  if (action === "manufacture") {
    setUsersViewFilter("all");
    setManufactureSubView("schedule");
    setView(canOpenView("manufacture") ? "manufacture" : "home");
    return;
  }

  if (action === "manufactureSchedule") {
    setUsersViewFilter("all");
    setManufactureSubView("schedule");
    setView(canOpenView("manufacture") ? "manufacture" : "home");
    return;
  }

  if (action === "manufactureCatalog") {
    setUsersViewFilter("all");
    setManufactureSubView("catalog");
    setView(canOpenView("manufacture") ? "manufacture" : "home");
    return;
  }

  if (action === "manufactureSolicitation") {
    setUsersViewFilter("all");
    setManufactureSubView("solicitation");
    setView(canOpenView("manufacture") ? "manufacture" : "home");
    return;
  }

  if (action === "warehouse") {
    openWarehouseWorkspace("operations", "warehouse");
    return;
  }

  if (action === "warehouseDeliverySchedule") {
    openWarehouseWorkspace("schedule", "delivery");
    return;
  }

  if (action === "warehouseDeliveryInventory") {
    openWarehouseWorkspace("inventory", "delivery");
    return;
  }

  if (action === "warehouseQrScan") {
    openWarehouseWorkspace("qr", "delivery");
    return;
  }

  if (action === "warehouseOperations") {
    openWarehouseWorkspace("operations", "warehouse");
    return;
  }

  if (action === "delivery") {
    openWarehouseWorkspace("schedule", "delivery");
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
    if (unitScopeWorkInput) unitScopeWorkInput.value = "";
    if (warehouseProjectSummary) warehouseProjectSummary.textContent = "Select a project to show floors, apartments, and unit checklist template.";
    return;
  }

  const project = projectById(projectId);
  const client = project ? clientById(project.clientId) : null;
  unitProjectName.value = project?.name || "";
  unitClientName.value = client?.name || "";
  unitJobSite.value = project?.address || "";
  if (unitScopeWorkInput) unitScopeWorkInput.value = projectScopeSummary(project);
  if (warehouseProjectSummary) {
    const templateText = projectChecklistTemplate(project).join(", ");
    warehouseProjectSummary.textContent = `Project details: Floors ${project?.floorsCount || "-"} | Apartments ${
      project?.apartmentsCount || "-"
    } | Unit checklist template: ${templateText || "-"}`;
  }
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

function syncContractProjectSelect() {
  const clientId = contractClientSelect?.value || "";
  if (!contractProjectSelect) return;
  populateProjectSelectByClient(contractProjectSelect, clientId, "Project (optional)", true);
}

function renderMasterSelects() {
  const previousProjectClient = projectClientSelect.value;
  const previousContactClient = contactClientSelect.value;
  const previousContactProject = contactProjectSelect.value;
  const previousContractClient = contractClientSelect?.value || "";
  const previousContractProject = contractProjectSelect?.value || "";

  populateSelect(projectClientSelect, clients, {
    includeEmpty: true,
    placeholder: clients.length ? "Select the client" : "Register a client",
  });
  if (previousProjectClient && clients.some((client) => client.id === previousProjectClient)) {
    projectClientSelect.value = previousProjectClient;
  } else if (selectedClientId && clients.some((client) => client.id === selectedClientId)) {
    projectClientSelect.value = selectedClientId;
  }

  populateSelect(contactClientSelect, clients, {
    includeEmpty: true,
    placeholder: "Client (optional)",
  });
  if (previousContactClient && clients.some((client) => client.id === previousContactClient)) {
    contactClientSelect.value = previousContactClient;
  } else if (selectedClientId && clients.some((client) => client.id === selectedClientId)) {
    contactClientSelect.value = selectedClientId;
  }

  syncContactProjectSelect();
  if (previousContactProject && projects.some((project) => project.id === previousContactProject)) {
    contactProjectSelect.value = previousContactProject;
  }

  if (contractClientSelect) {
    populateSelect(contractClientSelect, clients, {
      includeEmpty: true,
      placeholder: "Select client",
    });
    if (previousContractClient && clients.some((client) => client.id === previousContractClient)) {
      contractClientSelect.value = previousContractClient;
    } else if (selectedClientId && clients.some((client) => client.id === selectedClientId)) {
      contractClientSelect.value = selectedClientId;
    }
  }

  syncContractProjectSelect();
  if (contractProjectSelect && previousContractProject && projects.some((project) => project.id === previousContractProject)) {
    contractProjectSelect.value = previousContractProject;
  }

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
    .map((project) => {
      const client = clientById(project.clientId);
      return `<option value="${escapeHtml(project.id)}">${escapeHtml(project.name)}${client ? ` - ${escapeHtml(client.name)}` : ""}</option>`;
    })
    .join("")}`;
  if (projects.some((project) => project.id === previousReport)) projectReportSelect.value = previousReport;

  syncUnitProjectInfo();
}

function setClientsSectionMode(mode = "create") {
  const normalized = mode === "details" || mode === "hub" ? mode : "create";
  const detailsMode = normalized === "details";
  const hubMode = normalized === "hub";
  clientsCreateView?.classList.toggle("hidden", detailsMode || hubMode);
  clientDetailsView?.classList.toggle("hidden", !detailsMode);
  clientsHubView?.classList.toggle("hidden", !hubMode);
}

function setClientsWorkspaceMode(mode = "hub") {
  const normalized = ["hub", "newClients", "projects", "contracts"].includes(mode) ? mode : "hub";
  clientsWorkspaceMode = normalized;
  if (normalized === "newClients") {
    selectedClientDetailsProjectId = "";
    keepClientFormBlank = true;
    setClientsSectionMode("create");
    return;
  }
  if (normalized === "hub") setClientsSectionMode("hub");
}

function renderClientsHubList() {
  if (!clientsHubList) return;
  const rows = clients
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
    .map(
      (client) => `<button class="client-summary-item${client.id === selectedClientId ? " active" : ""}" type="button" data-client-hub="${client.id}">
        ${escapeHtml(client.name || "-")}
      </button>`
    )
    .join("");
  clientsHubList.innerHTML = rows || '<p class="hint">No clients registered yet.</p>';
  clientsHubList.querySelectorAll("[data-client-hub]").forEach((button) => {
    button.addEventListener("click", () => {
      selectedClientId = button.dataset.clientHub || "";
      selectedProjectId = "";
      selectedContractId = "";
      selectedClientDetailsProjectId = "";
      renderClientsHubList();
      renderProjectsTable();
      renderContactsTable();
      renderContractsTable();
    });
  });
}

function renderClientSelectedProjectDetails(project) {
  if (!clientProjectDetailsTable) return;
  if (!project) {
    clientProjectDetailsTable.innerHTML = '<p class="hint">Select a project to view project details and full scope of work.</p>';
    return;
  }

  const client = clientById(project.clientId);
  const scopeBase = ensureArray(project.scopeCategories).map((key) => PROJECT_SCOPE_LABELS[key] || key);
  const scopeExtra = uniqueTextList(project.scopeExtras);
  const scopeAll = uniqueTextList([...scopeBase, ...scopeExtra]);
  const checklistTemplate = projectChecklistTemplate(project);

  const detailRows = [
    ["Client", client?.name || "-"],
    ["Project Name", project.name || "-"],
    ["Project Code", project.code || "-"],
    ["Address / Job Site", project.address || "-"],
    ["Floors", project.floorsCount || "-"],
    ["Apartments", project.apartmentsCount || "-"],
  ]
    .map(([label, value]) => `<tr><th>${escapeHtml(label)}</th><td>${escapeHtml(value)}</td></tr>`)
    .join("");

  const scopeRows = scopeAll.length
    ? scopeAll.map((entry) => `<tr><td>${escapeHtml(entry)}</td></tr>`).join("")
    : '<tr><td>No scope items registered.</td></tr>';

  const checklistRows = checklistTemplate.length
    ? checklistTemplate.map((entry) => `<tr><td>${escapeHtml(entry)}</td></tr>`).join("")
    : '<tr><td>No checklist template items registered.</td></tr>';

  clientProjectDetailsTable.innerHTML = `
    <table class="data-table">
      <thead><tr><th>Field</th><th>Value</th></tr></thead>
      <tbody>${detailRows}</tbody>
    </table>
    <table class="data-table">
      <thead><tr><th>Scope of Work (all items)</th></tr></thead>
      <tbody>${scopeRows}</tbody>
    </table>
    <table class="data-table">
      <thead><tr><th>Unit checklist template</th></tr></thead>
      <tbody>${checklistRows}</tbody>
    </table>
  `;
}

function renderClientProjectsList(client) {
  if (!clientProjectsList) return;
  if (!client) {
    clientProjectsList.innerHTML = '<p class="hint">Select a client to view projects.</p>';
    renderClientSelectedProjectDetails(null);
    return;
  }

  const linkedProjects = projects
    .filter((project) => project.clientId === client.id)
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name));

  if (selectedClientDetailsProjectId && !linkedProjects.some((entry) => entry.id === selectedClientDetailsProjectId)) {
    selectedClientDetailsProjectId = "";
  }

  const rows = linkedProjects
    .map((project) => {
      const pmContact =
        contacts.find((entry) => entry.projectId === project.id && entry.role === "project-manager") ||
        contacts.find((entry) => entry.projectId === project.id && entry.role === "foreman") ||
        null;
      const pmName = pmContact?.name || client.contactSeniorProjectManager || "-";
      const pmPhone = pmContact?.phone || client.seniorProjectManagerPhone || "-";
      return `<tr class="${project.id === selectedClientDetailsProjectId ? "selected-row" : ""}">
        <td>${escapeHtml(project.name || "-")}</td>
        <td>${escapeHtml(project.address || "-")}</td>
        <td>${escapeHtml(project.code || "-")}</td>
        <td>${escapeHtml(pmName)}</td>
        <td>${escapeHtml(pmPhone)}</td>
        <td><button class="secondary xs-btn" type="button" data-client-detail-project="${project.id}">View details</button></td>
      </tr>`;
    })
    .join("");

  clientProjectsList.innerHTML = `<table class="data-table"><thead><tr><th>Project</th><th>Address</th><th>Code</th><th>PM</th><th>PM Phone</th><th>Action</th></tr></thead><tbody>${
    rows || '<tr><td colspan="6">No projects registered for this client.</td></tr>'
  }</tbody></table>`;

  clientProjectsList.querySelectorAll("[data-client-detail-project]").forEach((button) => {
    button.addEventListener("click", () => {
      const targetProject = projects.find((entry) => entry.id === button.dataset.clientDetailProject) || null;
      if (!targetProject) return;
      selectedClientDetailsProjectId = targetProject.id;
      renderClientProjectsList(client);
      renderClientSelectedProjectDetails(targetProject);
    });
  });

  const selectedProject = linkedProjects.find((entry) => entry.id === selectedClientDetailsProjectId) || null;
  renderClientSelectedProjectDetails(selectedProject);
}

function renderClientsTable() {
  if (!clientsSummaryScroll || !clientForm) return;

  const query = clientSearchQuery.trim().toLowerCase();
  const visibleClients = clients
    .filter((client) => {
      if (!query) return true;
      return String(client.name || "").toLowerCase().includes(query);
    })
    .slice()
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")));

  if (selectedClientId && !clients.some((entry) => entry.id === selectedClientId)) {
    selectedClientId = "";
    selectedContractId = "";
  }

  const listRows = visibleClients
    .map(
      (client) => `<button class="client-summary-item${client.id === selectedClientId ? " active" : ""}" type="button" data-client-summary="${client.id}">
        ${escapeHtml(client.name || "-")}
      </button>`
    )
    .join("");

  clientsSummaryScroll.innerHTML = listRows || '<p class="hint">No matching clients.</p>';

  clientsSummaryScroll.querySelectorAll("[data-client-summary]").forEach((button) => {
    button.addEventListener("click", () => {
      selectedClientId = button.dataset.clientSummary || "";
      selectedProjectId = "";
      selectedContractId = "";
      selectedClientDetailsProjectId = "";
      setClientsSectionMode("details");
      if (projectClientSelect && selectedClientId) projectClientSelect.value = selectedClientId;
      if (contactClientSelect && selectedClientId) {
        contactClientSelect.value = selectedClientId;
        syncContactProjectSelect();
      }
      if (contractClientSelect && selectedClientId) {
        contractClientSelect.value = selectedClientId;
        syncContractProjectSelect();
      }
      renderClientsTable();
      renderProjectsTable();
      renderContactsTable();
      renderContractsTable();
    });
  });

  const selectedClient = clients.find((entry) => entry.id === selectedClientId) || null;
  if (selectedClient && !clientDetailsView?.classList.contains("hidden")) {
    renderClientDetails(selectedClient);
    renderClientProjectsList(selectedClient);
  } else if (!selectedClient) {
    renderClientDetails(null);
    renderClientProjectsList(null);
  }

  const currentEditId = clientForm.elements.clientId?.value || "";
  if (!currentEditId && keepClientFormBlank) resetClientForm();
}

function splitLines(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function uniqueTextList(values) {
  const seen = new Set();
  const result = [];
  for (const raw of normalizeTextList(values)) {
    const key = raw.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(raw);
  }
  return result;
}

function projectScopeSummary(project) {
  if (!project) return "";
  const base = ensureArray(project.scopeCategories)
    .map((key) => PROJECT_SCOPE_LABELS[key] || key)
    .filter(Boolean);
  const extras = uniqueTextList(project.scopeExtras);
  return [...base, ...extras].join(", ");
}

function projectChecklistTemplate(project) {
  if (!project) return [];
  return uniqueTextList(project.unitChecklistTemplate);
}

function syncProjectDraftHiddenFields() {
  if (projectScopeExtrasJson) projectScopeExtrasJson.value = JSON.stringify(projectScopeExtrasDraft);
  if (projectChecklistExtrasJson) projectChecklistExtrasJson.value = JSON.stringify(projectChecklistExtrasDraft);
}

function renderProjectDraftPills() {
  if (projectScopeExtraList) {
    projectScopeExtraList.innerHTML = projectScopeExtrasDraft.length
      ? projectScopeExtrasDraft
          .map(
            (entry, index) =>
              `<div class="dynamic-check-row">
                <label><input type="checkbox" name="scopeExtrasDynamic" value="${escapeHtml(entry)}" checked /> ${escapeHtml(entry)}</label>
                <button class="danger xs-btn" type="button" data-project-scope-pill-remove="${index}" aria-label="Remove scope item">Remove</button>
              </div>`
          )
          .join("")
      : '<p class="hint">No extra scope items.</p>';
  }
  if (projectChecklistExtraList) {
    projectChecklistExtraList.innerHTML = projectChecklistExtrasDraft.length
      ? projectChecklistExtrasDraft
          .map(
            (entry, index) =>
              `<div class="dynamic-check-row">
                <label><input type="checkbox" name="unitChecklistExtrasDynamic" value="${escapeHtml(entry)}" checked /> ${escapeHtml(entry)}</label>
                <button class="danger xs-btn" type="button" data-project-checklist-pill-remove="${index}" aria-label="Remove checklist item">Remove</button>
              </div>`
          )
          .join("")
      : '<p class="hint">No extra checklist items.</p>';
  }
  syncProjectDraftHiddenFields();
}

function setProjectDrafts(scopeExtras = [], checklistExtras = []) {
  projectScopeExtrasDraft = uniqueTextList(scopeExtras);
  projectChecklistExtrasDraft = uniqueTextList(checklistExtras);
  renderProjectDraftPills();
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

function resetProjectForm({ clientId = "" } = {}) {
  if (!projectForm) return;
  projectForm.reset();
  if (projectForm.elements.projectId) projectForm.elements.projectId.value = "";
  if (clientId && projectClientSelect && clients.some((client) => client.id === clientId)) projectClientSelect.value = clientId;
  setCheckedValues(projectForm, "scopeCategories", []);
  setCheckedValues(projectForm, "unitChecklistBase", []);
  setProjectDrafts([], []);
  if (projectAdminTarget) projectAdminTarget.textContent = "Creating a new project.";
}

function populateProjectForm(project) {
  if (!projectForm || !project) return;
  if (projectForm.elements.projectId) projectForm.elements.projectId.value = project.id;
  projectForm.elements.clientId.value = project.clientId || "";
  projectForm.elements.name.value = project.name || "";
  projectForm.elements.code.value = project.code || "";
  projectForm.elements.address.value = project.address || "";
  projectForm.elements.floorsCount.value = project.floorsCount || "";
  projectForm.elements.apartmentsCount.value = project.apartmentsCount || "";
  setCheckedValues(projectForm, "scopeCategories", ensureArray(project.scopeCategories));

  const baseChecklist = projectChecklistTemplate(project).filter((item) =>
    Array.from(projectForm.querySelectorAll('input[name="unitChecklistBase"]')).some((entry) => entry.value === item)
  );
  const extraChecklist = projectChecklistTemplate(project).filter(
    (item) => !Array.from(projectForm.querySelectorAll('input[name="unitChecklistBase"]')).some((entry) => entry.value === item)
  );
  setCheckedValues(projectForm, "unitChecklistBase", baseChecklist);
  setProjectDrafts(project.scopeExtras, extraChecklist);

  if (projectAdminTarget) projectAdminTarget.textContent = `Editing project: ${project.name || "-"}`;
}

function openClientRegistrationShortcut(target = "projects") {
  if (!can("manageCatalog")) {
    setView("home");
    return;
  }
  if (target === "contracts") {
    setClientsWorkspaceMode("contracts");
    setView("clients");
    contractsSubpanel?.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  setClientsWorkspaceMode("projects");
  setView("clients");

  const selectedId = clientForm?.elements?.clientId?.value || selectedClientId || "";
  if (target === "projects") {
    if (selectedId) {
      selectedClientId = selectedId;
      if (projectClientSelect) projectClientSelect.value = selectedId;
    }
    projectsSubpanel?.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  if (selectedId && contactClientSelect) {
    selectedClientId = selectedId;
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

  const officeLines = splitLines(client.offices);
  const logoBlock = client.logoDataUrl
    ? `<img class="client-logo" src="${escapeHtml(client.logoDataUrl)}" alt="${escapeHtml(client.name)} logo" />`
    : '<div class="client-logo placeholder">No logo</div>';
  const locationsHtml = officeLines.length
    ? `<ul>${officeLines.map((entry) => `<li>${escapeHtml(entry)}</li>`).join("")}</ul>`
    : "<p>-</p>";
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
      </div>
      <div>
        <h5>Other Locations</h5>
        ${locationsHtml}
      </div>
    </div>
    <div class="row">
      <button class="secondary" type="button" data-client-edit-form="${escapeHtml(client.id)}">Edit this client</button>
    </div>
  `;

  clientDetailsPanel.querySelectorAll("[data-client-edit-form]").forEach((button) => {
    button.addEventListener("click", () => {
      const targetClient = clients.find((entry) => entry.id === button.dataset.clientEditForm);
      if (!targetClient) return;
      keepClientFormBlank = false;
      setClientsSectionMode("create");
      populateClientForm(targetClient);
    });
  });
}

function renderProjectsTable() {
  const canManage = can("manageCatalog");
  const scopedProjects = (selectedClientId ? projects.filter((project) => project.clientId === selectedClientId) : projects)
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name));

  if (selectedProjectId && !projects.some((project) => project.id === selectedProjectId)) selectedProjectId = "";

  const rows = scopedProjects
    .map((project) => {
      const client = clientById(project.clientId);
      const scopeSummary = projectScopeSummary(project) || "-";
      const unitTemplateCount = projectChecklistTemplate(project).length;
      return `<tr class="${project.id === selectedProjectId ? "selected-row" : ""}">
        <td>${escapeHtml(project.name)}</td>
        <td>${escapeHtml(client?.name || "-")}</td>
        <td>${escapeHtml(project.code)}</td>
        <td>${escapeHtml(project.address)}</td>
        <td>${escapeHtml(project.floorsCount || "-")}</td>
        <td>${escapeHtml(project.apartmentsCount || "-")}</td>
        <td>${escapeHtml(scopeSummary)}</td>
        <td>${unitTemplateCount}</td>
        <td>
          <div class="actions-inline">
            <button class="secondary xs-btn" data-select-project="${project.id}">Select</button>
            <button class="secondary xs-btn" data-open-project-warehouse="${project.id}">Warehouse</button>
            ${canManage ? `<button class="danger xs-btn" data-del-project="${project.id}">Delete</button>` : ""}
          </div>
        </td>
      </tr>`;
    })
    .join("");

  projectsTable.innerHTML = `<table class="data-table"><thead><tr><th>Project</th><th>Client</th><th>Code</th><th>Address</th><th>Floors</th><th>Apartments</th><th>Scope</th><th>Unit checklist</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="9">No projects for the selected client.</td></tr>'
  }</tbody></table>`;

  projectsTable.querySelectorAll("[data-select-project]").forEach((button) => {
    button.addEventListener("click", () => {
      const project = projects.find((entry) => entry.id === button.dataset.selectProject);
      if (!project) return;
      selectedProjectId = project.id;
      selectedClientId = project.clientId || selectedClientId;
      populateProjectForm(project);
      render();
      projectsSubpanel?.scrollIntoView({ behavior: "smooth", block: "center" });
    });
  });

  projectsTable.querySelectorAll("[data-open-project-warehouse]").forEach((button) => {
    button.addEventListener("click", () => {
      openProjectWarehouse(button.dataset.openProjectWarehouse || "");
    });
  });

  if (!canManage) return;

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
      if (!targetProject) return;
      const confirmed = confirmDeleteAction(`project "${targetProject.name}"`);
      if (!confirmed) return;
      pushEntityAudit("Projects", "deleted", `${targetProject.name}`, "projects");
      await trashDeleteRecord(PROJECT_STORE, targetProject, {
        label: `Project: ${targetProject.name}`,
        scope: "projects",
      });
      if (projectForm?.elements?.projectId?.value === id) resetProjectForm({ clientId: selectedClientId });
      if (selectedProjectId === id) selectedProjectId = "";
      await loadAll();
      render();
      queueAutoSync();
    });
  });
}

function renderContactsTable() {
  const canManage = can("manageCatalog");
  const scopedContacts = selectedClientId ? contacts.filter((contact) => contact.clientId === selectedClientId) : contacts;

  const rows = scopedContacts
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

  contactsTable.innerHTML = `<table class="data-table"><thead><tr><th>Name</th><th>Job Title Project</th><th>Company</th><th>Client</th><th>Project</th><th>Phone</th><th>Email</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="8">No people registered for the selected client.</td></tr>'
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
      if (!targetContact) return;
      const confirmed = confirmDeleteAction(`contact "${targetContact.name}"`);
      if (!confirmed) return;
      if (targetContact) pushEntityAudit("Contacts", "deleted", `${targetContact.name}`, "contacts");
      await trashDeleteRecord(CONTACT_STORE, targetContact, {
        label: `Contact: ${targetContact.name}`,
        scope: "contacts",
      });
      await loadAll();
      render();
      queueAutoSync();
    });
  });
}

function resetContractForm({ clientId = "", projectId = "" } = {}) {
  if (!contractForm) return;
  contractForm.reset();
  selectedContractId = "";
  if (contractForm.elements.contractId) contractForm.elements.contractId.value = "";
  if (contractClientSelect && clientId && clients.some((client) => client.id === clientId)) {
    contractClientSelect.value = clientId;
  }
  syncContractProjectSelect();
  if (contractProjectSelect && projectId && projects.some((project) => project.id === projectId)) {
    contractProjectSelect.value = projectId;
  }
  if (contractAdminTarget) contractAdminTarget.textContent = "Creating a new contract.";
  if (contractFormDeleteBtn) contractFormDeleteBtn.disabled = true;
}

function populateContractForm(contract) {
  if (!contractForm || !contract) return;
  selectedContractId = contract.id;
  if (contractForm.elements.contractId) contractForm.elements.contractId.value = contract.id;
  contractForm.elements.clientId.value = contract.clientId || "";
  syncContractProjectSelect();
  contractForm.elements.projectId.value = contract.projectId || "";
  contractForm.elements.title.value = contract.title || "";
  contractForm.elements.contractCode.value = contract.contractCode || "";
  contractForm.elements.status.value = contract.status || "draft";
  contractForm.elements.signedDate.value = normalizeDateField(contract.signedDate);
  contractForm.elements.startDate.value = normalizeDateField(contract.startDate);
  contractForm.elements.endDate.value = normalizeDateField(contract.endDate);
  contractForm.elements.amount.value = contract.amount || "";
  contractForm.elements.notes.value = contract.notes || "";
  if (contractAdminTarget) contractAdminTarget.textContent = `Editing contract: ${contract.title || "-"}`;
  if (contractFormDeleteBtn) contractFormDeleteBtn.disabled = false;
}

function contractStatusLabel(status) {
  const key = String(status || "").trim().toLowerCase();
  if (key === "active") return "Active";
  if (key === "on-hold") return "On hold";
  if (key === "completed") return "Completed";
  if (key === "cancelled") return "Cancelled";
  return "Draft";
}

function renderContractsTable() {
  if (!contractsTable) return;
  const canManage = can("manageCatalog");
  const scopedContracts = (selectedClientId ? contracts.filter((entry) => entry.clientId === selectedClientId) : contracts)
    .slice()
    .sort((a, b) => a.title.localeCompare(b.title) || a.contractCode.localeCompare(b.contractCode));

  const activeContract = selectedContractId ? contractById(selectedContractId) : null;
  if (!activeContract || (selectedClientId && activeContract.clientId !== selectedClientId)) {
    if (selectedContractId) selectedContractId = "";
    if (contractForm?.elements?.contractId?.value) resetContractForm({ clientId: selectedClientId || "" });
  }

  const rows = scopedContracts
    .map((contract) => {
      const client = clientById(contract.clientId);
      const project = projectById(contract.projectId);
      return `<tr class="${contract.id === selectedContractId ? "selected-row" : ""}">
        <td>${escapeHtml(contract.title || "-")}</td>
        <td>${escapeHtml(contract.contractCode || "-")}</td>
        <td>${escapeHtml(client?.name || "-")}</td>
        <td>${escapeHtml(project?.name || "-")}</td>
        <td>${escapeHtml(contractStatusLabel(contract.status))}</td>
        <td>${escapeHtml(contract.signedDate || "-")}</td>
        <td>${escapeHtml(contract.endDate || "-")}</td>
        <td>${escapeHtml(contract.amount || "-")}</td>
        <td>${escapeHtml(fmtDate(contract.updatedAt))}</td>
        <td>${
          canManage
            ? `<div class="actions-inline"><button class="secondary xs-btn" data-select-contract="${contract.id}">Select</button><button class="danger xs-btn" data-del-contract="${contract.id}">Delete</button></div>`
            : "-"
        }</td>
      </tr>`;
    })
    .join("");

  contractsTable.innerHTML = `<table class="data-table"><thead><tr><th>Title</th><th>Code</th><th>Client</th><th>Project</th><th>Status</th><th>Signed</th><th>End</th><th>Value (USD)</th><th>Updated</th><th>Actions</th></tr></thead><tbody>${
    rows || '<tr><td colspan="10">No contracts registered for the selected client.</td></tr>'
  }</tbody></table>`;

  if (!canManage) return;

  contractsTable.querySelectorAll("[data-select-contract]").forEach((button) => {
    button.addEventListener("click", () => {
      const contract = contractById(button.dataset.selectContract);
      if (!contract) return;
      selectedContractId = contract.id;
      selectedClientId = contract.clientId || selectedClientId;
      populateContractForm(contract);
      renderContractsTable();
    });
  });

  contractsTable.querySelectorAll("[data-del-contract]").forEach((button) => {
    button.addEventListener("click", async () => {
      const targetId = button.dataset.delContract || "";
      const targetContract = contractById(targetId);
      if (!targetContract) return;
      const confirmed = confirmDeleteAction(`contract "${targetContract.title || targetContract.contractCode || targetContract.id}"`);
      if (!confirmed) return;
      pushEntityAudit(
        "Contracts",
        "deleted",
        `${targetContract.title} | code "${auditValue(targetContract.contractCode)}"`,
        "contracts"
      );
      await trashDeleteRecord(CONTRACT_STORE, targetContract, {
        label: `Contract: ${targetContract.title || targetContract.contractCode || targetContract.id}`,
        scope: "contracts",
      });
      if (selectedContractId === targetId) resetContractForm({ clientId: selectedClientId || "" });
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
      if (!targetMaterial) return;
      const confirmed = confirmDeleteAction(`material "${targetMaterial.sku} | ${targetMaterial.description}"`);
      if (!confirmed) return;
      if (targetMaterial) pushEntityAudit("Materials", "deleted", `${targetMaterial.sku} | ${targetMaterial.description}`, "materials");
      await trashDeleteRecord(MATERIAL_STORE, targetMaterial, {
        label: `Material: ${targetMaterial.sku} | ${targetMaterial.description}`,
        scope: "materials",
      });
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
  setFormEnabled(contractForm, canManage);
  setFormEnabled(materialForm, canManage);

  if (selectedClientId && !clients.some((client) => client.id === selectedClientId)) selectedClientId = "";
  if (selectedProjectId && !projects.some((project) => project.id === selectedProjectId)) selectedProjectId = "";
  if (selectedContractId && !contracts.some((contract) => contract.id === selectedContractId)) selectedContractId = "";

  renderMasterSelects();
  if (projectForm?.elements?.projectId?.value) {
    const editingProject = projects.find((entry) => entry.id === projectForm.elements.projectId.value) || null;
    if (!editingProject) resetProjectForm({ clientId: selectedClientId });
  } else if (projectClientSelect && selectedClientId && clients.some((client) => client.id === selectedClientId)) {
    projectClientSelect.value = selectedClientId;
  }
  renderProjectDraftPills();
  renderMaterialsTable();
  if (!canManage) return;
  renderClientsHubList();
  renderClientsTable();
  renderProjectsTable();
  renderContactsTable();
  renderContractsTable();
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
  renderManufactureCatalogSummaries();

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
      const confirmed = confirmDeleteAction(`container "${container.containerCode || container.id}"`);
      if (!confirmed) return;
      pushContainerAudit(container, `Container deleted: ${container.containerCode}`);
      await trashDeleteRecord(CONTAINER_STORE, container, {
        label: `Container: ${container.containerCode || container.id}`,
        scope: container.projectId || "containers",
      });
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
          flowStage: "warehouse",
          flowStatus: "generated",
          flowUpdatedAt: now,
          flowEvents: [
            {
              id: uid(),
              at: now,
              byUserId: currentUser?.id || "",
              byName: currentUser?.name || "System",
              stage: "warehouse",
              status: "generated",
              issueType: "",
              note: "QR generated from manifest.",
              unitId: "",
              unitLabel: "",
              source: "manifest",
            },
          ],
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
          if (!removed) return;
          const confirmed = confirmDeleteAction(`manifest item "${removed.description || removed.code || removed.id}"`);
          if (!confirmed) return;
          const removedQrItems = container.qrItems.filter(
            (item) => item.manifestItemId === btn.dataset.manifestDel && (item.status === "in-warehouse" || item.status === "dispatched")
          );
          await saveTrashRecord({
            storeName: CONTAINER_STORE,
            recordId: `${container.id}:${removed.id}`,
            label: `Container manifest: ${removed.description || removed.code || removed.id}`,
            scope: container.containerCode || "containers",
            payload: removed,
            restoreType: "container-manifest-item",
            context: { containerId: container.id, qrItems: removedQrItems },
          });
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
    const qrScanBtn = qrDispatchForm.querySelector(".qd-scan-btn");
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
    qrScanBtn?.addEventListener("click", async () => {
      if (!qrDispatchEditable) return;
      await scanQrIntoInput(qrScanInput, "Scan QR for dispatch");
    });

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
        alert("QR not found in this container.");
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
      appendQrFlowEvent(container, qrItem, {
        stage: "delivery",
        status: "in-progress",
        note: `Dispatched from warehouse to unit ${targetUnit.unitCode}`,
        source: "dispatch",
        unit: targetUnit,
        at: now,
      });
      applyStageToUnitFromQr(targetUnit, "warehouse", "completed", `QR dispatched: ${qrItem.qrCode}`, now);
      targetUnit.updatedAt = now;
      qrDispatchForm.reset();
      await saveUnitAndContainer(targetUnit, container, `QR dispatch recorded: ${qrItem.qrCode} -> ${targetUnit.unitCode}`);
    });

    applySectorVisibility(node, currentView === "manufacture" ? "fabrica" : currentProjectSector);
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
  const qrCount = checkItems.reduce((sum, item) => sum + unitChecklistQrCodesText(item).length, 0);
  return `Items: ${total} | QR labels: ${qrCount} | Missing: ${missing} | Damaged: ${damaged} | Adjustment: ${adjustment}`;
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
          <th>QR codes</th>
          <th>Expected</th>
          <th>Checked</th>
          <th>Status</th>
          <th>Notes</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        ${unit.checkItems
          .map((item) => {
            const qrCodes = unitChecklistQrCodesText(item);
            const firstQr = qrCodes[0] || "";
            const extraQrCount = qrCodes.length > 1 ? qrCodes.length - 1 : 0;
            const qrPreview = firstQr
              ? `<img class="qr-thumb" src="${qrImageUrl(firstQr, 120)}" alt="QR for ${escapeHtml(item.description || item.code || "item")}" loading="lazy" />`
              : `<span class="hint">No QR</span>`;
            return `
            <tr>
              <td>${escapeHtml(item.code)}</td>
              <td>${escapeHtml(item.description)}</td>
              <td>
                <div class="qr-item-cell">
                  ${qrPreview}
                  ${qrCodes.length ? `<small>${qrCodes.length} label(s)${extraQrCount ? ` (+${extraQrCount} extra)` : ""}</small>` : ""}
                  <div class="actions-inline">
                    <button type="button" class="secondary xs-btn" data-item-qr-copy="${item.id}">Copy</button>
                    <button type="button" class="secondary xs-btn" data-item-qr-print="${item.id}">Print</button>
                    <button type="button" class="secondary xs-btn" data-item-qr-png="${item.id}">PNG</button>
                  </div>
                </div>
              </td>
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
          })
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
    button.addEventListener("click", async () => {
      const removed = unit.checkItems.find((entry) => entry.id === button.dataset.itemRemove);
      if (!removed) return;
      const confirmed = confirmDeleteAction(`checklist item "${removed.description || removed.code || removed.id}"`);
      if (!confirmed) return;
      await saveTrashRecord({
        storeName: UNIT_STORE,
        recordId: `${unit.id}:${removed.id}`,
        label: `Checklist item: ${removed.description || removed.code || removed.id}`,
        scope: unit.projectName || unit.unitCode || "units",
        payload: removed,
        restoreType: "unit-check-item",
        context: { unitId: unit.id },
      });
      unit.checkItems = unit.checkItems.filter((entry) => entry.id !== button.dataset.itemRemove);
      await saveUnitWithAudit(unit, `Item removed from checklist: ${removed.description || removed.id}`);
    });
  });

  wrapper.querySelectorAll("[data-item-qr-print]").forEach((button) => {
    button.addEventListener("click", () => {
      const item = unit.checkItems.find((entry) => entry.id === button.dataset.itemQrPrint);
      if (!item) return;
      openUnitChecklistQrLabels(unit, item, { autoPrint: true });
    });
  });

  wrapper.querySelectorAll("[data-item-qr-png]").forEach((button) => {
    button.addEventListener("click", () => {
      const item = unit.checkItems.find((entry) => entry.id === button.dataset.itemQrPng);
      if (!item) return;
      downloadUnitChecklistQrPng(item);
    });
  });

  wrapper.querySelectorAll("[data-item-qr-copy]").forEach((button) => {
    button.addEventListener("click", async () => {
      const item = unit.checkItems.find((entry) => entry.id === button.dataset.itemQrCopy);
      if (!item) return;
      const qrCodes = unitChecklistQrCodesText(item);
      if (!qrCodes.length) return;
      const text = qrCodes.join("\n");
      try {
        await navigator.clipboard.writeText(text);
      } catch {
        window.prompt("Copy QR code(s):", text);
      }
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
        button.addEventListener("click", async () => {
          const row = unit.dispatchTasks.find((entry) => entry.id === button.dataset.dispatchDel);
          if (!row) return;
          const confirmed = confirmDeleteAction(`dispatch task "${row.item || row.id}"`);
          if (!confirmed) return;
          await saveTrashRecord({
            storeName: UNIT_STORE,
            recordId: `${unit.id}:${row.id}`,
            label: `Dispatch task: ${row.item || row.id}`,
            scope: unit.projectName || unit.unitCode || "units",
            payload: row,
            restoreType: "unit-dispatch-task",
            context: { unitId: unit.id },
          });
          unit.dispatchTasks = unit.dispatchTasks.filter((entry) => entry.id !== button.dataset.dispatchDel);
          await saveUnitWithAudit(unit, `Dispatch task removed: ${row?.item || button.dataset.dispatchDel}`);
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
    const siteScanBtn = siteForm.querySelector(".sr-scan-btn");
    const canSiteReceive = can("siteReceive");
    if (!canSiteReceive) siteForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));
    siteScanBtn?.addEventListener("click", async () => {
      if (!canSiteReceive) return;
      await scanQrIntoInput(siteForm.querySelector(".sr-qr"), "Scan QR for site receive");
    });
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
          alert("QR not found for this project.");
          return;
        }
        const qrItem = qrContainer.qrItems.find((entry) => entry.qrCode === qrCode);
        if (!qrItem) {
          alert("QR not found.");
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
        appendQrFlowEvent(qrContainer, qrItem, {
          stage: "delivery",
          status: issue === "ok" ? "completed" : "issue",
          issueType: issue === "ok" ? "" : issue,
          note,
          source: "site-receive",
          unit,
          at: now,
        });
        applyStageToUnitFromQr(unit, "delivery", issue === "ok" ? "completed" : "issue", note, now);
        unit.updatedAt = now;

        if (issue !== "ok") {
          const issueLine = `${effectiveItem} (${effectiveQty}) -> ${issue}${note ? ` | ${note}` : ""}`;
          unit.issuesText = unit.issuesText ? `${unit.issuesText}\n${issueLine}` : issueLine;
          pushQrIssueQueue(qrContainer, qrItem, unit, issue, note, { source: "site-receive", at: now });
        }

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
        button.addEventListener("click", async () => {
          const row = unit.siteReceipts.find((entry) => entry.id === button.dataset.receiveDel);
          if (!row) return;
          const confirmed = confirmDeleteAction(`site receipt "${row.item || row.id}"`);
          if (!confirmed) return;
          await saveTrashRecord({
            storeName: UNIT_STORE,
            recordId: `${unit.id}:${row.id}`,
            label: `Site receipt: ${row.item || row.id}`,
            scope: unit.projectName || unit.unitCode || "units",
            payload: row,
            restoreType: "unit-site-receipt",
            context: { unitId: unit.id },
          });
          unit.siteReceipts = unit.siteReceipts.filter((entry) => entry.id !== button.dataset.receiveDel);
          await saveUnitWithAudit(unit, `Receipt removed: ${row?.item || button.dataset.receiveDel}`);
        });
      });
    }

    const auditTable = node.querySelector(".audit-table");
    if (auditTable) renderAuditTable(auditTable, unit);

    const checkForm = node.querySelector(".check-item-form");
    const addDefaultsBtn = node.querySelector(".add-default-items-btn");
    if (!can("checklist")) checkForm.querySelectorAll("input,select,button").forEach((el) => (el.disabled = true));
    if (!can("checklist") && addDefaultsBtn) addDefaultsBtn.disabled = true;

    addDefaultsBtn?.addEventListener("click", () => {
      if (!can("checklist")) return;
      const added = addDefaultChecklistItemsToUnit(unit);
      if (!added) {
        alert("All default materials are already in this unit.");
        return;
      }
      saveUnitWithAudit(unit, `${added} default material item(s) with QR codes added.`);
    });

    checkForm.addEventListener("submit", (event) => {
      event.preventDefault();
      if (!can("checklist")) return;

      const code = checkForm.querySelector(".ci-code").value.trim();
      const description = checkForm.querySelector(".ci-desc").value.trim();
      const expectedQty = Number(checkForm.querySelector(".ci-exp").value || 1);
      const checkedQty = Number(checkForm.querySelector(".ci-got").value || 0);
      const status = checkForm.querySelector(".ci-status").value;

      if (!description) return;

      const createdItem = createCheckItem(
        {
          code,
          description,
          expectedQty: Math.max(1, Math.trunc(expectedQty || 1)),
          checkedQty,
          status,
        },
        unit,
        unit.checkItems.length
      );
      unit.checkItems.push(createdItem);

      checkForm.reset();
      saveUnitWithAudit(
        unit,
        `Checklist item added: ${description} (${createdItem.expectedQty} piece${createdItem.expectedQty === 1 ? "" : "s"}) with ${
          createdItem.qrItems.length
        } QR label(s)`
      );
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
        const confirmed = confirmDeleteAction(`photo "${photo.name || photo.id}"`);
        if (!confirmed) return;
        pushAppAudit(
          `[Unit ${unit.unitCode || unit.id}] Photo removed: ${photo.name || photo.id}`,
          "unit-change",
          unit.projectName || "-"
        );
        await trashDeleteRecord(PHOTO_STORE, photo, {
          label: `Photo: ${photo.name || photo.id}`,
          scope: unit.projectName || unit.unitCode || "photos",
        });
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
      const confirmed = confirmDeleteAction(`unit "${unit.unitCode || unit.id}"`);
      if (!confirmed) return;
      pushAppAudit(`[Unit ${unit.unitCode || unit.id}] Unit deleted`, "unit-change", unit.projectName || "-");
      const toDelete = photosByUnit.get(unit.id) || [];
      await trashDeleteRecord(UNIT_STORE, unit, {
        label: `Unit: ${unit.unitCode || unit.id}`,
        scope: unit.projectName || "units",
        relatedRecords: toDelete.map((photo) => ({ storeName: PHOTO_STORE, payload: photo })),
      });
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

function isPendingUserAssignment(user) {
  return userAccessProfile(user) === "visitor";
}

function setUsersSubView(view = "directory") {
  const normalized = view === "registration" ? "registration" : "directory";
  usersSubView = normalized;
  usersDirectoryView?.classList.toggle("hidden", normalized !== "directory");
  usersRegistrationView?.classList.toggle("hidden", normalized !== "registration");
  usersRegistrationTabBtn?.classList.toggle("active", normalized === "registration");
  usersDirectoryTabBtn?.classList.toggle("active", normalized === "directory");
}

function openUsersRegistrationClean() {
  resetAdminUserForm();
  setUserAdminFormOpen(false);
  setUsersSubView("registration");
  syncSelectedUserRow();
}

function resetAdminUserForm() {
  if (!userForm) return;
  adminEditingUserId = "";
  userForm.reset();
  userForm.accessProfile.value = "visitor";
  userForm.employmentType.value = "tag";
  if (userAdminGenderSelect) userAdminGenderSelect.value = "unspecified";
  toggleSubcontractorExtras("admin", "tag");
  if (userAdminTarget) userAdminTarget.textContent = "Creating a new user.";
  setAvatarPreview(userAdminPhotoPreview, defaultAvatarForGender(userAdminGenderSelect?.value || "unspecified"));
  if (userAdminCoiFileStatus) userAdminCoiFileStatus.textContent = "No COI file attached.";
  if (openUserAdminCoiFileBtn) openUserAdminCoiFileBtn.classList.add("hidden");
  if (userFormDeleteBtn) userFormDeleteBtn.disabled = true;
  renderUserIdCard(null);
}

function setUserAdminFormOpen(open) {
  userAdminFormOpen = Boolean(open);
  if (!userFormPanel) return;
  userFormPanel.classList.remove("hidden");
  userFormPanel.classList.toggle("users-form-editing", userAdminFormOpen);
}

function populateAdminUserForm(user) {
  if (!userForm || !user) return;
  adminEditingUserId = user.id;
  setUserAdminFormOpen(true);
  userForm.firstName.value = user.firstName || "";
  userForm.lastName.value = user.lastName || "";
  userForm.birthDate.value = normalizeDateField(user.birthDate);
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
  userForm.accessProfile.value = userAccessProfile(user);
  if (userAdminGenderSelect) userAdminGenderSelect.value = normalizeGender(user.gender);
  const photoInput = userForm.querySelector('input[name="photo"]');
  if (photoInput) photoInput.value = "";
  const contractorFileInput = userForm.querySelector('input[name="contractorCoiFile"]');
  if (contractorFileInput) contractorFileInput.value = "";
  if (userAdminTarget) userAdminTarget.textContent = `Editing user: ${user.name || user.username}`;
  setAvatarPreview(userAdminPhotoPreview, userAvatarSrc(user));
  if (userAdminCoiFileStatus) {
    userAdminCoiFileStatus.textContent = user.contractorCoiFile?.name
      ? `Current file: ${user.contractorCoiFile.name}`
      : "No COI file attached.";
  }
  if (openUserAdminCoiFileBtn) openUserAdminCoiFileBtn.classList.toggle("hidden", !user.contractorCoiFile?.dataUrl);
  if (userFormDeleteBtn) userFormDeleteBtn.disabled = user.id === currentUser?.id;
  renderUserIdCard(user);
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
    setUsersSubView("directory");
    setUserAdminFormOpen(false);
    usersPanel.classList.add("hidden");
    return;
  }

  usersPanel.classList.remove("hidden");
  setUsersSubView(usersSubView);

  const visibleUsers = users.filter((user) => {
    if (!canAccessEmployeeRecord(user)) return false;
    if (isDeveloper()) return usersFilterMatches(user);
    if (userAccessProfile(user) === "developer" || isPrimaryDeveloperUser(user)) return false;
    return usersFilterMatches(user);
  });

  const userAccessProfileSelect = userForm?.querySelector('select[name="accessProfile"]');
  if (userAccessProfileSelect) {
    const developerOption = userAccessProfileSelect.querySelector('option[value="developer"]');
    if (developerOption) developerOption.hidden = !isDeveloper();
    if (!isDeveloper() && userAccessProfileSelect.value === "developer") userAccessProfileSelect.value = "admin";
  }

  if (!adminEditingUserId && userForm && !userForm.dataset.ready) {
    resetAdminUserForm();
    userForm.dataset.ready = "1";
    setUserAdminFormOpen(false);
  }

  if (usersNameRail) {
    const nameRows = visibleUsers
      .map((user) => {
        const profile = userAccessProfile(user);
        const pendingAssignment = isPendingUserAssignment(user);
        return `<button class="users-name-btn${pendingAssignment ? " pending-assignment" : ""}" type="button" data-user-rail="${user.id}">
          ${escapeHtml(user.name || user.username || "-")}
          <small>${escapeHtml(roleLabel(profile))} • ${escapeHtml(user.companyName || "No company")}</small>
          ${pendingAssignment ? '<span class="pending-assignment-chip">Pending assignment</span>' : ""}
        </button>`;
      })
      .join("");
    usersNameRail.innerHTML = nameRows || '<p class="hint">No users available.</p>';
    usersNameRail.querySelectorAll("[data-user-rail]").forEach((button) => {
      button.addEventListener("click", () => {
        selectAdminUser(button.dataset.userRail || "");
      });
    });
  }

  const rows = visibleUsers
    .map((user) => {
      const avatarSrc = userAvatarSrc(user);
      const profile = userAccessProfile(user);
      const pendingAssignment = isPendingUserAssignment(user);
      return `<tr data-user-row="${user.id}" class="${pendingAssignment ? "user-pending-assignment" : ""}">
      <td>${avatarSrc ? `<img class="avatar-xs" src="${escapeHtml(avatarSrc)}" alt="${escapeHtml(user.name || user.username || "User")}" />` : ""}</td>
      <td>${escapeHtml(user.name)}</td>
      <td>${escapeHtml(user.companyName || "-")}</td>
      <td>${escapeHtml(user.jobTitle || "-")}</td>
      <td>${escapeHtml(user.employmentType || "-")}</td>
      <td>${escapeHtml(user.email || "-")}</td>
      <td>${escapeHtml(user.username)}</td>
      <td>${
        pendingAssignment
          ? '<span class="pending-assignment-pill">Pending assignment (Visitor)</span>'
          : escapeHtml(roleLabel(profile))
      }</td>
      <td>${fmtDate(user.updatedAt)}</td>
    </tr>`;
    })
    .join("");

  usersTable.innerHTML = `<table class="data-table"><thead><tr><th>Photo</th><th>Name</th><th>Company</th><th>Job Title</th><th>Type</th><th>Email</th><th>User</th><th>Access Profile</th><th>Updated</th></tr></thead><tbody>${
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
    else {
      resetAdminUserForm();
      setUserAdminFormOpen(false);
    }
  } else {
    if (!userAdminFormOpen) renderUserIdCard(null);
  }
  syncSelectedUserRow();
}

function populateUserEditForm(user) {
  if (!userEditForm || !user) return;
  userEditForm.firstName.value = user.firstName || "";
  userEditForm.lastName.value = user.lastName || "";
  userEditForm.birthDate.value = normalizeDateField(user.birthDate);
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
  if (userEditGenderSelect) userEditGenderSelect.value = normalizeGender(user.gender);
  const photoInput = userEditForm.querySelector('input[name="photo"]');
  if (photoInput) photoInput.value = "";
  const contractorFileInput = userEditForm.querySelector('input[name="contractorCoiFile"]');
  if (contractorFileInput) contractorFileInput.value = "";
  if (userEditTarget) userEditTarget.textContent = `Editing: ${user.name || user.username}`;
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
  setAvatarPreview(userEditPhotoPreview, userAvatarSrc(user));
}

function openUserEdit(userId, returnView = "users") {
  if (!currentUser) return;
  if (userAccessProfile(currentUser) === "warehouse" && userId !== currentUser.id) return;
  if (!isDeveloper() && !isAdmin() && userId !== currentUser.id) return;
  const user = users.find((entry) => entry.id === userId);
  if (!user) return;
  if (!canAccessEmployeeRecord(user)) return;
  if (!isDeveloper() && (userAccessProfile(user) === "developer" || isPrimaryDeveloperUser(user))) {
    alert("Only developer can edit developer accounts.");
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
  const canManufacture = canOpenView("manufacture");
  const canReports = can("report");
  const isDev = isDeveloper();
  const visibleSectors = new Set(allowedProjectSectors());

  const toggleNav = (action, visible) => {
    document.querySelectorAll(`[data-nav-action="${action}"]`).forEach((el) => {
      el.classList.toggle("hidden", !visible);
    });
  };
  const toggleNavGroup = (group, visible) => {
    document.querySelectorAll(`[data-nav-group="${group}"]`).forEach((el) => {
      el.classList.toggle("hidden", !visible);
    });
  };

  toggleNavGroup("users", canManageUsers);
  toggleNav("developer", isDev);
  toggleNav("users", canManageUsers);
  toggleNav("usersRegistration", canManageUsers);
  toggleNav("usersPeople", canManageUsers);
  toggleNav("subcontractors", canManageUsers);
  toggleNavGroup("partners", canManageUsers);
  toggleNav("partners", canManageUsers);
  toggleNav("partnersManufacturers", canManageUsers);
  toggleNav("partnersConstructionMaterials", canManageUsers);
  toggleNav("partnersImporters", canManageUsers);
  toggleNav("partnersCarriers", canManageUsers);
  toggleNav("partnersTruckDeliveries", canManageUsers);
  toggleNavGroup("manufacture", canManufacture);
  toggleNav("manufacture", canManufacture);
  toggleNav("manufactureSchedule", canManufacture);
  toggleNav("manufactureCatalog", canManufacture);
  toggleNav("manufactureSolicitation", canManufacture);
  toggleNavGroup("clients", isCatalogAdmin);
  toggleNav("clients", isCatalogAdmin);
  toggleNav("clientsNew", isCatalogAdmin);
  toggleNav("clientsProjects", isCatalogAdmin);
  toggleNav("clientsContracts", isCatalogAdmin);
  toggleNav("projects", false);
  const canWarehouseGroup = canProjects && (visibleSectors.has("warehouse") || visibleSectors.has("delivery"));
  toggleNavGroup("warehouse", canWarehouseGroup);
  toggleNav("warehouse", canWarehouseGroup);
  toggleNav("warehouseDeliverySchedule", canProjects && visibleSectors.has("delivery"));
  toggleNav("warehouseDeliveryInventory", canProjects && visibleSectors.has("delivery"));
  toggleNav("warehouseQrScan", canProjects && visibleSectors.has("delivery"));
  toggleNav("warehouseOperations", canWarehouseGroup);
  toggleNav("delivery", false);
  toggleNav("distribution", canProjects && visibleSectors.has("distribuicao"));
  toggleNav("installation", canProjects && visibleSectors.has("instalacao"));
  toggleNav("changeOrders", canProjects && visibleSectors.has("punchlist"));
  toggleNav("projectReports", canProjects && canReports);
  toggleNav("ocrImporter", can("manageCatalog") || can("manageContainerManifest") || isAdmin() || isDeveloper());
  renderCoiReminderPanel();
}

function renderRoleStrip() {
  const accessProfile = userAccessProfile(currentUser);
  roleLine.textContent = `User: ${currentUser.name} | Access Profile: ${roleLabel(accessProfile)}`;
  permissionLine.textContent = rolePermissionsSummary(accessProfile);
}

function backupCountsSnapshot() {
  return {
    units: units.length,
    users: users.length,
    clients: clients.length,
    projects: projects.length,
    contacts: contacts.length,
    contracts: contracts.length,
    containers: containers.length,
    materials: materials.length,
    delivery: deliverySkuItems.length,
    trash: trashRecords.length,
  };
}

function backupCountsSummary(counts) {
  return `units:${counts.units}, users:${counts.users}, clients:${counts.clients}, projects:${counts.projects}, contacts:${counts.contacts}, contracts:${counts.contracts}, containers:${counts.containers}, materials:${counts.materials}, delivery:${counts.delivery}, trash:${counts.trash}`;
}

function lastLocalBackupAt() {
  try {
    return String(localStorage.getItem(LOCAL_BACKUP_META_KEY) || "").trim();
  } catch {
    return "";
  }
}

function localBackupStatusLabel() {
  const lastBackupAt = lastLocalBackupAt();
  if (!lastBackupAt) return "No local backup downloaded on this device yet.";
  return `Last local backup downloaded: ${fmtDate(lastBackupAt)}.`;
}

function renderDeveloperAuditPanel() {
  if (!developerAuditPanel) return;
  if (!isDeveloper()) {
    developerAuditPanel.classList.add("hidden");
    return;
  }

  const backupStatus = localBackupStatusLabel();

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

  const deletedRows = ensureArray(trashRecords)
    .slice()
    .sort((a, b) => new Date(b.deletedAt || 0).getTime() - new Date(a.deletedAt || 0).getTime())
    .slice(0, 600)
    .map(
      (entry) => `<tr>
        <td>${escapeHtml(fmtDate(entry.deletedAt))}</td>
        <td>${escapeHtml(entry.deletedByName || "-")}</td>
        <td>${escapeHtml(trashStoreLabel(entry.storeName))}</td>
        <td>${escapeHtml(entry.scope || "-")}</td>
        <td>${escapeHtml(entry.label || entry.recordId || "-")}</td>
        <td>${escapeHtml(formatTrashTimeLeft(entry))}</td>
        <td>
          <div class="actions-inline">
            <button class="secondary xs-btn" type="button" data-trash-restore="${escapeHtml(entry.id)}">Restore</button>
            <button class="danger xs-btn" type="button" data-trash-delete="${escapeHtml(entry.id)}">Delete forever</button>
          </div>
        </td>
      </tr>`
    )
    .join("");

  developerAuditPanel.classList.remove("hidden");
  developerAuditPanel.innerHTML = `
    <h3>Audit Log (Developer)</h3>
    <p class="hint">Tracks exact field changes and navigation history by user.</p>
    <div class="developer-tools">
      <button class="secondary developer-tool-btn" type="button" data-developer-backup-download>
        <span class="developer-tool-code">BK</span>
        <span class="developer-tool-text">
          <strong>Local Full Backup</strong>
          <small>Download complete JSON copy to this device</small>
        </span>
      </button>
      <label class="file-label secondary developer-tool-btn">
        <span class="developer-tool-code">RS</span>
        <span class="developer-tool-text">
          <strong>Restore Backup File</strong>
          <small>Select a JSON backup from your computer</small>
        </span>
        <input data-developer-backup-import type="file" accept="application/json" />
      </label>
    </div>
    <p class="hint developer-backup-status" data-developer-backup-status>${escapeHtml(backupStatus)}</p>
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
    <details class="trash-recovery-box">
      <summary><strong>Deleted Records Recovery (${RESTORE_WINDOW_LABEL})</strong></summary>
      <p class="hint">Hidden recycle area for deletes. Restore window is ${RESTORE_WINDOW_LABEL}.</p>
      <div class="row">
        <button class="secondary xs-btn" type="button" data-trash-purge-expired>Purge expired now</button>
      </div>
      <div class="table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th>Deleted at</th>
              <th>Deleted by</th>
              <th>Store</th>
              <th>Scope</th>
              <th>Item</th>
              <th>Time left</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>${deletedRows || '<tr><td colspan="7">No deleted records available for restore.</td></tr>'}</tbody>
        </table>
      </div>
    </details>
  `;

  developerAuditPanel.querySelectorAll("[data-trash-restore]").forEach((button) => {
    button.addEventListener("click", async () => {
      const target = trashRecords.find((entry) => entry.id === button.dataset.trashRestore);
      if (!target) return;
      const confirmed = window.confirm(`Restore "${target.label || target.recordId || target.id}" now?`);
      if (!confirmed) return;
      await restoreTrashRecord(target.id);
    });
  });

  const backupDownloadBtn = developerAuditPanel.querySelector("[data-developer-backup-download]");
  backupDownloadBtn?.addEventListener("click", () => {
    triggerLocalBackupDownload("developer-tools");
    renderDeveloperAuditPanel();
  });

  const backupImportInput = developerAuditPanel.querySelector("[data-developer-backup-import]");
  backupImportInput?.addEventListener("change", async () => {
    const file = backupImportInput.files?.[0];
    if (!file) return;
    await importBackupFile(file);
    backupImportInput.value = "";
    renderDeveloperAuditPanel();
  });

  developerAuditPanel.querySelectorAll("[data-trash-delete]").forEach((button) => {
    button.addEventListener("click", async () => {
      const target = trashRecords.find((entry) => entry.id === button.dataset.trashDelete);
      if (!target) return;
      const confirmed = window.confirm(`Delete "${target.label || target.recordId || target.id}" permanently?`);
      if (!confirmed) return;
      await permanentlyDeleteTrashRecord(target.id);
      await loadAll();
      render();
      queueAutoSync();
    });
  });

  const purgeBtn = developerAuditPanel.querySelector("[data-trash-purge-expired]");
  purgeBtn?.addEventListener("click", async () => {
    const purgedCount = await purgeExpiredTrashRecords();
    if (purgedCount) {
      pushEntityAudit("Deleted Records", "purged-expired", `${purgedCount} expired record(s)`, "trash");
    }
    await loadAll();
    render();
    if (purgedCount) queueAutoSync();
  });
}

function renderSyncPanel() {
  const allowed = can("sync");
  syncPanel.classList.toggle("hidden", !allowed);
  if (!allowed) return;

  syncConfigForm.supabaseUrl.value = syncConfig.supabaseUrl;
  syncConfigForm.supabaseAnonKey.value = syncConfig.supabaseAnonKey;
  syncConfigForm.tenant.value = syncConfig.tenant;
  syncConfigForm.autoSync.value = String(syncConfig.autoSync);

  const presetLocked = hasPresetSyncConfig();
  syncConfigForm.querySelectorAll('input[name="supabaseUrl"], input[name="supabaseAnonKey"], input[name="tenant"]').forEach((field) => {
    field.disabled = presetLocked;
  });
  const syncSaveBtn = syncConfigForm.querySelector('button[type="submit"]');
  if (syncSaveBtn) syncSaveBtn.disabled = presetLocked;

  const statusText = presetLocked
    ? "Cloud sync is centrally configured for all devices."
    : syncConfig.lastSyncAt
      ? `Last sync: ${fmtDate(syncConfig.lastSyncAt)}`
      : t("Sem sincronizacao");
  updateSyncStatus(statusText, false);
  renderDeveloperAuditPanel();
}

function renderQrLookupPanel() {
  if (!qrLookupPanel || !qrLookupResult) return;
  qrLookupPanel.classList.toggle("hidden", !currentUser);
  if (!currentUser) return;

  populateQrFlowUnitSelect();
  syncQrFlowStageOptions();
  updateQrFlowStageCustomVisibility();
  updateQrFlowIssueVisibility();
  renderQrFlowAccessHint();
  const canEditWorkflow = canUpdateQrWorkflow();
  if (qrFlowForm) {
    qrFlowForm.querySelectorAll("input,select,button").forEach((field) => {
      field.disabled = !canEditWorkflow;
    });
    if (qrFlowStageSelect) {
      qrFlowStageSelect.disabled = !canEditWorkflow || Boolean(qrLockedStageForCurrentUser());
    }
    if (canEditWorkflow) {
      syncQrFlowStageOptions();
      updateQrFlowStageCustomVisibility();
      updateQrFlowIssueVisibility();
    }
  }
  if (!canEditWorkflow) {
    setQrFlowResult("You have read-only access for QR workflow updates.");
  }

  const lookupCode = String(lastQrLookupCode || "").trim();
  if (!lookupCode) {
    qrLookupResult.innerHTML = `<p class="hint">Search a QR code to view complete status, routing, and history.</p>`;
    return;
  }

  const snapshot = buildQrHistory(lookupCode);
  if (!snapshot) {
    qrLookupResult.innerHTML = `<p class="hint">QR not found: ${escapeHtml(lookupCode)}</p>`;
    return;
  }

  const { container, item, client, project, unit, events } = snapshot;
  if (qrFlowCodeInput && !qrFlowCodeInput.value.trim()) qrFlowCodeInput.value = item.qrCode;
  if (qrFlowUnitSelect) {
    const flowUnit = item.unitId || unit?.id || "";
    if (flowUnit && Array.from(qrFlowUnitSelect.options).some((entry) => entry.value === flowUnit)) {
      qrFlowUnitSelect.value = flowUnit;
    }
  }
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
      <div class="kpi-box"><span>Stage</span><strong>${escapeHtml(workflowStageLabel(item.flowStage || "warehouse"))}</strong></div>
      <div class="kpi-box"><span>Status</span><strong>${escapeHtml(workflowStatusLabel(item.flowStatus || item.status))}</strong></div>
      <div class="kpi-box"><span>Container</span><strong>${escapeHtml(container.containerCode || "-")}</strong></div>
      <div class="kpi-box"><span>Client</span><strong>${escapeHtml(client?.name || "-")}</strong></div>
      <div class="kpi-box"><span>Project</span><strong>${escapeHtml(project?.name || "-")}</strong></div>
      <div class="kpi-box"><span>Unit</span><strong>${escapeHtml(item.unitLabel || unit?.unitCode || "-")}</strong></div>
      <div class="kpi-box"><span>Kitchen Type</span><strong>${escapeHtml(item.kitchenType || "-")}</strong></div>
    </div>
    <div class="row">
      <button class="secondary" type="button" data-lookup-print="${escapeHtml(item.id)}">Print label (PDF/Print)</button>
      <button class="secondary" type="button" data-lookup-png="${escapeHtml(item.id)}">Download label PNG</button>
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
  renderDeliveryInventoryPanel();
  renderOcrImporterPanel();
  renderUsers();
  renderUserEditPanel();
  renderAccessControl();
  renderUnits();
  applyViewMode();
  applyLanguageToUi();
}

function toUnit(formData) {
  const unitId = uid();
  const projectId = formData.get("projectId")?.toString();
  const project = projectById(projectId);
  const client = project ? clientById(project.clientId) : null;
  const now = new Date().toISOString();
  const projectScope = projectScopeSummary(project);
  const manualScope = formData.get("scopeWork")?.toString().trim() || "";
  const scopeWork = manualScope || projectScope;
  const unitCode = formData.get("unitCode")?.toString().trim();
  const baseChecklist = uniqueTextList([...projectChecklistTemplate(project), ...UNIT_DEFAULT_QR_MATERIALS]);
  const unitDraft = {
    id: unitId,
    clientId: client?.id || "",
    projectId: project?.id || "",
    clientName: client?.name || formData.get("clientName")?.toString().trim(),
    projectName: project?.name || formData.get("projectName")?.toString().trim(),
    unitCode,
  };
  const checkItems = baseChecklist.map((description, index) =>
    createCheckItem({ code: "", description, expectedQty: 1, checkedQty: 0, status: "ok", createdAt: now, updatedAt: now }, unitDraft, index)
  );

  return normalizeUnit({
    ...unitDraft,
    jobSite: formData.get("jobSite")?.toString().trim(),
    building: formData.get("building")?.toString().trim(),
    floor: formData.get("floor")?.toString().trim(),
    category: formData.get("category")?.toString().trim(),
    unitType: formData.get("unitType")?.toString().trim(),
    kitchenModel: formData.get("kitchenModel")?.toString().trim(),
    shopdrawingRef: formData.get("shopdrawingRef")?.toString().trim(),
    scopeWork,
    deliveryQuality: "pending",
    installationStatus: "not-started",
    deliveryNotes: "",
    issuesText: "",
    checkItems,
    dispatchTasks: [],
    siteReceipts: [],
    dispatchSignature: null,
    auditLog: [
      {
        id: uid(),
        at: now,
        byUserId: currentUser?.id || "",
        byName: currentUser?.name || "System",
        message: "Unit created",
      },
    ],
    stages: Object.fromEntries(STAGES.map((stage) => [stage.key, { done: false, at: null }])),
    createdAt: now,
    updatedAt: now,
  });
}

function syncEndpoint() {
  const url = syncConfig.supabaseUrl.trim().replace(/\/$/, "");
  if (!url || !syncConfig.supabaseAnonKey || !syncConfig.tenant) return null;
  return `${url}/rest/v1/app_records`;
}

function normalizeSyncKinds(kinds) {
  if (!Array.isArray(kinds) || kinds.length === 0) return null;
  const validKinds = new Set(["unit", "photo", "user", "client", "project", "contact", "contract", "container", "material", "deliverySku", "trash"]);
  const normalized = kinds
    .map((kind) => String(kind || "").trim())
    .filter((kind) => validKinds.has(kind));
  return normalized.length ? normalized : null;
}

function toCloudRecords({ kinds = null } = {}) {
  const tenant = syncConfig.tenant.trim();
  const kindFilter = normalizeSyncKinds(kinds);
  const sanitizeUserForSync = (user) => {
    const clean = { ...user };
    delete clean.legacyPassword;
    delete clean.password;
    return clean;
  };

  const mapRecords = (arr, kind, updatedField = "updatedAt") =>
    (kindFilter && !kindFilter.includes(kind) ? [] : arr).map((item) => ({
      tenant,
      kind,
      id: item.id,
      payload: item,
      updated_at: item[updatedField] || item.createdAt || new Date().toISOString(),
    }));

  const userRecords =
    kindFilter && !kindFilter.includes("user")
      ? []
      : users.map((item) => ({
          tenant,
          kind: "user",
          id: item.id,
          payload: sanitizeUserForSync(item),
          updated_at: item.updatedAt || item.createdAt || new Date().toISOString(),
        }));

  return [
    ...mapRecords(units, "unit"),
    ...mapRecords(photos, "photo", "updatedAt"),
    ...userRecords,
    ...mapRecords(clients, "client"),
    ...mapRecords(projects, "project"),
    ...mapRecords(contacts, "contact"),
    ...mapRecords(contracts, "contract"),
    ...mapRecords(containers, "container"),
    ...mapRecords(materials, "material"),
    ...mapRecords(deliverySkuItems, "deliverySku"),
    ...mapRecords(trashRecords, "trash"),
  ];
}

async function pushCloud({ silent = false, force = false, kinds = null } = {}) {
  if (!force && !can("sync")) return false;
  const endpoint = syncEndpoint();
  if (!endpoint) {
    if (!silent) updateSyncStatus(t("Configurar Supabase URL/Key/Tenant"), true);
    return false;
  }
  if (!navigator.onLine) {
    if (!silent) updateSyncStatus(t("Sem internet para sincronizar"), true);
    return false;
  }

  const records = toCloudRecords({ kinds });
  if (!records.length) return true;

  try {
    if (!silent) updateSyncStatus(t("Enviando dados..."));

    const response = await fetch(`${endpoint}?on_conflict=tenant,kind,id`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: syncConfig.supabaseAnonKey,
        Authorization: `Bearer ${syncConfig.supabaseAnonKey}`,
        Prefer: "resolution=merge-duplicates,return=minimal",
      },
      body: JSON.stringify(records),
    });

    if (!response.ok) throw new Error(await response.text());

    syncConfig.lastSyncAt = new Date().toISOString();
    await saveSetting({ id: "syncConfig", ...syncConfig });
    if (!silent) updateSyncStatus(`${t("Push concluido:")} ${fmtDate(syncConfig.lastSyncAt)}`);
    return true;
  } catch (error) {
    if (!silent) updateSyncStatus(t("Falha no push. Verifique a configuracao."), true);
    if (!silent) console.error(error);
    return false;
  }
}

function newerThan(remoteTs, localTs) {
  return new Date(remoteTs || 0).getTime() > new Date(localTs || 0).getTime();
}

async function pullCloud({ silent = false, force = false, kinds = null, full = false } = {}) {
  if (!force && !can("sync")) return false;
  const endpoint = syncEndpoint();
  if (!endpoint) {
    if (!silent) updateSyncStatus(t("Configurar Supabase URL/Key/Tenant"), true);
    return false;
  }
  if (!navigator.onLine) {
    if (!silent) updateSyncStatus(t("Sem internet para sincronizar"), true);
    return false;
  }

  try {
    if (!silent) updateSyncStatus(t("Baixando dados..."));
    const normalizedKinds = normalizeSyncKinds(kinds);
    const pullCursor = typeof syncConfig.lastPullCursor === "string" ? syncConfig.lastPullCursor : "";
    const qs = new URLSearchParams({
      select: "kind,id,payload,updated_at",
      tenant: `eq.${syncConfig.tenant.trim()}`,
      order: "updated_at.desc",
      limit: "20000",
    });

    if (normalizedKinds) qs.set("kind", `in.(${normalizedKinds.join(",")})`);
    if (!full && pullCursor) qs.set("updated_at", `gt.${pullCursor}`);

    const response = await fetch(`${endpoint}?${qs.toString()}`, {
      headers: {
        apikey: syncConfig.supabaseAnonKey,
        Authorization: `Bearer ${syncConfig.supabaseAnonKey}`,
      },
    });

    if (!response.ok) throw new Error(await response.text());

    const records = await response.json();
    if (!records.length) {
      syncConfig.lastSyncAt = new Date().toISOString();
      await saveSetting({ id: "syncConfig", ...syncConfig });
      if (!silent) updateSyncStatus(`${t("Pull concluido:")} ${fmtDate(syncConfig.lastSyncAt)}`);
      return true;
    }

    let nextPullCursor = pullCursor;
    let hasChanges = false;
    const unitMap = new Map(units.map((item) => [item.id, item]));
    const photoMap = new Map(photos.map((item) => [item.id, item]));
    const userMap = new Map(users.map((item) => [item.id, item]));
    const clientMap = new Map(clients.map((item) => [item.id, item]));
    const projectMap = new Map(projects.map((item) => [item.id, item]));
    const contactMap = new Map(contacts.map((item) => [item.id, item]));
    const contractMap = new Map(contracts.map((item) => [item.id, item]));
    const containerMap = new Map(containers.map((item) => [item.id, item]));
    const materialMap = new Map(materials.map((item) => [item.id, item]));
    const deliverySkuMap = new Map(deliverySkuItems.map((item) => [item.id, item]));
    const trashMap = new Map(trashRecords.map((item) => [item.id, item]));

    for (const row of records) {
      const item = row.payload || {};
      const remoteUpdatedAt = item.updatedAt || row.updated_at || new Date().toISOString();
      item.updatedAt = remoteUpdatedAt;
      if (!nextPullCursor || newerThan(remoteUpdatedAt, nextPullCursor)) nextPullCursor = remoteUpdatedAt;

      if (row.kind === "unit") {
        const local = unitMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(UNIT_STORE, normalizeUnit(item));
          hasChanges = true;
        }
      }

      if (row.kind === "photo") {
        const local = photoMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt || local.createdAt)) {
          await put(PHOTO_STORE, item);
          hasChanges = true;
        }
      }

      if (row.kind === "user") {
        const local = userMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(USER_STORE, normalizeUser(item));
          hasChanges = true;
        }
      }

      if (row.kind === "client") {
        const local = clientMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(CLIENT_STORE, normalizeClient(item));
          hasChanges = true;
        }
      }

      if (row.kind === "project") {
        const local = projectMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(PROJECT_STORE, normalizeProject(item));
          hasChanges = true;
        }
      }

      if (row.kind === "contact") {
        const local = contactMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(CONTACT_STORE, normalizeContact(item));
          hasChanges = true;
        }
      }

      if (row.kind === "contract") {
        const local = contractMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(CONTRACT_STORE, normalizeContract(item));
          hasChanges = true;
        }
      }

      if (row.kind === "container") {
        const local = containerMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(CONTAINER_STORE, normalizeContainer(item));
          hasChanges = true;
        }
      }

      if (row.kind === "material") {
        const local = materialMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(MATERIAL_STORE, normalizeMaterial(item));
          hasChanges = true;
        }
      }

      if (row.kind === "deliverySku") {
        const local = deliverySkuMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(DELIVERY_SKU_STORE, normalizeDeliverySkuItem(item));
          hasChanges = true;
        }
      }

      if (row.kind === "trash") {
        const local = trashMap.get(item.id);
        if (!local || newerThan(remoteUpdatedAt, local.updatedAt)) {
          await put(TRASH_STORE, normalizeTrashRecord(item));
          hasChanges = true;
        }
      }
    }

    syncConfig.lastSyncAt = new Date().toISOString();
    if (nextPullCursor) syncConfig.lastPullCursor = nextPullCursor;
    await saveSetting({ id: "syncConfig", ...syncConfig });
    if (hasChanges) {
      await loadAll();
      if (currentUser) currentUser = users.find((user) => user.id === currentUser.id) || currentUser;
      render();
    }
    if (!silent) updateSyncStatus(`${t("Pull concluido:")} ${fmtDate(syncConfig.lastSyncAt)}`);
    return true;
  } catch (error) {
    if (!silent) updateSyncStatus(t("Falha no pull. Verifique a configuracao."), true);
    if (!silent) console.error(error);
    return false;
  }
}

function queueAutoSync() {
  clearPendingFormChanges();
  if (!syncEndpoint()) return;
  if (!syncConfig.autoSync) return;
  if (!navigator.onLine) return;

  clearTimeout(autoSyncTimer);
  autoSyncTimer = setTimeout(() => {
    pushCloud({ silent: true, force: true });
  }, 1200);
}

function formForSyncLock(target) {
  if (!target) return null;
  if (target.tagName === "FORM") return target;
  if (typeof target.closest === "function") return target.closest("form");
  return null;
}

function setFormSyncDirty(form, dirty) {
  if (!form || !appMain?.contains(form)) return;
  if (dirty) form.dataset.syncDirty = "1";
  else delete form.dataset.syncDirty;
}

function clearPendingFormChanges() {
  if (!appMain) return;
  appMain.querySelectorAll("form[data-sync-dirty='1']").forEach((form) => {
    delete form.dataset.syncDirty;
  });
}

function hasPendingFormChanges() {
  if (!appMain) return false;
  return Boolean(appMain.querySelector("form[data-sync-dirty='1']"));
}

function blockAutoPullTemporarily(ms = AUTO_PULL_SUBMIT_GRACE_MS) {
  autoPullBlockedUntil = Math.max(autoPullBlockedUntil, Date.now() + ms);
}

function isAutoPullTemporarilyBlocked() {
  return Date.now() < autoPullBlockedUntil;
}

async function runAutoPullCycle() {
  if (autoPullInFlight) return;
  if (!currentUser) return;
  if (!navigator.onLine) return;
  if (!syncEndpoint()) return;
  if (isAutoPullTemporarilyBlocked()) return;
  if (hasPendingFormChanges()) return;

  autoPullInFlight = true;
  try {
    await pullCloud({ silent: true, force: true, kinds: AUTO_PULL_KINDS });
  } finally {
    autoPullInFlight = false;
  }
}

function stopAutoPullLoop() {
  if (autoPullTimer) clearInterval(autoPullTimer);
  autoPullTimer = null;
}

function startAutoPullLoop() {
  stopAutoPullLoop();
  if (!currentUser) return;
  if (!navigator.onLine) return;
  if (!syncEndpoint()) return;

  autoPullTimer = setInterval(() => {
    void runAutoPullCycle();
  }, AUTO_PULL_INTERVAL_MS);
  void runAutoPullCycle();
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
      <td>${escapeHtml(unitChecklistQrCodesText(item).join(", ") || "-")}</td>
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
        <thead><tr><th>Code</th><th>Description</th><th>QR Codes</th><th>Expected Qty</th><th>Checked Qty</th><th>Status</th><th>Notes</th></tr></thead>
        <tbody>${checklistRows || '<tr><td colspan="7">No items.</td></tr>'}</tbody>
      </table>

      <h3>Warehouse movements (containers)</h3>
      <table>
        <thead><tr><th>Date</th><th>Container</th><th>Movement</th><th>Qty</th><th>Destination</th><th>Operator</th></tr></thead>
        <tbody>${warehouseRows || '<tr><td colspan="6">No linked movements.</td></tr>'}</tbody>
      </table>

      <h3>Project team</h3>
      <table>
        <thead><tr><th>Name</th><th>Job Title Project</th><th>Company</th><th>Phone</th><th>Email</th></tr></thead>
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

setupForm?.addEventListener("submit", (event) => {
  event.preventDefault();
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

const userAdminPhotoInput = userForm?.querySelector('input[name="photo"]');
const userEditPhotoInput = userEditForm?.querySelector('input[name="photo"]');

userAdminGenderSelect?.addEventListener("change", () => {
  void refreshAdminPhotoPreview();
});

userEditGenderSelect?.addEventListener("change", () => {
  void refreshUserEditPhotoPreview();
});

userAdminPhotoInput?.addEventListener("change", () => {
  void refreshAdminPhotoPreview();
});

userEditPhotoInput?.addEventListener("change", () => {
  void refreshUserEditPhotoPreview();
});

function refreshUserIdCardFromAdminForm() {
  const selected = adminEditingUserId ? users.find((entry) => entry.id === adminEditingUserId) || null : null;
  renderUserIdCard(selected);
}

userForm?.addEventListener("input", () => {
  refreshUserIdCardFromAdminForm();
});

userForm?.addEventListener("change", () => {
  refreshUserIdCardFromAdminForm();
});

signupForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(signupForm);
  const username = data.get("username")?.toString().trim().toLowerCase();
  const firstName = data.get("firstName")?.toString().trim() || "";
  const lastName = data.get("lastName")?.toString().trim() || "";
  const birthDate = normalizeDateField(data.get("birthDate"));
  const fullName = `${firstName} ${lastName}`.trim();
  if (!username) return;
  if (!syncEndpoint()) {
    alert("System database is not available on this device. Contact your administrator.");
    return;
  }
  const syncedUsers = await pullCloud({ silent: true, force: true, kinds: ["user"] });
  if (!syncedUsers) {
    alert("Could not validate users in database right now. Check connection and try again.");
    return;
  }
  if (users.some((user) => user.username === username)) {
    alert(t("Usuario ja existe."));
    return;
  }
  const photoFile = signupForm.querySelector('input[name="photo"]')?.files?.[0];
  const photoDataUrl = photoFile ? await fileToDataUrl(photoFile) : "";
  const employmentType = data.get("employmentType")?.toString() || "";
  const gender = normalizeGender(data.get("gender"));
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
    birthDate,
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
    gender,
    photoDataUrl,
    username,
    passwordHash: await hashPassword(data.get("password")?.toString() || ""),
    accessProfile: "visitor",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
  await put(USER_STORE, normalizeUser(user));
  await loadAll();
  signupForm.reset();
  toggleSubcontractorExtras("signup", "tag");
  showSignupMode(false);
  const pushed = await pushCloud({ silent: true, force: true, kinds: ["user"] });
  if (!pushed) {
    await del(USER_STORE, user.id);
    await loadAll();
    alert("Could not save registration in the central database. Try again.");
    return;
  }
  pushEntityAudit(
    "Users",
    "created",
    `${user.name} (${user.username}) self-registration submitted with access profile "${user.accessProfile}"`,
    "users"
  );
  alert("Registration submitted. An admin or developer can adjust your access profile and permissions.");
  queueAutoSync();
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(loginForm);
  const username = data.get("username")?.toString().trim().toLowerCase();
  const plainPassword = data.get("password")?.toString() || "";
  const passwordHash = await hashPassword(plainPassword);
  if (!syncEndpoint()) {
    alert("System database is not available on this device. Contact your administrator.");
    return;
  }
  const syncedUsers = await pullCloud({ silent: true, force: true, kinds: ["user"] });
  if (!syncedUsers) {
    alert("Could not validate users in database right now. Check connection and try again.");
    return;
  }

  const userByUsername = users.find((entry) => entry.username === username);
  const user = users.find((entry) => entry.username === username && userPasswordMatches(entry, plainPassword, passwordHash));
  if (!user) {
    if (!userByUsername) {
      alert("User is not registered. Click 'Create Account' to register.");
      return;
    }
    alert(t("Usuario ou senha invalidos."));
    return;
  }

  const shouldMigratePassword =
    Boolean(userLegacyPassword(user)) ||
    !isSha256Hash(user.passwordHash) ||
    String(user.passwordHash || "").trim().toLowerCase() !== passwordHash.toLowerCase();

  if (shouldMigratePassword) {
    const migratedUser = normalizeUser({
      ...user,
      passwordHash,
      legacyPassword: "",
      password: "",
      updatedAt: new Date().toISOString(),
    });
    await put(USER_STORE, migratedUser);
    await loadAll();
    currentUser = users.find((entry) => entry.id === migratedUser.id) || migratedUser;
    await pushCloud({ silent: true, force: true, kinds: ["user"] });
  } else {
    currentUser = user;
  }

  hasAutoDispatchedCoiReminder = false;
  setSession(user.id);
  pushAppAudit("Session login", "session", "auth");
  window.history.replaceState(null, "", "#home");
  loginForm.reset();
  await pullCloud({ silent: true, force: true });
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
  const birthDate = normalizeDateField(data.get("birthDate"));
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
  const selectedAccessProfile = data.get("accessProfile")?.toString() || "visitor";
  const gender = normalizeGender(data.get("gender"));

  if (!firstName || !lastName || !companyName || !jobTitle || !address || !cellPhone || !email || !username) {
    alert("Fill in all required fields.");
    return;
  }
  if (!isDeveloper() && selectedAccessProfile === "developer") {
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
    birthDate,
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
    gender,
    photoDataUrl,
    username,
    passwordHash,
    accessProfile: selectedAccessProfile,
    createdAt: targetUser?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });

  await put(USER_STORE, savedUser);
  if (!targetUser) {
    pushEntityAudit(
      "Users",
      "created",
      `${savedUser.name} (${savedUser.username}) access profile "${savedUser.accessProfile}" company "${auditValue(savedUser.companyName)}"`,
      "users"
    );
  } else {
    const changes = collectAuditChanges(targetUser, savedUser, [
      { key: "firstName", label: "First name" },
      { key: "lastName", label: "Last name" },
      { key: "birthDate", label: "Date of birth" },
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
      { key: "gender", label: "Gender" },
      { key: "username", label: "Username" },
      { key: "accessProfile", label: "Access Profile" },
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
  adminEditingUserId = "";
  resetAdminUserForm();
  setUserAdminFormOpen(false);
  setUsersSubView("directory");
  render();
  queueAutoSync();
});

usersRegistrationTabBtn?.addEventListener("click", () => {
  if (!can("manageUsers")) return;
  openUsersRegistrationClean();
});

usersDirectoryTabBtn?.addEventListener("click", () => {
  if (!can("manageUsers")) return;
  setUsersSubView("directory");
});

userDirectoryEditBtn?.addEventListener("click", () => {
  if (!can("manageUsers")) return;
  if (!adminEditingUserId) {
    alert("Select a user from the list first.");
    return;
  }
  const selected = users.find((entry) => entry.id === adminEditingUserId);
  if (!selected) {
    alert("Selected user is no longer available.");
    return;
  }
  populateAdminUserForm(selected);
  setUserAdminFormOpen(true);
  setUsersSubView("registration");
  userForm?.querySelector('input[name="firstName"]')?.focus();
});

userFormNewBtn?.addEventListener("click", () => {
  if (!can("manageUsers")) return;
  resetAdminUserForm();
  setUserAdminFormOpen(true);
  setUsersSubView("registration");
  syncSelectedUserRow();
  userForm?.querySelector('input[name="firstName"]')?.focus();
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
  if (!isDeveloper() && (userAccessProfile(targetUser) === "developer" || isPrimaryDeveloperUser(targetUser))) {
    alert("Developer account cannot be deleted by admin.");
    return;
  }
  const confirmed = confirmDeleteAction(`user "${targetUser.name} (${targetUser.username})"`);
  if (!confirmed) return;
  pushEntityAudit(
    "Users",
    "deleted",
    `${targetUser.name} (${targetUser.username}) access profile "${userAccessProfile(targetUser)}"`,
    "users"
  );
  await trashDeleteRecord(USER_STORE, targetUser, {
    label: `User: ${targetUser.name} (${targetUser.username})`,
    scope: "users",
  });
  await loadAll();
  if (currentUser) currentUser = users.find((entry) => entry.id === currentUser.id) || currentUser;
  resetAdminUserForm();
  setUserAdminFormOpen(false);
  setUsersSubView("directory");
  render();
  queueAutoSync();
});

userFormCloseBtn?.addEventListener("click", () => {
  adminEditingUserId = "";
  resetAdminUserForm();
  setUserAdminFormOpen(false);
  setUsersSubView("directory");
  syncSelectedUserRow();
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
  if (!isDeveloper() && userAccessProfile(targetUser) === "developer" && targetUser.id !== currentUser.id) return;
  if (!canAccessEmployeeRecord(targetUser)) return;

  const data = new FormData(userEditForm);
  const firstName = data.get("firstName")?.toString().trim() || "";
  const lastName = data.get("lastName")?.toString().trim() || "";
  const birthDate = normalizeDateField(data.get("birthDate"));
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
  const gender = normalizeGender(data.get("gender"));

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
    birthDate,
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
    gender,
    photoDataUrl,
    updatedAt: new Date().toISOString(),
  });

  await put(USER_STORE, savedUser);
  const changes = collectAuditChanges(targetUser, savedUser, [
    { key: "firstName", label: "First name" },
    { key: "lastName", label: "Last name" },
    { key: "birthDate", label: "Date of birth" },
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
    { key: "gender", label: "Gender" },
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
  selectedProjectId = "";
  selectedContractId = "";
  keepClientFormBlank = true;
  selectedClientDetailsProjectId = "";
  setClientsSectionMode("create");
  await loadAll();
  resetClientForm();
  resetProjectForm({ clientId: selectedClientId });
  resetContractForm({ clientId: selectedClientId });
  if (contactClientSelect) {
    contactClientSelect.value = selectedClientId;
    syncContactProjectSelect();
  }
  render();
  queueAutoSync();
});

clientSearchInput?.addEventListener("input", () => {
  clientSearchQuery = (clientSearchInput.value || "").trim();
  renderClientsTable();
});

clientDetailsBackBtn?.addEventListener("click", () => {
  selectedClientDetailsProjectId = "";
  keepClientFormBlank = true;
  setClientsSectionMode("create");
  resetClientForm();
});

goProjectsFromClientBtn?.addEventListener("click", () => {
  openClientRegistrationShortcut("projects");
});

goPeopleFromClientBtn?.addEventListener("click", () => {
  openClientRegistrationShortcut("people");
});

function addProjectScopeExtraDraft(rawValue) {
  const value = String(rawValue || "").trim();
  if (!value) return;
  projectScopeExtrasDraft = uniqueTextList([...projectScopeExtrasDraft, value]);
  renderProjectDraftPills();
}

function addProjectChecklistExtraDraft(rawValue) {
  const value = String(rawValue || "").trim();
  if (!value) return;
  projectChecklistExtrasDraft = uniqueTextList([...projectChecklistExtrasDraft, value]);
  renderProjectDraftPills();
}

projectScopeExtraAddBtn?.addEventListener("click", () => {
  addProjectScopeExtraDraft(projectScopeExtraInput?.value || "");
  if (projectScopeExtraInput) projectScopeExtraInput.value = "";
  projectScopeExtraInput?.focus();
});

projectChecklistExtraAddBtn?.addEventListener("click", () => {
  addProjectChecklistExtraDraft(projectChecklistExtraInput?.value || "");
  if (projectChecklistExtraInput) projectChecklistExtraInput.value = "";
  projectChecklistExtraInput?.focus();
});

projectScopeExtraInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  addProjectScopeExtraDraft(projectScopeExtraInput.value);
  projectScopeExtraInput.value = "";
});

projectChecklistExtraInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  addProjectChecklistExtraDraft(projectChecklistExtraInput.value);
  projectChecklistExtraInput.value = "";
});

projectScopeExtraList?.addEventListener("click", (event) => {
  const btn = event.target.closest("[data-project-scope-pill-remove]");
  if (!btn) return;
  const index = Number(btn.dataset.projectScopePillRemove);
  if (!Number.isInteger(index) || index < 0 || index >= projectScopeExtrasDraft.length) return;
  const value = projectScopeExtrasDraft[index];
  const confirmed = confirmDeleteAction(`scope item "${value || index}"`, { restorable: false });
  if (!confirmed) return;
  projectScopeExtrasDraft = projectScopeExtrasDraft.filter((_, idx) => idx !== index);
  renderProjectDraftPills();
});

projectChecklistExtraList?.addEventListener("click", (event) => {
  const btn = event.target.closest("[data-project-checklist-pill-remove]");
  if (!btn) return;
  const index = Number(btn.dataset.projectChecklistPillRemove);
  if (!Number.isInteger(index) || index < 0 || index >= projectChecklistExtrasDraft.length) return;
  const value = projectChecklistExtrasDraft[index];
  const confirmed = confirmDeleteAction(`checklist template item "${value || index}"`, { restorable: false });
  if (!confirmed) return;
  projectChecklistExtrasDraft = projectChecklistExtrasDraft.filter((_, idx) => idx !== index);
  renderProjectDraftPills();
});

clientFormNewBtn?.addEventListener("click", () => {
  if (!can("manageCatalog")) return;
  selectedClientId = "";
  selectedProjectId = "";
  selectedContractId = "";
  selectedClientDetailsProjectId = "";
  keepClientFormBlank = true;
  setClientsSectionMode("create");
  resetClientForm();
  resetProjectForm();
  resetContractForm();
  render();
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
  if (contracts.some((contract) => contract.clientId === targetId)) {
    alert("Cannot delete client with linked contracts.");
    return;
  }
  if (containers.some((container) => container.clientId === targetId)) {
    alert(t("Nao e possivel excluir cliente com containers vinculados."));
    return;
  }
  const targetClient = clients.find((entry) => entry.id === targetId);
  if (!targetClient) return;
  const confirmed = confirmDeleteAction(`client "${targetClient.name}"`);
  if (!confirmed) return;
  if (targetClient) pushEntityAudit("Clients", "deleted", `${targetClient.name}`, "clients");
  await trashDeleteRecord(CLIENT_STORE, targetClient, {
    label: `Client: ${targetClient.name}`,
    scope: "clients",
  });
  selectedClientId = "";
  selectedProjectId = "";
  selectedContractId = "";
  selectedClientDetailsProjectId = "";
  keepClientFormBlank = true;
  setClientsSectionMode("create");
  resetClientForm();
  resetProjectForm();
  resetContractForm();
  await loadAll();
  render();
  queueAutoSync();
});

projectFormNewBtn?.addEventListener("click", () => {
  if (!can("manageCatalog")) return;
  selectedProjectId = "";
  resetProjectForm({ clientId: selectedClientId || projectClientSelect?.value || "" });
  render();
});

projectFormDeleteBtn?.addEventListener("click", async () => {
  if (!can("manageCatalog")) return;
  const targetId = projectForm?.elements?.projectId?.value || "";
  if (!targetId) return;
  if (units.some((unit) => unit.projectId === targetId)) {
    alert(t("Nao e possivel excluir projeto com unidades vinculadas."));
    return;
  }
  if (contacts.some((contact) => contact.projectId === targetId)) {
    alert(t("Nao e possivel excluir projeto com pessoas vinculadas."));
    return;
  }
  if (contracts.some((contract) => contract.projectId === targetId)) {
    alert("Cannot delete project with linked contracts.");
    return;
  }
  if (containers.some((container) => container.projectId === targetId)) {
    alert(t("Nao e possivel excluir projeto com containers vinculados."));
    return;
  }
  const targetProject = projects.find((entry) => entry.id === targetId);
  if (!targetProject) return;
  const confirmed = confirmDeleteAction(`project "${targetProject.name}"`);
  if (!confirmed) return;
  pushEntityAudit("Projects", "deleted", `${targetProject.name}`, "projects");
  await trashDeleteRecord(PROJECT_STORE, targetProject, {
    label: `Project: ${targetProject.name}`,
    scope: "projects",
  });
  if (selectedProjectId === targetId) selectedProjectId = "";
  resetProjectForm({ clientId: selectedClientId || targetProject.clientId || "" });
  await loadAll();
  render();
  queueAutoSync();
});

projectForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(projectForm);
  const projectId = data.get("projectId")?.toString() || "";
  const targetProject = projectId ? projects.find((entry) => entry.id === projectId) : null;
  const clientId = data.get("clientId")?.toString();
  const name = data.get("name")?.toString().trim();
  const code = data.get("code")?.toString().trim() || "";
  const address = data.get("address")?.toString().trim() || "";
  const floorsCount = data.get("floorsCount")?.toString().trim() || "";
  const apartmentsCount = data.get("apartmentsCount")?.toString().trim() || "";
  const scopeCategories = collectCheckedValues(projectForm, "scopeCategories");
  const scopeExtras = uniqueTextList(collectCheckedValues(projectForm, "scopeExtrasDynamic"));
  const checklistBase = collectCheckedValues(projectForm, "unitChecklistBase");
  const checklistExtras = uniqueTextList(collectCheckedValues(projectForm, "unitChecklistExtrasDynamic"));
  const unitChecklistTemplate = uniqueTextList([...checklistBase, ...checklistExtras]);

  if (!clientId || !name) {
    alert(t("Cliente e nome do projeto sao obrigatorios."));
    return;
  }

  if (
    projects.some(
      (project) => project.clientId === clientId && project.name.toLowerCase() === name.toLowerCase() && project.id !== targetProject?.id
    )
  ) {
    alert(t("Projeto ja cadastrado para este cliente."));
    return;
  }

  if (targetProject && targetProject.clientId !== clientId && contracts.some((contract) => contract.projectId === targetProject.id)) {
    alert("This project has linked contracts. Keep the same client or update contracts first.");
    return;
  }

  const savedProject = normalizeProject({
    ...(targetProject || {}),
    id: targetProject?.id || uid(),
    clientId,
    name,
    code,
    address,
    floorsCount,
    apartmentsCount,
    scopeCategories,
    scopeExtras,
    unitChecklistTemplate,
    createdAt: targetProject?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });

  await put(PROJECT_STORE, savedProject);
  const linkedClient = clientById(savedProject.clientId);
  if (!targetProject) {
    pushEntityAudit(
      "Projects",
      "created",
      `${savedProject.name} | client "${auditValue(linkedClient?.name || "-")}" | code "${auditValue(savedProject.code)}" | scope "${auditValue(
        projectScopeSummary(savedProject)
      )}"`,
      "projects"
    );
  } else {
    const changes = collectAuditChanges(targetProject, savedProject, [
      { key: "name", label: "Project name" },
      { key: "code", label: "Project code" },
      { key: "address", label: "Project address" },
      { key: "floorsCount", label: "Floors" },
      { key: "apartmentsCount", label: "Apartments" },
      { key: "scopeCategories", label: "Scope categories", map: (value) => ensureArray(value).join(", ") },
      { key: "scopeExtras", label: "Scope extras", map: (value) => ensureArray(value).join(", ") },
      { key: "unitChecklistTemplate", label: "Unit checklist template", map: (value) => ensureArray(value).join(", ") },
    ]);
    pushEntityAudit(
      "Projects",
      "updated",
      `${savedProject.name}${changes.length ? ` | ${changes.join("; ")}` : " | no field changes"}`,
      "projects"
    );
  }

  selectedClientId = clientId;
  selectedProjectId = "";
  await loadAll();
  resetProjectForm({ clientId: selectedClientId });
  render();
  queueAutoSync();
});

contactForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(contactForm);
  const name = data.get("name")?.toString().trim();
  if (!name) return;
  const clientId = data.get("clientId")?.toString() || selectedClientId || "";
  const projectId = data.get("projectId")?.toString() || "";
  if (!clientId) {
    alert("Select a client before adding people to project.");
    return;
  }
  if (projectId) {
    const linkedProject = projectById(projectId);
    if (linkedProject && linkedProject.clientId !== clientId) {
      alert("Selected project does not belong to the selected client.");
      return;
    }
  }

  const createdContact = normalizeContact({
    id: uid(),
    name,
    role: data.get("role")?.toString(),
    company: data.get("company")?.toString().trim(),
    clientId,
    projectId,
    phone: data.get("phone")?.toString().trim(),
    email: data.get("email")?.toString().trim(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });

  await put(CONTACT_STORE, createdContact);
  pushEntityAudit(
    "Contacts",
    "created",
    `${createdContact.name} | job title project "${auditValue(contactRoleLabel(createdContact.role))}" | company "${auditValue(createdContact.company)}"`,
    "contacts"
  );

  selectedClientId = clientId;
  contactForm.reset();
  if (contactClientSelect) contactClientSelect.value = selectedClientId;
  syncContactProjectSelect();
  await loadAll();
  render();
  queueAutoSync();
});

contractFormNewBtn?.addEventListener("click", () => {
  if (!can("manageCatalog")) return;
  resetContractForm({ clientId: selectedClientId || contractClientSelect?.value || "" });
  renderContractsTable();
});

contractFormDeleteBtn?.addEventListener("click", async () => {
  if (!can("manageCatalog")) return;
  const targetId = contractForm?.elements?.contractId?.value || selectedContractId || "";
  if (!targetId) return;
  const targetContract = contractById(targetId);
  if (!targetContract) return;
  const confirmed = confirmDeleteAction(`contract "${targetContract.title || targetContract.contractCode || targetContract.id}"`);
  if (!confirmed) return;
  pushEntityAudit(
    "Contracts",
    "deleted",
    `${targetContract.title} | code "${auditValue(targetContract.contractCode)}"`,
    "contracts"
  );
  await trashDeleteRecord(CONTRACT_STORE, targetContract, {
    label: `Contract: ${targetContract.title || targetContract.contractCode || targetContract.id}`,
    scope: "contracts",
  });
  resetContractForm({ clientId: selectedClientId || targetContract.clientId || "" });
  await loadAll();
  render();
  queueAutoSync();
});

contractForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("manageCatalog")) return;

  const data = new FormData(contractForm);
  const contractId = data.get("contractId")?.toString() || "";
  const targetContract = contractId ? contractById(contractId) : null;
  const clientId = data.get("clientId")?.toString().trim() || "";
  const projectId = data.get("projectId")?.toString().trim() || "";
  const title = data.get("title")?.toString().trim() || "";
  const contractCode = data.get("contractCode")?.toString().trim() || "";
  const status = data.get("status")?.toString().trim() || "draft";
  const signedDate = data.get("signedDate")?.toString().trim() || "";
  const startDate = data.get("startDate")?.toString().trim() || "";
  const endDate = data.get("endDate")?.toString().trim() || "";
  const amount = data.get("amount")?.toString().trim() || "";
  const notes = data.get("notes")?.toString().trim() || "";

  if (!clientId || !title) {
    alert("Client and contract title are required.");
    return;
  }

  if (projectId) {
    const linkedProject = projectById(projectId);
    if (!linkedProject) {
      alert("Selected project was not found.");
      return;
    }
    if (linkedProject.clientId !== clientId) {
      alert("Selected project does not belong to the selected client.");
      return;
    }
  }

  if (
    contractCode &&
    contracts.some(
      (entry) => entry.clientId === clientId && entry.contractCode.toLowerCase() === contractCode.toLowerCase() && entry.id !== targetContract?.id
    )
  ) {
    alert("This contract code already exists for the selected client.");
    return;
  }

  const nowIso = new Date().toISOString();
  const savedContract = normalizeContract({
    ...(targetContract || {}),
    id: targetContract?.id || uid(),
    clientId,
    projectId,
    title,
    contractCode,
    status,
    signedDate,
    startDate,
    endDate,
    amount,
    notes,
    createdAt: targetContract?.createdAt || nowIso,
    updatedAt: nowIso,
  });

  await put(CONTRACT_STORE, savedContract);
  const linkedClient = clientById(savedContract.clientId);
  const linkedProject = projectById(savedContract.projectId);
  if (!targetContract) {
    pushEntityAudit(
      "Contracts",
      "created",
      `${savedContract.title} | client "${auditValue(linkedClient?.name || "-")}" | project "${auditValue(linkedProject?.name || "-")}" | status "${auditValue(contractStatusLabel(savedContract.status))}"`,
      "contracts"
    );
  } else {
    const changes = collectAuditChanges(targetContract, savedContract, [
      { key: "title", label: "Title" },
      { key: "contractCode", label: "Code" },
      { key: "status", label: "Status", map: (value) => contractStatusLabel(value) },
      { key: "signedDate", label: "Signed date" },
      { key: "startDate", label: "Start date" },
      { key: "endDate", label: "End date" },
      { key: "amount", label: "Amount (USD)" },
      { key: "notes", label: "Notes" },
      { key: "clientId", label: "Client", map: (value) => clientById(value)?.name || "-" },
      { key: "projectId", label: "Project", map: (value) => projectById(value)?.name || "-" },
    ]);
    pushEntityAudit(
      "Contracts",
      "updated",
      `${savedContract.title}${changes.length ? ` | ${changes.join("; ")}` : " | no field changes"}`,
      "contracts"
    );
  }

  selectedClientId = clientId;
  selectedContractId = "";
  await loadAll();
  resetContractForm({ clientId: selectedClientId, projectId: projectId || "" });
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
projectClientSelect.addEventListener("change", () => {
  selectedClientId = projectClientSelect.value || selectedClientId;
  if (contactClientSelect && projectClientSelect.value) {
    contactClientSelect.value = projectClientSelect.value;
    syncContactProjectSelect();
  }
  if (contractClientSelect && projectClientSelect.value) {
    contractClientSelect.value = projectClientSelect.value;
    syncContractProjectSelect();
  }
  renderProjectsTable();
  renderContactsTable();
  renderContractsTable();
});
contactClientSelect.addEventListener("change", () => {
  if (contactClientSelect.value) selectedClientId = contactClientSelect.value;
  syncContactProjectSelect();
  renderContactsTable();
  renderContractsTable();
});
contractClientSelect?.addEventListener("change", () => {
  if (contractClientSelect.value) selectedClientId = contractClientSelect.value;
  syncContractProjectSelect();
  renderContractsTable();
});
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
  pushAppAudit(
    `[Unit ${unit.unitCode || unit.id}] Unit created${unit.checkItems.length ? ` with ${unit.checkItems.length} checklist template item(s)` : ""}`,
    "unit-change",
    unit.projectName || "-"
  );
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
  if (qrFlowCodeInput && lastQrLookupCode) qrFlowCodeInput.value = lastQrLookupCode;
  renderQrLookupPanel();
});
qrLookupInput?.addEventListener("input", () => {
  if (!qrLookupInput.value.trim()) {
    lastQrLookupCode = "";
    renderQrLookupPanel();
  }
});
qrLookupScanBtn?.addEventListener("click", async () => {
  await scanQrIntoInput(qrLookupInput, "Scan QR to search history");
  lastQrLookupCode = qrLookupInput?.value?.trim() || "";
  if (qrFlowCodeInput && lastQrLookupCode) qrFlowCodeInput.value = lastQrLookupCode;
  renderQrLookupPanel();
});
qrFlowScanBtn?.addEventListener("click", async () => {
  if (!canUpdateQrWorkflow()) return;
  const scanned = await scanQrIntoInput(qrFlowCodeInput, "Scan QR to update workflow");
  syncQrFlowStageOptions();
  if (!scanned) return;
  await submitQrFlowUpdate({ fromScan: true });
});
qrFlowStageSelect?.addEventListener("change", () => {
  updateQrFlowStageCustomVisibility();
});
qrFlowStatusSelect?.addEventListener("change", () => {
  updateQrFlowIssueVisibility();
});
qrFlowForm?.addEventListener("reset", () => {
  setTimeout(() => {
    syncQrFlowStageOptions({ preserveSelection: false });
    if (qrFlowStatusSelect) qrFlowStatusSelect.value = "in-progress";
    updateQrFlowStageCustomVisibility();
    updateQrFlowIssueVisibility();
    renderQrFlowAccessHint();
    setQrFlowResult("");
  }, 0);
});
qrFlowForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  await submitQrFlowUpdate();
});

deliveryAssignClientSelect?.addEventListener("change", () => {
  syncDeliveryAssignProjectSelect();
  if (deliveryAssignProjectSelect) deliveryAssignProjectSelect.value = "";
  syncDeliveryAssignUnitSelect();
});

deliveryAssignProjectSelect?.addEventListener("change", () => {
  syncDeliveryAssignUnitSelect();
  if (deliveryAssignUnitSelect) deliveryAssignUnitSelect.value = "";
});

deliveryScanBtn?.addEventListener("click", async () => {
  if (deliveryScanBtn.disabled) return;
  await scanQrIntoInput(deliveryScanQrInput, "Scan delivery inventory QR");
});

deliveryImportTargetSelect?.addEventListener("change", () => {
  refreshDeliveryImportTargetUi();
  syncDeliveryImportApplyState();
});

deliveryImportFileInput?.addEventListener("change", () => {
  clearDeliveryImportDraft({ clearFile: false, clearStatus: true });
  const file = deliveryImportFileInput.files?.[0];
  if (!file) return;
  const info = detectImportFileKind(file);
  setDeliveryImportStatus(`Selected file: ${file.name} (${info.label}, ${formatBytes(file.size)}). Click Analyze file.`);
});

deliveryImportAnalyzeBtn?.addEventListener("click", async () => {
  if (deliveryImportAnalyzeBtn.disabled) return;
  await analyzeDeliveryImportFile();
});

deliveryImportApplyBtn?.addEventListener("click", async () => {
  if (deliveryImportApplyBtn.disabled) return;
  await applyDeliveryImportDraft();
});

deliveryImportClearBtn?.addEventListener("click", () => {
  if (deliveryImportClearBtn.disabled) return;
  clearDeliveryImportDraft({ clearFile: true, clearStatus: true });
});

ocrImportTargetSelect?.addEventListener("change", () => {
  refreshOcrImportTargetUi();
  syncOcrImportApplyState();
});

ocrImportFileInput?.addEventListener("change", () => {
  clearOcrImportDraft({ clearFile: false, clearStatus: true });
  const file = ocrImportFileInput.files?.[0];
  if (!file) return;
  const info = detectImportFileKind(file);
  setOcrImportStatus(`Selected file: ${file.name} (${info.label}, ${formatBytes(file.size)}). Click Run OCR / Analyze.`);
});

ocrImportAnalyzeBtn?.addEventListener("click", async () => {
  if (ocrImportAnalyzeBtn.disabled) return;
  await analyzeOcrImportFile();
});

ocrImportApplyBtn?.addEventListener("click", async () => {
  if (ocrImportApplyBtn.disabled) return;
  await applyOcrImportDraft();
});

ocrImportClearBtn?.addEventListener("click", () => {
  if (ocrImportClearBtn.disabled) return;
  clearOcrImportDraft({ clearFile: true, clearStatus: true });
});

deliveryInventorySeedBtn?.addEventListener("click", async () => {
  if (deliveryInventorySeedBtn.disabled) return;
  const seedRows = deliverySkuSeedRows();
  const hasExistingRows = deliverySkuItems.length > 0;
  if (hasExistingRows) {
    const proceed = window.confirm(
      "Inventory already has rows. Reloading seed will update container baseline quantities and keep your current assignments/scans."
    );
    if (!proceed) return;
  }
  const now = new Date().toISOString();
  const containerSources = new Set(["Container 2", "Container 3"]);
  const existingByKey = new Map(deliverySkuItems.map((entry) => [deliverySkuKey(entry.sku, entry.finish), entry]));
  const output = [...deliverySkuItems];

  seedRows.forEach((seedItem) => {
    const key = deliverySkuKey(seedItem.sku, seedItem.finish);
    const existing = existingByKey.get(key);
    if (!existing) {
      output.push(seedItem);
      return;
    }

    const preservedRefs = ensureArray(existing.sourceRefs).filter((entry) => !containerSources.has(entry.source));
    const updated = normalizeDeliverySkuItem({
      ...existing,
      totalQty: Math.max(existing.totalQty, seedItem.totalQty),
      sourceRefs: [...seedItem.sourceRefs, ...preservedRefs],
      updatedAt: now,
    });
    const index = output.findIndex((entry) => entry.id === existing.id);
    if (index >= 0) output[index] = updated;
  });

  await writeDeliverySkuItems(output, { replace: true });
  pushEntityAudit("Delivery Inventory", "seeded", `Loaded ${seedRows.length} consolidated rows from container lists`, "delivery");
  setDeliveryInventoryStatus(`Inventory loaded: ${seedRows.length} unique material variants cataloged with QR.`);
});

deliveryInventoryResetBtn?.addEventListener("click", async () => {
  if (deliveryInventoryResetBtn.disabled) return;
  const totalRows = deliverySkuItems.length;
  const confirmed = window.confirm(
    `This will remove all delivery inventory rows (${totalRows}). You can restore this reset for up to ${RESTORE_WINDOW_LABEL}. Continue?`
  );
  if (!confirmed) return;
  if (deliverySkuItems.length) {
    await saveTrashRecord({
      storeName: DELIVERY_SKU_STORE,
      recordId: "delivery-reset",
      label: `Delivery inventory reset (${deliverySkuItems.length} rows)`,
      scope: "delivery",
      payload: deliverySkuItems,
      restoreType: "delivery-inventory-reset",
      context: {},
    });
  }
  await writeDeliverySkuItems([], { replace: true });
  pushEntityAudit("Delivery Inventory", "cleared", "All delivery inventory rows removed", "delivery");
  setDeliveryInventoryStatus("Delivery inventory cleared.");
  resetDeliveryScanDestination();
  if (deliveryScanQrInput) deliveryScanQrInput.value = "";
});

deliveryAddForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!deliveryAddForm || deliveryAddForm.querySelector("button[type='submit']")?.disabled) return;
  const sku = normalizeDeliverySkuValue(deliveryAddSkuInput?.value || "");
  const finish = normalizeFinishValue(deliveryAddFinishInput?.value || "");
  const qty = Math.max(1, Math.trunc(toNumber(deliveryAddQtyInput?.value || 1)));
  const description = String(deliveryAddDescriptionInput?.value || "").trim();
  if (!sku || !finish || qty <= 0) {
    setDeliveryInventoryStatus("Enter SKU, finish, and a valid quantity.", true);
    return;
  }

  const matchKey = deliverySkuKey(sku, finish);
  const existing = deliverySkuItems.find((entry) => deliverySkuKey(entry.sku, entry.finish) === matchKey);
  const now = new Date().toISOString();

  if (existing) {
    const updatedItem = normalizeDeliverySkuItem({
      ...existing,
      description: description || existing.description || sku,
      totalQty: existing.totalQty + qty,
      sourceRefs: [
        ...ensureArray(existing.sourceRefs),
        { source: "Manual entry", qty },
      ],
      updatedAt: now,
    });
    await saveDeliverySkuItem(updatedItem);
    pushEntityAudit(
      "Delivery Inventory",
      "updated",
      `${updatedItem.sku} ${updatedItem.finish}: total qty +${qty} (new total ${updatedItem.totalQty})`,
      "delivery"
    );
    if (deliveryScanQrInput) deliveryScanQrInput.value = updatedItem.qrCode;
    setDeliveryInventoryStatus(`Quantity increased for ${updatedItem.sku} (${updatedItem.finish}).`);
  } else {
    const createdItem = normalizeDeliverySkuItem({
      id: uid(),
      sku,
      description: description || sku,
      finish,
      totalQty: qty,
      scannedQty: 0,
      sourceRefs: [{ source: "Manual entry", qty }],
      scanLog: [],
      createdAt: now,
      updatedAt: now,
    });
    await saveDeliverySkuItem(createdItem);
    pushEntityAudit(
      "Delivery Inventory",
      "created",
      `${createdItem.sku} ${createdItem.finish}: qty ${createdItem.totalQty} with QR ${createdItem.qrCode}`,
      "delivery"
    );
    if (deliveryScanQrInput) deliveryScanQrInput.value = createdItem.qrCode;
    setDeliveryInventoryStatus(`Item created: ${createdItem.sku} (${createdItem.finish}).`);
  }

  deliveryAddForm.reset();
  if (deliveryAddQtyInput) deliveryAddQtyInput.value = "1";
});

deliverySaveDestinationBtn?.addEventListener("click", async () => {
  if (deliverySaveDestinationBtn.disabled) return;
  const qrCode = String(deliveryScanQrInput?.value || "").trim();
  if (!qrCode) {
    setDeliveryInventoryStatus("Scan or paste a QR code first to save destination.", true);
    return;
  }
  const target = deliverySkuByQrCode(qrCode);
  if (!target) {
    setDeliveryInventoryStatus(`QR not found in delivery inventory: ${qrCode}`, true);
    return;
  }
  const now = new Date().toISOString();
  const updatedItem = normalizeDeliverySkuItem({
    ...target,
    clientId: deliveryAssignClientSelect?.value || "",
    projectId: deliveryAssignProjectSelect?.value || "",
    unitId: deliveryAssignUnitSelect?.value || "",
    destinationNote: String(deliveryDestinationNoteInput?.value || "").trim(),
    updatedAt: now,
  });
  await saveDeliverySkuItem(updatedItem);
  pushEntityAudit(
    "Delivery Inventory",
    "destination-updated",
    `${updatedItem.sku} ${updatedItem.finish}: ${deliveryDestinationLabel(updatedItem)}`,
    "delivery"
  );
  setDeliveryInventoryStatus(`Destination saved for ${updatedItem.sku} (${updatedItem.finish}).`);
});

deliveryScanForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (deliveryScanForm.querySelector("button[type='submit']")?.disabled) return;

  const qrCode = String(deliveryScanQrInput?.value || "").trim();
  const qty = Math.max(1, Math.trunc(toNumber(deliveryScanQtyInput?.value || 1)));
  if (!qrCode || qty <= 0) {
    setDeliveryInventoryStatus("Scan a QR code and enter quantity before consuming.", true);
    return;
  }

  const item = deliverySkuByQrCode(qrCode);
  if (!item) {
    setDeliveryInventoryStatus(`QR not found in delivery inventory: ${qrCode}`, true);
    return;
  }

  const availableBefore = deliverySkuAvailableQty(item);
  if (availableBefore < qty) {
    setDeliveryInventoryStatus(
      `Insufficient available quantity for ${item.sku} (${item.finish}). Available: ${availableBefore}.`,
      true
    );
    return;
  }

  const now = new Date().toISOString();
  const clientId = deliveryAssignClientSelect?.value || item.clientId || "";
  const projectId = deliveryAssignProjectSelect?.value || item.projectId || "";
  const unitId = deliveryAssignUnitSelect?.value || item.unitId || "";
  const destinationNote = String(deliveryDestinationNoteInput?.value || item.destinationNote || "").trim();

  const updatedItem = normalizeDeliverySkuItem({
    ...item,
    scannedQty: item.scannedQty + qty,
    clientId,
    projectId,
    unitId,
    destinationNote,
    scanLog: [
      ...ensureArray(item.scanLog),
      {
        id: uid(),
        qty,
        scannedAt: now,
        scannedByUserId: currentUser?.id || "",
        scannedByName: currentUser?.name || "User",
        clientId,
        projectId,
        unitId,
        destinationNote,
      },
    ],
    updatedAt: now,
  });

  await saveDeliverySkuItem(updatedItem);
  const availableAfter = deliverySkuAvailableQty(updatedItem);
  pushEntityAudit(
    "Delivery Inventory",
    "consumed-by-qr",
    `${updatedItem.sku} ${updatedItem.finish}: -${qty} (available ${availableAfter}) via ${updatedItem.qrCode}`,
    "delivery"
  );
  setDeliveryInventoryStatus(
    `QR processed: ${updatedItem.sku} (${updatedItem.finish}) -${qty}. Remaining available: ${availableAfter}.`
  );
  if (deliveryScanQtyInput) deliveryScanQtyInput.value = "1";
});

projectReportBtn.addEventListener("click", () => {
  if (!can("report")) return;
  generateProjectReport(projectReportSelect.value);
});

function buildFullBackupPayload() {
  return {
    exportedAt: new Date().toISOString(),
    units,
    photos,
    users,
    clients,
    projects,
    contacts,
    contracts,
    containers,
    materials,
    deliverySkuItems,
    trashRecords,
    settings: [{ id: "syncConfig", ...syncConfig }],
  };
}

function triggerLocalBackupDownload(source = "toolbar") {
  if (!isDeveloper()) return false;
  const payload = buildFullBackupPayload();
  const counts = backupCountsSnapshot();
  pushEntityAudit("Backups", "exported", `${backupCountsSummary(counts)} | source:${source}`, "backup");

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `cabinets-control-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);

  try {
    localStorage.setItem(LOCAL_BACKUP_META_KEY, payload.exportedAt);
  } catch {
    // Keep backup download flow even if local storage is unavailable.
  }
  return true;
}

async function importBackupFile(file) {
  if (!isDeveloper() || !file) return false;
  try {
    const text = await file.text();
    const payload = JSON.parse(text);

    if (!Array.isArray(payload.units) || !Array.isArray(payload.photos)) {
      throw new Error("Invalid backup format");
    }

    for (const unit of payload.units) await put(UNIT_STORE, normalizeUnit(unit));
    for (const photo of payload.photos) await put(PHOTO_STORE, photo);
    if (Array.isArray(payload.users)) for (const user of payload.users) await put(USER_STORE, normalizeUser(user));
    if (Array.isArray(payload.clients)) for (const client of payload.clients) await put(CLIENT_STORE, normalizeClient(client));
    if (Array.isArray(payload.projects)) for (const project of payload.projects) await put(PROJECT_STORE, normalizeProject(project));
    if (Array.isArray(payload.contacts)) for (const contact of payload.contacts) await put(CONTACT_STORE, normalizeContact(contact));
    if (Array.isArray(payload.contracts)) for (const contract of payload.contracts) await put(CONTRACT_STORE, normalizeContract(contract));
    if (Array.isArray(payload.containers)) for (const container of payload.containers) await put(CONTAINER_STORE, normalizeContainer(container));
    if (Array.isArray(payload.materials)) for (const material of payload.materials) await put(MATERIAL_STORE, normalizeMaterial(material));
    if (Array.isArray(payload.deliverySkuItems)) {
      for (const item of payload.deliverySkuItems) await put(DELIVERY_SKU_STORE, normalizeDeliverySkuItem(item));
    }
    if (Array.isArray(payload.trashRecords)) for (const item of payload.trashRecords) await put(TRASH_STORE, normalizeTrashRecord(item));
    if (Array.isArray(payload.settings)) for (const setting of payload.settings) await put(SETTINGS_STORE, setting);

    await loadAll();
    if (currentUser) currentUser = users.find((user) => user.id === currentUser.id) || currentUser;
    pushEntityAudit(
      "Backups",
      "imported",
      `${file.name} | units:${ensureArray(payload.units).length}, users:${ensureArray(payload.users).length}, clients:${ensureArray(payload.clients).length}, projects:${ensureArray(payload.projects).length}, contacts:${ensureArray(payload.contacts).length}, contracts:${ensureArray(payload.contracts).length}, containers:${ensureArray(payload.containers).length}, materials:${ensureArray(payload.materials).length}, delivery:${ensureArray(payload.deliverySkuItems).length}, trash:${ensureArray(payload.trashRecords).length}`,
      "backup"
    );
    render();
    queueAutoSync();
    return true;
  } catch {
    alert("Failed to import backup. Check the selected file.");
    return false;
  }
}

exportBtn.addEventListener("click", () => {
  triggerLocalBackupDownload("toolbar");
});

importInput.addEventListener("change", async () => {
  const file = importInput.files?.[0];
  if (!file) return;
  await importBackupFile(file);
  importInput.value = "";
});

syncConfigForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!can("sync")) return;
  if (hasPresetSyncConfig()) {
    updateSyncStatus("Cloud sync is centrally configured for all devices.", false);
    return;
  }

  const data = new FormData(syncConfigForm);
  const nextSupabaseUrl = data.get("supabaseUrl")?.toString().trim();
  const nextSupabaseAnonKey = data.get("supabaseAnonKey")?.toString().trim();
  const nextTenant = data.get("tenant")?.toString().trim();
  const syncTargetChanged =
    nextSupabaseUrl !== (syncConfig.supabaseUrl || "").trim() ||
    nextSupabaseAnonKey !== (syncConfig.supabaseAnonKey || "").trim() ||
    nextTenant !== (syncConfig.tenant || "").trim();

  syncConfig = {
    ...syncConfig,
    id: "syncConfig",
    supabaseUrl: nextSupabaseUrl,
    supabaseAnonKey: nextSupabaseAnonKey,
    tenant: nextTenant,
    autoSync: data.get("autoSync")?.toString() === "true",
    lastPullCursor: syncTargetChanged ? null : syncConfig.lastPullCursor || null,
    lastSyncAt: syncTargetChanged ? null : syncConfig.lastSyncAt || null,
  };

  await saveSetting(syncConfig);
  pushEntityAudit(
    "Cloud Sync",
    "updated",
    `target:${syncTargetChanged ? "changed" : "same"} | autoSync:${syncConfig.autoSync ? "on" : "off"} | tenant:${auditValue(syncConfig.tenant || "-")}`,
    "sync"
  );
  updateSyncStatus(t("Configuracao salva."));
  if (navigator.onLine) startAutoPullLoop();
  else stopAutoPullLoop();
});

pushSyncBtn.addEventListener("click", () => pushCloud());
pullSyncBtn.addEventListener("click", () => pullCloud());

function updateConnectivity() {
  const online = navigator.onLine;
  connectionBadge.textContent = online ? t("Online") : t("Offline");
  connectionBadge.style.borderColor = online ? "#94c5a8" : "#e7b6b6";
  connectionBadge.style.background = online ? "#f4fcf7" : "#fff5f5";

  if (online) {
    queueAutoSync();
    startAutoPullLoop();
  } else {
    stopAutoPullLoop();
  }
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

cameraScanCloseBtn?.addEventListener("click", () => closeCameraScanModal(""));
cameraScanManualBtn?.addEventListener("click", () => {
  const manual = window.prompt("Paste or type the QR code:");
  closeCameraScanModal(manual ? String(manual).trim() : "");
});
cameraScanPhotoBtn?.addEventListener("click", () => {
  cameraScanPhotoInput?.click();
});
cameraScanPhotoInput?.addEventListener("change", async () => {
  const file = Array.from(cameraScanPhotoInput.files || [])[0];
  if (!file) return;
  setCameraScanStatus("Reading photo...");
  try {
    const scanned = await decodeQrFromImageFile(file);
    if (scanned) {
      playScanFeedbackSound();
      closeCameraScanModal(scanned);
      return;
    }
    setCameraScanStatus("No QR code found in this image. Try another one.");
  } catch {
    setCameraScanStatus("Could not read this image. Try another one or enter manually.");
  } finally {
    cameraScanPhotoInput.value = "";
  }
});
cameraScanModal?.addEventListener("click", (event) => {
  if (event.target === cameraScanModal) closeCameraScanModal("");
});
document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  if (!cameraScanModal || cameraScanModal.classList.contains("hidden")) return;
  closeCameraScanModal("");
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
if (signupGenderSelect && !signupGenderSelect.value) signupGenderSelect.value = "unspecified";
if (userAdminGenderSelect && !userAdminGenderSelect.value) userAdminGenderSelect.value = "unspecified";
if (userEditGenderSelect && !userEditGenderSelect.value) userEditGenderSelect.value = "unspecified";
setAvatarPreview(userAdminPhotoPreview, defaultAvatarForGender(userAdminGenderSelect?.value || "unspecified"));
setAvatarPreview(userEditPhotoPreview, defaultAvatarForGender(userEditGenderSelect?.value || "unspecified"));
renderUserIdCard(null);
setUserAdminFormOpen(false);
resetQrFlowForm();

document.addEventListener(
  "input",
  (event) => {
    const form = formForSyncLock(event.target);
    if (!form) return;
    setFormSyncDirty(form, true);
  },
  true
);

document.addEventListener(
  "change",
  (event) => {
    const form = formForSyncLock(event.target);
    if (!form) return;
    setFormSyncDirty(form, true);
  },
  true
);

document.addEventListener(
  "reset",
  (event) => {
    const form = formForSyncLock(event.target);
    if (!form) return;
    setFormSyncDirty(form, false);
  },
  true
);

document.addEventListener(
  "submit",
  (event) => {
    const form = formForSyncLock(event.target);
    if (!form) return;
    blockAutoPullTemporarily();
  },
  true
);

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

clientsNavGroup?.addEventListener("mouseenter", () => {
  if (!window.matchMedia("(hover: hover)").matches) return;
  if (!currentUser || !can("manageCatalog")) return;
  clearTimeout(clientsNavHoverTimer);
  clientsNavHoverTimer = window.setTimeout(() => {
    if (!currentUser || !can("manageCatalog")) return;
    openClientsHub();
  }, 180);
});

clientsNavGroup?.addEventListener("mouseleave", () => {
  clearTimeout(clientsNavHoverTimer);
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
    await migrateLegacyUserRoleKey();
    await ensureDeveloperRolePresence();
    await pullCloud({ silent: true, force: true });
    await migrateLegacyUserRoleKey();
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
