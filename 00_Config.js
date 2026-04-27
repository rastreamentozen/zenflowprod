// ====================================================================================
// 🧠 ARQUIVO: 00_Config.gs (Configurações e Constantes Globais)
// ====================================================================================

const PLANILHA_ID = "1wcgYDTH7C9vRuu2CE43WMB1Rh0h9xf3JplTJWqQtsZA";
const ID_PLANILHA_TECNICOS = "1yrYwyE0iy4aYKHEthMxPF3rOUiygfQkSLqXBzIGtK4I";
const EMAIL_REMETENTE = "rastreamento@zenseguro.com";

const MAPA_COLUNAS = {
  DATA: 1, NOME: 2, PLACA: 3, CHASSI: 4, FIPE: 5, EMAIL: 6, TELEFONE: 7,
  CHECK_EMAIL: 8, DATA_EMAIL: 9, RESPONDEU_EMAIL: 10, CHECK_WHATS: 11, 
  DATA_WHATS: 12, RESPONDEU_WHATS: 13, RESPONSAVEL: 14, FIPE_BAIXA: 15, 
  TECNICO_INDISPONIVEL: 16, ESTADO: 17
};

const SGA_CONFIG = {
  URL_AUTH: "https://api.hinova.com.br/api/sga/v2/usuario/autenticar",
  URL_CONSULTA_BASE: "https://api.hinova.com.br/api/sga/v2/veiculo/buscar/",
  TOKEN_ASSOCIACAO: "041e6d561c08d16fce2a5beead2ca02fa4f4ee113d51a6f56e9c9e7c89694e4187b271c6cf997f4ece7bf4180e13ee750cb67ffbaa15d5bc82260c3b57a27cfc114322ab6e29a6dc368f1c2b4ea59702456d6c9df528c97f3386aee5978276f6",
  USUARIO: "victor rodrigues", SENHA: "ZEN0102"
};

const MAPA_SITUACAO_SGA = { "1": "Ativo", "2": "Inativo", "3": "Pendente", "4": "Inadimplente", "5": "Negado", "6": "Cancelado", "7": "Evento", "8": "Indenizado", "11": "Cancelado com rastreador", "12": "Inativos com rastreador", "13": "Inativos sem rastreador", "14": "Ativo com adesivo", "17": "Cancelamento pendente", "18": "Envio de termos", "19": "Desligado do corpo associativo", "22": "Aguardando indenização" };