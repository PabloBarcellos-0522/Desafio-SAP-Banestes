/**
 * SERVIDOR.GS - Arquivo Central de Backend
 * Refatorado para seguir padrões SRP e DRY.
 */

const CONFIG = {
  SHEETS: {
    PROCESSOS: "Processos",
    CLIENTES: "Clientes",
    PRODUTOS: "Produtos",
    UNIDADES: "Unidades",
  },
  COLUMNS: {
    ID: "ID",
    CLIENTE_ID: "CLIENTE ID",
    PRODUTO_ID: "PRODUTO ID",
    UNIDADE_ID: "UNIDADE ID",
    DATA_CRIACAO: "DATA DA CRIAÇÃO",
    DATA_ATUALIZACAO: "DATA DA ATUALIZAÇÃO",
    VALOR: "VALOR DO CONTRATO",
    CONTRATO: "CÓD. DO CONTRATO",
    ANDAMENTO: "ANDAMENTO",
    GARANTIA: "TIPO DE GARANTIA",
  },
  CACHE: {
    KEY_PROCESSOS: "cache_processos_ajuizamento",
    TTL_SECONDS: 900, // Expira em 15 minutos
  },
};

/**
 * Remove a chave de cache dos processos para forçar atualização na próxima leitura.
 * Chamada obrigatoriamente após qualquer operação de escrita (CRUD).
 */
function limparCacheProcessos() {
  const cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE.KEY_PROCESSOS);
}

/**
 * Retorna a instância da planilha ativa.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Busca dados brutos e cabeçalhos de uma aba específica.
 * @param {string} sheetName Nome da aba.
 * @returns {{headers: string[], data: any[][]}}
 */
function getSheetData(sheetName) {
  const sheet = getSS().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Aba '${sheetName}' não encontrada`);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return { headers, data };
}

/**
 * Gera mapas de tradução (De/Para) para as tabelas de referência.
 * @returns {Object} Mapas de Clientes, Produtos e Unidades.
 */
function getRefMaps() {
  const clientes = getSheetData(CONFIG.SHEETS.CLIENTES);
  const produtos = getSheetData(CONFIG.SHEETS.PRODUTOS);
  const unidades = getSheetData(CONFIG.SHEETS.UNIDADES);

  const idIdx = (h) => h.indexOf("ID");
  const nomeIdx = (h, name = "NOME") => h.indexOf(name);

  return {
    clientes: {
      idToNome: new Map(
        clientes.data.map((r) => [r[idIdx(clientes.headers)], r[nomeIdx(clientes.headers)]]),
      ),
      nomeToId: new Map(
        clientes.data.map((r) => [r[nomeIdx(clientes.headers)], r[idIdx(clientes.headers)]]),
      ),
    },
    produtos: {
      idToNome: new Map(
        produtos.data.map((r) => [r[idIdx(produtos.headers)], r[nomeIdx(produtos.headers)]]),
      ),
      nomeToId: new Map(
        produtos.data.map((r) => [r[nomeIdx(produtos.headers)], r[idIdx(produtos.headers)]]),
      ),
    },
    unidades: {
      idToNome: new Map(
        unidades.data.map((r) => [r[idIdx(unidades.headers)], r[nomeIdx(unidades.headers)]]),
      ),
      nomeToId: new Map(
        unidades.data.map((r) => [r[nomeIdx(unidades.headers)], r[idIdx(unidades.headers)]]),
      ),
    },
  };
}

function doGet() {
  // CRITICAL: Usa createTemplateFromFile para processar os <?!= include(...) ?>
  var template = HtmlService.createTemplateFromFile("Cliente");

  return template
    .evaluate()
    .setTitle("SAP - Banestes")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função para incluir arquivos HTML (CSS e JS) no template principal
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- BUSCA DE DADOS ---

/**
 * Função pública para busca de processos com suporte a Cache.
 * Melhora drasticamente a performance de carregamento para o usuário final.
 * @returns {Object[]}
 */
function getProcessos() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.CACHE.KEY_PROCESSOS);

  if (cached) {
    console.log("Retornando dados do Cache");
    return JSON.parse(cached);
  }

  // Se não houver cache, busca na planilha e armazena
  const dados = getProcessosFromSheet();

  try {
    // CacheService tem limite de 100KB por entrada.
    cache.put(CONFIG.CACHE.KEY_PROCESSOS, JSON.stringify(dados), CONFIG.CACHE.TTL_SECONDS);
  } catch (e) {
    console.warn("Falha ao gravar no cache (limite de tamanho): " + e.message);
  }

  return dados;
}

/**
 * Busca e formata os dados diretamente da planilha (lógica pesada).
 * @returns {Object[]}
 */
function getProcessosFromSheet() {
  const { headers, data } = getSheetData(CONFIG.SHEETS.PROCESSOS);
  const maps = getRefMaps();

  return data.map((row) => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      switch (h) {
        case CONFIG.COLUMNS.CLIENTE_ID:
          obj["CLIENTE"] = maps.clientes.idToNome.get(val) || val;
          break;
        case CONFIG.COLUMNS.PRODUTO_ID:
          obj["PRODUTO"] = maps.produtos.idToNome.get(val) || val;
          break;
        case CONFIG.COLUMNS.UNIDADE_ID:
          obj["UNIDADE"] = maps.unidades.idToNome.get(val) || val;
          break;

        case CONFIG.COLUMNS.DATA_CRIACAO:
        case CONFIG.COLUMNS.DATA_ATUALIZACAO:
          if (val instanceof Date) {
            obj[h] = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
            obj[h + "_SORT"] = val.getTime();
          } else {
            obj[h] = val;
            obj[h + "_SORT"] = 0;
          }
          break;

        case CONFIG.COLUMNS.VALOR:
          let num = 0;
          if (typeof val === "number") num = val;
          else if (typeof val === "string")
            num = parseFloat(val.replace(/[^\d,-]/g, "").replace(",", ".")) || 0;
          obj["VALOR_SORT"] = num;
          obj[h] =
            "R$ " +
            num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
          break;

        default:
          obj[h] = val;
      }
    });
    return obj;
  });
}

// --- FUNÇÕES DE APOIO ---

function getClientes() {
  return getSheetData(CONFIG.SHEETS.CLIENTES);
}
function getProdutos() {
  return getSheetData(CONFIG.SHEETS.PRODUTOS);
}
function getUnidades() {
  return getSheetData(CONFIG.SHEETS.UNIDADES);
}

function getClientesForDropdown() {
  return getClientes().data.map((r) => ({ name: r[1] }));
}
function getProdutosForDropdown() {
  return getProdutos().data.map((r) => ({ name: r[2] }));
}
function getUnidadesForDropdown() {
  return getUnidades().data.map((r) => ({ name: r[2] }));
}

// --- CRUD ---

/**
 * Gera o próximo ID disponível no padrão PROCESSO_X.
 * Garante que o ID não exista na tabela.
 * @returns {string} Próximo ID.
 */
function gerarProximoID() {
  const { headers, data } = getSheetData(CONFIG.SHEETS.PROCESSOS);
  const idIdx = headers.indexOf(CONFIG.COLUMNS.ID);
  const existingIds = new Set(data.map((r) => String(r[idIdx]).toUpperCase()));

  let counter = 1;
  while (existingIds.has(`PROCESSO_${counter}`)) {
    counter++;
  }
  return `PROCESSO_${counter}`;
}

/**
 * Gera um código de contrato numérico aleatório e único.
 * @returns {string} Código de contrato.
 */
function gerarCodigoContratoUnico() {
  const { headers, data } = getSheetData(CONFIG.SHEETS.PROCESSOS);
  const contratoIdx = headers.indexOf(CONFIG.COLUMNS.CONTRATO);
  const existingContratos = new Set(data.map((r) => String(r[contratoIdx]).replace(/\D/g, "")));

  let novoContrato;
  do {
    // Gera um número entre 10 e 12 dígitos
    novoContrato = Math.floor(Math.random() * 900000000000 + 10000000000).toString();
  } while (existingContratos.has(novoContrato));

  return novoContrato;
}

/**
 * Adiciona um novo processo à planilha com controle de concorrência e invalidação de cache.
 * ID e Código de Contrato seguem os novos padrões de nomenclatura e unicidade.
 */
function adicionarProcesso(processo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);

    const sheet = getSS().getSheetByName(CONFIG.SHEETS.PROCESSOS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const maps = getRefMaps();

    // Geração baseada nos novos padrões
    const novoID = gerarProximoID();
    const novoContrato = processo[CONFIG.COLUMNS.CONTRATO]
      ? processo[CONFIG.COLUMNS.CONTRATO].trim()
      : gerarCodigoContratoUnico();

    const newRow = headers.map((h) => {
      switch (h) {
        case CONFIG.COLUMNS.ID:
          return novoID;
        case CONFIG.COLUMNS.CLIENTE_ID:
          return maps.clientes.nomeToId.get(processo.Cliente) || "";
        case CONFIG.COLUMNS.PRODUTO_ID:
          return maps.produtos.nomeToId.get(processo.Produto) || "";
        case CONFIG.COLUMNS.UNIDADE_ID:
          return maps.unidades.nomeToId.get(processo.Unidade) || "";
        case CONFIG.COLUMNS.DATA_CRIACAO:
        case CONFIG.COLUMNS.DATA_ATUALIZACAO:
          return new Date();
        case CONFIG.COLUMNS.CONTRATO:
          return "'" + novoContrato;
        case CONFIG.COLUMNS.VALOR:
          return formatarParaTextoMoeda(processo[h]);
        default:
          return processo[h] || "";
      }
    });

    sheet.appendRow(newRow);
    limparCacheProcessos();
    return `Processo (${novoID}) criado com sucesso!`;
  } catch (e) {
    throw new Error("Erro ao adicionar processo: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Edita um processo existente com controle de concorrência e invalidação de cache.
 * @param {Object} processoData Dados atualizados do processo.
 * @returns {string} Mensagem de sucesso.
 */
function editarProcesso(processoData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);

    const sheet = getSS().getSheetByName(CONFIG.SHEETS.PROCESSOS);
    const { headers, data } = getSheetData(CONFIG.SHEETS.PROCESSOS);
    const idIdx = headers.indexOf(CONFIG.COLUMNS.ID);

    // Usa o ID fixo enviado pelo front (que não muda na edição)
    const targetID = processoData.editProcessoID;

    let rIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][idIdx] == targetID) {
        rIdx = i;
        break;
      }
    }
    if (rIdx === -1) throw new Error(`Processo (${targetID}) não encontrado`);

    const maps = getRefMaps();

    const updatedRow = headers.map((h) => {
      switch (h) {
        case CONFIG.COLUMNS.ID:
          return targetID;
        case CONFIG.COLUMNS.CLIENTE_ID:
          return maps.clientes.nomeToId.get(processoData.Cliente) || processoData.Cliente;
        case CONFIG.COLUMNS.PRODUTO_ID:
          return maps.produtos.nomeToId.get(processoData.Produto) || processoData.Produto;
        case CONFIG.COLUMNS.UNIDADE_ID:
          return maps.unidades.nomeToId.get(processoData.Unidade) || processoData.Unidade;
        case CONFIG.COLUMNS.DATA_CRIACAO:
          return data[rIdx][headers.indexOf(h)] || new Date();
        case CONFIG.COLUMNS.DATA_ATUALIZACAO:
          return new Date();
        case CONFIG.COLUMNS.CONTRATO:
          return "'" + (processoData[h] || "");
        case CONFIG.COLUMNS.VALOR:
          return formatarParaTextoMoeda(processoData[h]);
        default:
          return processoData[h] !== undefined ? processoData[h] : data[rIdx][headers.indexOf(h)];
      }
    });

    sheet.getRange(rIdx + 2, 1, 1, updatedRow.length).setValues([updatedRow]);
    limparCacheProcessos();
    return `Processo (${targetID}) atualizado com sucesso!`;
  } catch (e) {
    throw new Error("Erro ao editar processo: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Exclui um processo com controle de concorrência e invalidação de cache.
 * @param {string|number} id ID do processo.
 * @returns {string} Mensagem de sucesso.
 */
function excluirProcesso(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);

    const sheet = getSS().getSheetByName(CONFIG.SHEETS.PROCESSOS);
    const { headers, data } = getSheetData(CONFIG.SHEETS.PROCESSOS);
    const idIdx = headers.indexOf(CONFIG.COLUMNS.ID);

    for (let i = 0; i < data.length; i++) {
      if (data[i][idIdx] == id) {
        sheet.deleteRow(i + 2);
        limparCacheProcessos();
        return "Processo excluído com sucesso!";
      }
    }
    throw new Error("Processo não encontrado para exclusão");
  } catch (e) {
    throw new Error("Erro ao excluir processo: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function formatarParaTextoMoeda(valor) {
  if (!valor) return "'R$ 0,00";
  let num =
    typeof valor === "number" ? valor : parseFloat(valor.toString().replace(/[^\d]/g, "")) / 100;
  if (isNaN(num)) return "'R$ 0,00";
  return (
    "'R$ " + num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  );
}
