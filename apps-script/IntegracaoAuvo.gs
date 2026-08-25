/**
 * INTEGRACAO AUVO -> GOOGLE SHEETS
 * Versao 6.0.5 completa e consolidada para clientes e questionarios.
 * Inclui alerta automatico de clientes sem grupo por e-mail e WhatsApp.
 * Inclui diagnostico e integracao mensal da BASE_QUESTIONARIO.
 * Inclui selecao, vigencia, carga mensal e atualizacao incremental dos questionarios.
 *
 * ALTERACOES NESTA VERSAO (revisao 2026-08-24), em relacao a 6.0.3:
 * 1. iniciarCargaQuestionariosAuvo_ agora verifica QUESTIONARIOS_CARGA_ATIVA antes de
 *    comecar uma nova carga, assim como atualizarBaseAuvo ja fazia para AUVO_ATUALIZACAO_ATIVA.
 *    Sem essa checagem, o gatilho diario das 23h podia interromper uma carga mensal em
 *    andamento, limpando a aba temporaria e sobrescrevendo o estado no meio do processo.
 * 2. processarCargaQuestionariosAuvo_ agora reagenda a continuacao (agendarContinuacaoQuestionariosAuvo_)
 *    quando o lock nao e adquirido, em vez de apenas retornar. Sem isso, uma disputa de lock
 *    passageira interrompia silenciosamente a cadeia de continuacao de uma carga em andamento.
 * 3. A formula "Chave Tec Data" (coluna 36 de BASE_QUESTIONARIO) referenciava a coluna 26
 *    ("CPF/CNPJ do Cliente", texto) em vez da coluna 25 ("Data da tarefa"), corrigido.
 * 4. Nova funcao listarTarefasComRetentativaQuestionariosAuvo_: se a rota de tarefas
 *    falhar num dia especifico (ex.: HTTP 404 pontual do AUVO), o script redescobre o
 *    estilo de rota e tenta esse dia mais uma vez antes de abortar a carga inteira do
 *    periodo. Corrigido apos falha real observada em producao (carga mensal de 01/08 a
 *    24/08 abortou no dia 02/08 por um HTTP 404 transitorio em /Tasks).
 *
 * REQUISITO PARA A PRIMEIRA EXECUCAO:
 * - Deve existir a aba BACKUP_BASE_AUVO com a exportacao anterior do AUVO.
 * - O backup e usado para preservar campos que o GET /customers nao entrega.
 *
 * FLUXO:
 * 1. Configure as credenciais pelo menu.
 * 2. Teste a conexao.
 * 3. Teste os catalogos.
 * 4. Execute Atualizar BASE_AUVO uma unica vez.
 * 5. Depois do SUCESSO, instale a atualizacao diaria.
 */

const AUVO = {
  VERSAO: '6.0.5-2026-08-24',
  FUSO_HORARIO: 'America/Sao_Paulo',
  BASE_URL: 'https://api.auvo.com.br/v2',
  ABA_CLIENTES: 'BASE_AUVO',
  ABA_BACKUP: 'BACKUP_BASE_AUVO',
  ABA_LOG: 'LOG_API',
  ABA_TEMP: 'TEMP_AUVO_API',
  ABA_QUESTIONARIOS: 'BASE_QUESTIONARIO',
  ABA_DIAGNOSTICO_QUESTIONARIOS: 'DIAGNOSTICO_QUESTIONARIOS',
  ABA_CONFIG_QUESTIONARIOS: 'CONFIG_QUESTIONARIOS_AUVO',
  ABA_TEMP_QUESTIONARIOS: 'TEMP_QUESTIONARIOS_AUVO',
  TOTAL_COLUNAS: 34,
  PAGE_SIZE: 100,
  MAX_PAGINAS_POR_EXECUCAO: 5,
  TEMPO_MAXIMO_EXECUCAO: 240000,
  CACHE_CATALOGOS_SEGUNDOS: 21600
};

const CABECALHOS_BASE_QUESTIONARIO_ = [
  'Código Tarefa', 'Colaborador', 'Cliente', 'Código do cliente',
  'Qual o modelo do equipamento?', 'Código do patrimônio do equipamento',
  'Foto do patrimonio (QR Code)', 'Foto do ambiente e equipamento antes',
  'Local da instalacao', 'Foto do refil antes',
  'Qual o volume do refil antes da reposiçao?',
  'Foi identificado volume de fragrância acima do esperado para o período de consumo. Qual o motivo provável?',
  'Serviços Executados', 'Foto da Pilha',
  'Foto da configuração (potência/programação)',
  'Situação da trava do equipamento', 'Foto(s) Frasco Lote/Fragrância',
  'Foto(s) do refil depois e Limpeza do equipamento', 'Observações',
  'Data resposta', 'Hora resposta', 'Equipamento', 'Identificador',
  'Código externo', 'Data da tarefa', 'CPF/CNPJ do Cliente',
  'Coluna1', 'Coluna2', 'Coluna3', 'Coluna4', 'Coluna5', 'Coluna6',
  'Filtro Painel', 'Codigo Cliente Texto', 'Chave Tec Cliente',
  'Chave Tec Data', 'Cliente Único Válido',
  'ID Questionário AUVO', 'Nome Questionário AUVO', 'Chave Integração AUVO',
  'Contabiliza Produtividade?', 'Código Cliente AUVO', 'Assinatura da Instância'
];

const CABECALHOS_BASE_AUVO_ = [
  'Código',
  'Código externo',
  'cliente',
  'CPF ou CNPJ',
  'Razão social',
  'Endereço',
  'Complemento do Endereço',
  'Telefone corporativo',
  'E-mail corporativo',
  'Falar com',
  'Usuário Auvo responsável',
  'Grupos',
  'Segmento',
  'Observação',
  'Observação Interna',
  'Latitude',
  'Longitude',
  'Status',
  'Anotação',
  'Equipes responsáveis',
  'Última visita',
  'Data de cadastro',
  'Usuario cadastro',
  'Contribuinte do ICMS',
  'Inscrição Estadual',
  'Inscrição Municipal',
  'E-mail de cobrança',
  'CEP de cobrança',
  'Logradouro endereço de cobrança',
  'Número endereço de cobrança',
  'Complemento endereço de cobrança',
  'Bairro de cobrança',
  'Cidade de cobrança',
  'Estado de cobrança'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Integração AUVO')
    .addItem('1. Configurar credenciais', 'configurarCredenciaisAuvo')
    .addItem('2. Testar conexão', 'testarConexaoAuvo')
    .addItem('3. Testar catálogos', 'testarCatalogosAuvo')
    .addItem('3A. Verificar versão instalada', 'verificarVersaoIntegracaoAuvo')
    .addItem('4. Atualizar BASE_AUVO', 'atualizarBaseAuvo')
    .addSeparator()
    .addItem('5. Configurar destinatários dos alertas', 'configurarDestinatariosAlertasAuvo')
    .addItem('5A. Autorizar envio de e-mail', 'autorizarEnvioEmailAuvo')
    .addItem('6. Configurar WhatsApp oficial (Meta)', 'configurarWhatsAppMetaAuvo')
    .addItem('7. Testar alerta de clientes sem grupo', 'testarAlertaClientesSemGrupoAuvo')
    .addSeparator()
    .addItem('8. Diagnosticar questionários', 'diagnosticarQuestionariosAuvo')
    .addItem('9. Sincronizar e selecionar questionários', 'sincronizarConfiguracaoQuestionariosAuvo')
    .addItem('10. Validar questionários contabilizados', 'validarConfiguracaoQuestionariosAuvo')
    .addItem('11. Configurar período mensal', 'configurarPeriodoQuestionariosAuvo')
    .addItem('12. Testar rota de tarefas do período', 'testarRotaTarefasQuestionariosAuvo')
    .addItem('13. Carregar mês na BASE_QUESTIONARIO', 'carregarMesQuestionariosAuvo')
    .addItem('14. Atualizar hoje e ontem', 'atualizarQuestionariosRecentesAuvo')
    .addItem('15. Instalar atualização diária dos questionários', 'instalarAtualizacaoDiariaQuestionariosAuvo')
    .addItem('16. Remover atualização diária dos questionários', 'removerAtualizacaoDiariaQuestionariosAuvo')
    .addSeparator()
    .addItem('Cancelar atualização em andamento', 'cancelarAtualizacaoAuvo')
    .addItem('Instalar atualização diária', 'instalarAtualizacaoDiaria')
    .addItem('Remover atualização diária', 'removerAtualizacaoDiaria')
    .addToUi();
}

function verificarVersaoIntegracaoAuvo() {
  const mensagem = 'Versão instalada: ' + AUVO.VERSAO;
  registrarLogAuvo_('VERSÃO', 'OK', mensagem);
  SpreadsheetApp.getUi().alert('Integração AUVO', mensagem, SpreadsheetApp.getUi().ButtonSet.OK);
}

function configurarCredenciaisAuvo() {
  const ui = SpreadsheetApp.getUi();

  const respostaKey = ui.prompt(
    'Credenciais do AUVO',
    'Cole a apiKey:',
    ui.ButtonSet.OK_CANCEL
  );

  if (respostaKey.getSelectedButton() !== ui.Button.OK) return;

  const respostaToken = ui.prompt(
    'Credenciais do AUVO',
    'Cole o apiToken:',
    ui.ButtonSet.OK_CANCEL
  );

  if (respostaToken.getSelectedButton() !== ui.Button.OK) return;

  const apiKey = respostaKey.getResponseText().trim();
  const apiToken = respostaToken.getResponseText().trim();

  if (!apiKey || !apiToken) {
    ui.alert('A apiKey e o apiToken são obrigatórios.');
    return;
  }

  PropertiesService.getScriptProperties().setProperties({
    AUVO_API_KEY: apiKey,
    AUVO_API_TOKEN: apiToken
  });

  ui.alert('Credenciais armazenadas com sucesso.');
}

function testarConexaoAuvo() {
  try {
    autenticarAuvo_();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Conexão realizada com sucesso.',
      'Integração AUVO',
      8
    );
  } catch (erro) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Falha na conexão: ' + erro.message,
      'Integração AUVO',
      12
    );
    throw erro;
  }
}

/**
 * Inicia uma carga nova. Se ja existir carga ativa, nao reinicia nem apaga o
 * temporario. Use o cancelamento explicito antes de recomecar.
 */
function atualizarBaseAuvo() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Outra execução da integração AUVO está em andamento.');
  }

  try {
    const propriedades = PropertiesService.getScriptProperties();

    if (propriedades.getProperty('AUVO_ATUALIZACAO_ATIVA') === 'SIM') {
      registrarLogAuvo_(
        'CLIENTES',
        'AVISO',
        'Nova solicitação ignorada: já existe atualização em andamento.'
      );
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Já existe uma atualização em andamento.',
        'Integração AUVO',
        10
      );
      return;
    }

    validarFonteHistoricoAuvo_();
    removerGatilhosContinuacaoAuvo_();

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    let temporaria = planilha.getSheetByName(AUVO.ABA_TEMP);

    if (!temporaria) {
      temporaria = planilha.insertSheet(AUVO.ABA_TEMP);
    }

    temporaria.clearContents();
    temporaria.hideSheet();

    propriedades.setProperties({
      AUVO_PAGINA_ATUAL: '1',
      AUVO_TOTAL_IMPORTADO: '0',
      AUVO_ATUALIZACAO_ATIVA: 'SIM',
      AUVO_INICIO_ATUALIZACAO: new Date().toISOString()
    });

    registrarLogAuvo_('CLIENTES', 'INÍCIO', 'Nova atualização iniciada.');
  } finally {
    lock.releaseLock();
  }

  processarPaginasAuvo_();
}

function continuarAtualizacaoBaseAuvo() {
  processarPaginasAuvo_();
}

function processarPaginasAuvo_() {
  const lock = LockService.getScriptLock();
  const inicioExecucao = Date.now();

  if (!lock.tryLock(30000)) {
    agendarContinuacaoAuvo_();
    return;
  }

  try {
    const propriedades = PropertiesService.getScriptProperties();

    if (propriedades.getProperty('AUVO_ATUALIZACAO_ATIVA') !== 'SIM') {
      removerGatilhosContinuacaoAuvo_();
      return;
    }

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const temporaria = planilha.getSheetByName(AUVO.ABA_TEMP);

    if (!temporaria) {
      throw new Error('A aba temporária da integração não foi encontrada.');
    }

    const token = autenticarAuvo_();
    const catalogos = carregarCatalogosAuvo_(token);

    let pagina = Number(propriedades.getProperty('AUVO_PAGINA_ATUAL') || 1);
    let totalImportado = Number(
      propriedades.getProperty('AUVO_TOTAL_IMPORTADO') || 0
    );
    let paginasProcessadas = 0;

    while (
      paginasProcessadas < AUVO.MAX_PAGINAS_POR_EXECUCAO &&
      Date.now() - inicioExecucao < AUVO.TEMPO_MAXIMO_EXECUCAO
    ) {
      const url =
        AUVO.BASE_URL +
        '/customers?page=' + pagina +
        '&pageSize=' + AUVO.PAGE_SIZE +
        '&order=asc';

      const resposta = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: {
          Authorization: 'Bearer ' + token,
          Accept: 'application/json'
        },
        muteHttpExceptions: true
      });

      validarHttpAuvo_(resposta, 'consulta da página ' + pagina);

      const json = JSON.parse(resposta.getContentText());
      const resultado = json.result || {};
      const clientes = resultado.entityList || [];

      if (!Array.isArray(clientes)) {
        throw new Error(
          'A página ' + pagina + ' não retornou result.entityList.'
        );
      }

      if (clientes.length > 0) {
        const linhas = clientes.map(function(cliente) {
          return mapearClienteAuvo_(cliente, catalogos);
        });

        temporaria
          .getRange(
            totalImportado + 1,
            1,
            linhas.length,
            AUVO.TOTAL_COLUNAS
          )
          .setValues(linhas);

        totalImportado += linhas.length;
      }

      const informacaoPagina =
        resultado.pageInformation ||
        resultado.pagedSearchReturnData ||
        {};

      const totalPaginas = Number(informacaoPagina.totalPages || 0);
      const ultimaPagina =
        clientes.length === 0 ||
        clientes.length < AUVO.PAGE_SIZE ||
        (totalPaginas > 0 && pagina >= totalPaginas);

      pagina++;
      paginasProcessadas++;

      propriedades.setProperties({
        AUVO_PAGINA_ATUAL: String(pagina),
        AUVO_TOTAL_IMPORTADO: String(totalImportado)
      });

      if (ultimaPagina) {
        finalizarAtualizacaoAuvo_(temporaria, totalImportado);
        return;
      }
    }

    SpreadsheetApp.flush();

    registrarLogAuvo_(
      'CLIENTES',
      'CONTINUAÇÃO',
      totalImportado +
        ' clientes armazenados. Próxima página: ' +
        pagina + '.'
    );

    agendarContinuacaoAuvo_();
  } catch (erro) {
    removerGatilhosContinuacaoAuvo_();
    encerrarEstadoAtualizacaoAuvo_();
    registrarLogAuvo_('CLIENTES', 'ERRO', erro.message);
    throw erro;
  } finally {
    lock.releaseLock();
  }
}

function finalizarAtualizacaoAuvo_(temporaria, totalImportado) {
  if (totalImportado <= 0) {
    throw new Error(
      'Nenhum cliente foi importado. A BASE_AUVO foi preservada.'
    );
  }

  const linhas = temporaria
    .getRange(1, 1, totalImportado, AUVO.TOTAL_COLUNAS)
    .getValues();

  const linhasMescladas = mesclarCamposHistoricosAuvo_(linhas);
  gravarBaseAuvo_(linhasMescladas);

  const inicioIso = PropertiesService.getScriptProperties().getProperty(
    'AUVO_INICIO_ATUALIZACAO'
  );
  let duracao = '';

  if (inicioIso) {
    const milissegundos = Date.now() - new Date(inicioIso).getTime();
    if (!isNaN(milissegundos)) {
      duracao = ' Duração: ' + formatarDuracaoAuvo_(milissegundos) + '.';
    }
  }

  encerrarEstadoAtualizacaoAuvo_();
  removerGatilhosContinuacaoAuvo_();
  temporaria.clearContents();

  registrarLogAuvo_(
    'CLIENTES',
    'SUCESSO',
    totalImportado + ' clientes atualizados.' + duracao
  );

  // O alerta e um processo independente. Uma falha de comunicacao nao
  // desfaz nem altera a atualizacao da BASE_AUVO que ja foi concluida.
  try {
    dispararAlertaClientesSemGrupoAuvo_(linhasMescladas, false);
  } catch (erroAlerta) {
    registrarLogAuvo_(
      'ALERTA SEM GRUPO',
      'ERRO',
      erroAlerta.message
    );
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    totalImportado + ' clientes atualizados com sucesso.',
    'Integração AUVO',
    10
  );
}

function cancelarAtualizacaoAuvo() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('A execução atual ainda não liberou o bloqueio.');
  }

  try {
    removerGatilhosContinuacaoAuvo_();
    encerrarEstadoAtualizacaoAuvo_();
    registrarLogAuvo_(
      'CLIENTES',
      'CANCELADO',
      'Atualização cancelada. A BASE_AUVO foi preservada.'
    );
  } finally {
    lock.releaseLock();
  }

  SpreadsheetApp.getUi().alert(
    'Atualização cancelada. A BASE_AUVO não foi alterada.'
  );
}

function encerrarEstadoAtualizacaoAuvo_() {
  const propriedades = PropertiesService.getScriptProperties();
  [
    'AUVO_PAGINA_ATUAL',
    'AUVO_TOTAL_IMPORTADO',
    'AUVO_ATUALIZACAO_ATIVA',
    'AUVO_INICIO_ATUALIZACAO'
  ].forEach(function(chave) {
    propriedades.deleteProperty(chave);
  });
}

function agendarContinuacaoAuvo_() {
  removerGatilhosContinuacaoAuvo_();

  ScriptApp.newTrigger('continuarAtualizacaoBaseAuvo')
    .timeBased()
    .after(60000)
    .create();
}

function removerGatilhosContinuacaoAuvo_() {
  ScriptApp.getProjectTriggers().forEach(function(gatilho) {
    if (gatilho.getHandlerFunction() === 'continuarAtualizacaoBaseAuvo') {
      ScriptApp.deleteTrigger(gatilho);
    }
  });
}

/**
 * Mapeamento confirmado pelo diagnostico real do GET /customers.
 */
function mapearClienteAuvo_(cliente, catalogos) {
  return [
    valorAuvo_(cliente.id),
    valorAuvo_(cliente.externalId),
    valorAuvo_(cliente.description),
    valorAuvo_(cliente.cpfCnpj),
    valorAuvo_(cliente.legalName),
    valorAuvo_(cliente.address),
    valorAuvo_(cliente.adressComplement),
    juntarListaAuvo_(cliente.phoneNumber),
    juntarListaAuvo_(cliente.email),
    valorAuvo_(cliente.manager),
    resolverIdsAuvo_(cliente.managersId, catalogos.usuarios),
    resolverIdsAuvo_(cliente.groupsId, catalogos.grupos),
    resolverIdAuvo_(cliente.segmentId, catalogos.segmentos),
    valorAuvo_(cliente.note),
    '',
    valorAuvo_(cliente.latitude),
    valorAuvo_(cliente.longitude),
    normalizarStatusAuvo_(cliente.active),
    '',
    resolverIdsAuvo_(cliente.managerTeamsId, catalogos.equipes),
    '',
    converterDataAuvo_(cliente.creationDate),
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    ''
  ];
}

/**
 * Catalogos que convertem os IDs do cliente em nomes legiveis.
 * Se um catalogo falhar, a importacao continua e usa o historico como fallback.
 */
function carregarCatalogosAuvo_(token) {
  return {
    usuarios: obterMapaCatalogoAuvo_(token, 'USUARIOS', ['/users/']),
    grupos: obterMapaCatalogoAuvo_(
      token,
      'GRUPOS',
      [
        '/customerGroups/',
        '/customersGroups/',
        '/customer-groups/',
        '/groups/'
      ]
    ),
    segmentos: obterMapaCatalogoAuvo_(token, 'SEGMENTOS', ['/segments/']),
    equipes: obterMapaCatalogoAuvo_(token, 'EQUIPES', ['/teams/'])
  };
}

function obterMapaCatalogoAuvo_(token, chave, endpoints) {
  const cache = CacheService.getScriptCache();
  const chaveCache = 'AUVO_CATALOGO_' + chave;
  const armazenado = cache.get(chaveCache);

  if (armazenado) {
    try {
      return JSON.parse(armazenado);
    } catch (erroCache) {
      cache.remove(chaveCache);
    }
  }

  const erros = [];

  for (let i = 0; i < endpoints.length; i++) {
    try {
      const mapa = consultarCatalogoAuvo_(token, endpoints[i]);

      if (Object.keys(mapa).length === 0) {
        throw new Error(endpoints[i] + ' retornou zero itens reconhecidos');
      }

      try {
        cache.put(
          chaveCache,
          JSON.stringify(mapa),
          AUVO.CACHE_CATALOGOS_SEGUNDOS
        );
      } catch (erroGravacaoCache) {
        registrarLogAuvo_(
          'CATÁLOGO ' + chave,
          'AVISO',
          'Catálogo carregado, mas não coube no cache.'
        );
      }

      return mapa;
    } catch (erro) {
      erros.push(erro.message);
    }
  }

  registrarLogAuvo_(
    'CATÁLOGO ' + chave,
    'AVISO',
    'Não foi possível resolver nomes. Tentativas: ' + erros.join(' | ')
  );

  return {};
}

function consultarCatalogoAuvo_(token, endpoint) {
  const estilos = endpoint === '/users/'
    ? ['MAIUSCULO', 'MINUSCULO']
    : ['MINUSCULO', 'MAIUSCULO'];
  const tentativas = [];
  let melhorMapa = {};

  for (let i = 0; i < estilos.length; i++) {
    try {
      const resultado = consultarCatalogoPaginadoAuvo_(
        token,
        endpoint,
        estilos[i]
      );

      if (Object.keys(resultado.mapa).length > Object.keys(melhorMapa).length) {
        melhorMapa = resultado.mapa;
      }

      if (resultado.completo && Object.keys(resultado.mapa).length > 0) {
        return resultado.mapa;
      }

      tentativas.push(
        estilos[i] + ': ' + resultado.motivo +
        ' (' + Object.keys(resultado.mapa).length + ' nomes)'
      );
    } catch (erro) {
      tentativas.push(estilos[i] + ': ' + erro.message);
    }
  }

  if (Object.keys(melhorMapa).length > 0) {
    throw new Error(
      endpoint + ' retornou dados, mas a paginação não pôde ser validada. ' +
      tentativas.join(' | ')
    );
  }

  throw new Error(
    endpoint + ' não retornou itens reconhecidos. ' + tentativas.join(' | ')
  );
}

/**
 * O AUVO usa maiúsculas em algumas rotas e minúsculas em outras. Esta função
 * testa um estilo por vez. Algumas rotas, como /customerGroups/, ignoram os
 * parâmetros de paginação e devolvem mais itens do que o pageSize solicitado.
 * Nesse caso específico, uma segunda resposta idêntica confirma que a rota
 * trabalha como catálogo completo de página única.
 */
function consultarCatalogoPaginadoAuvo_(token, endpoint, estilo) {
  const mapa = {};
  let pagina = 1;
  let assinaturaAnterior = '';
  const tamanhoPagina = 100;

  while (pagina <= 100) {
    const separador = endpoint.indexOf('?') >= 0 ? '&' : '?';
    const parametros = estilo === 'MAIUSCULO'
      ? 'Page=' + pagina + '&PageSize=' + tamanhoPagina + '&Order=Asc'
      : 'page=' + pagina + '&pageSize=' + tamanhoPagina + '&order=asc';
    const url =
      AUVO.BASE_URL + endpoint + separador +
      'paramFilter=' + encodeURIComponent('{}') +
      '&' + parametros;

    const resposta = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept: 'application/json'
      },
      muteHttpExceptions: true
    });

    validarHttpAuvo_(resposta, 'catálogo ' + endpoint);

    const json = JSON.parse(resposta.getContentText());
    const resultado = json.result !== undefined ? json.result : json;
    const entidades = extrairListaCatalogoAuvo_(resultado);

    if (entidades.length === 0) {
      return {
        mapa: mapa,
        completo: Object.keys(mapa).length > 0,
        motivo: Object.keys(mapa).length > 0
          ? 'fim da listagem'
          : 'zero entidades reconhecidas'
      };
    }

    const assinaturaAtual = criarAssinaturaPaginaCatalogoAuvo_(entidades);
    if (pagina > 1 && assinaturaAtual === assinaturaAnterior) {
      const quantidadeMapeada = Object.keys(mapa).length;
      const respostaIgnorouPageSize =
        entidades.length > tamanhoPagina &&
        quantidadeMapeada >= tamanhoPagina;

      if (respostaIgnorouPageSize) {
        return {
          mapa: mapa,
          completo: true,
          motivo:
            'catálogo completo em resposta única; a rota ignorou o pageSize'
        };
      }

      return {
        mapa: mapa,
        completo: false,
        motivo: 'a API repetiu a página anterior'
      };
    }
    assinaturaAnterior = assinaturaAtual;

    entidades.forEach(function(entidade) {
      const id = extrairIdCatalogoAuvo_(entidade);
      const nome = extrairNomeCatalogoAuvo_(entidade);
      if (temValorAuvo_(id) && temValorAuvo_(nome)) {
        mapa[String(id)] = nome;
      }
    });

    const info = extrairPaginacaoCatalogoAuvo_(resultado);

    const totalPaginas = Number(info.totalPages || 0);
    const totalItens = Number(info.totalItems || 0);
    const tamanhoInformado = Number(info.pageSize || tamanhoPagina);
    const terminou =
      entidades.length < tamanhoPagina ||
      info.hasMore === false ||
      (totalPaginas > 0 && pagina >= totalPaginas) ||
      (totalItens > 0 && pagina * tamanhoInformado >= totalItens);

    if (terminou) {
      return {
        mapa: mapa,
        completo: true,
        motivo: 'paginação concluída'
      };
    }
    pagina++;
  }

  return {
    mapa: mapa,
    completo: false,
    motivo: 'catálogo excedeu 100 páginas'
  };
}

function criarAssinaturaPaginaCatalogoAuvo_(entidades) {
  return entidades.map(function(entidade, indice) {
    const id = extrairIdCatalogoAuvo_(entidade);
    return temValorAuvo_(id) ? String(id) : 'SEM_ID_' + indice;
  }).join('|');
}

function extrairListaCatalogoAuvo_(resultado) {
  if (Array.isArray(resultado)) return resultado;
  if (!resultado || typeof resultado !== 'object') return [];

  const candidatos = [
    resultado.entityList,
    resultado.items,
    resultado.entities,
    resultado.users,
    resultado.teams,
    resultado.segments,
    resultado.groups,
    resultado.customerGroups,
    resultado.customersGroups,
    resultado.pagedSearchReturnData &&
      resultado.pagedSearchReturnData.entityList,
    resultado.pagedSearchReturnData &&
      resultado.pagedSearchReturnData.items,
    resultado.pageInformation &&
      resultado.pageInformation.entityList
  ];

  for (let i = 0; i < candidatos.length; i++) {
    if (Array.isArray(candidatos[i])) return candidatos[i];
  }

  return localizarListaEntidadesCatalogoAuvo_(resultado, 0);
}

/**
 * Algumas rotas da API envolvem a lista em objetos diferentes. Esta busca
 * limitada encontra a primeira lista de entidades sem depender do nome do
 * contêiner e sem percorrer indefinidamente a resposta.
 */
function localizarListaEntidadesCatalogoAuvo_(valor, profundidade) {
  if (!valor || typeof valor !== 'object' || profundidade > 4) return [];

  if (Array.isArray(valor)) {
    const pareceListaEntidades = valor.some(function(item) {
      return item &&
        typeof item === 'object' &&
        temValorAuvo_(extrairIdCatalogoAuvo_(item));
    });
    return pareceListaEntidades ? valor : [];
  }

  const chaves = Object.keys(valor);
  for (let i = 0; i < chaves.length; i++) {
    const encontrada = localizarListaEntidadesCatalogoAuvo_(
      valor[chaves[i]],
      profundidade + 1
    );
    if (encontrada.length) return encontrada;
  }

  return [];
}

function extrairPaginacaoCatalogoAuvo_(resultado) {
  if (!resultado || typeof resultado !== 'object') return {};

  const candidatos = [
    resultado,
    resultado.pageInformation,
    resultado.pagedSearchReturnData,
    resultado.pagination,
    resultado.paging
  ];

  for (let i = 0; i < candidatos.length; i++) {
    const item = candidatos[i];
    if (
      item &&
      typeof item === 'object' &&
      (
        temValorAuvo_(obterCampoCatalogoAuvo_(item, [
          'totalPages', 'pageCount', 'totalPage', 'totalItems',
          'totalCount', 'pageSize', 'hasMore', 'hasNextPage'
        ]))
      )
    ) {
      return {
        totalPages: obterCampoCatalogoAuvo_(
          item,
          ['totalPages', 'pageCount', 'totalPage']
        ),
        totalItems: obterCampoCatalogoAuvo_(
          item,
          ['totalItems', 'totalCount', 'count']
        ),
        pageSize: obterCampoCatalogoAuvo_(
          item,
          ['pageSize', 'itemsPerPage', 'perPage']
        ),
        hasMore: normalizarBooleanoCatalogoAuvo_(
          obterCampoCatalogoAuvo_(item, ['hasMore', 'hasNextPage'])
        )
      };
    }
  }

  return {};
}

function extrairIdCatalogoAuvo_(entidade) {
  if (!entidade || typeof entidade !== 'object') return '';

  return obterCampoCatalogoAuvo_(entidade, [
    'id', 'taskId', 'taskID', 'userId', 'userID', 'teamId', 'teamID', 'segmentId', 'segmentID',
    'groupId', 'groupID', 'customerGroupId', 'customerGroupID',
    'customerGroupsId', 'customerGroupsID', 'auvoUserId', 'auvoUserID',
    'questionnaireId', 'questionnaireID', 'questionaryId', 'questionaryID',
    'code', 'publicId', 'publicID'
  ]);
}

function extrairNomeCatalogoAuvo_(entidade) {
  if (!entidade || typeof entidade !== 'object') return '';

  const valor = obterCampoCatalogoAuvo_(entidade, [
    'description', 'name', 'userName', 'fullName', 'groupName', 'teamName',
    'segmentName', 'questionnaireName', 'questionaryName', 'title',
    'legalName', 'login', 'email'
  ]);

  return temValorAuvo_(valor) ? String(valor).trim() : '';
}

/**
 * Busca campos ignorando diferenças como userId x userID. A comparação é
 * case-insensitive, pois a própria documentação do AUVO mistura os padrões.
 */
function obterCampoCatalogoAuvo_(objeto, nomes) {
  if (!objeto || typeof objeto !== 'object') return '';

  const chaves = Object.keys(objeto);
  const indice = {};
  chaves.forEach(function(chave) {
    indice[chave.toLowerCase()] = chave;
  });

  for (let i = 0; i < nomes.length; i++) {
    const chaveReal = indice[String(nomes[i]).toLowerCase()];
    if (chaveReal !== undefined && temValorAuvo_(objeto[chaveReal])) {
      return objeto[chaveReal];
    }
  }

  return '';
}

function normalizarBooleanoCatalogoAuvo_(valor) {
  if (valor === true || String(valor).toLowerCase() === 'true') return true;
  if (valor === false || String(valor).toLowerCase() === 'false') return false;
  return '';
}

/**
 * Preenche somente os vazios da API com o historico. Assim, um valor novo da
 * API prevalece, mas a ausencia do campo no endpoint nao apaga a informacao.
 */
function mesclarCamposHistoricosAuvo_(linhasNovas) {
  const historico = criarMapaHistoricoAuvo_();
  const colunasComFallback = [
    9, 10, 11, 12, 13, 14, 18, 19, 20, 22,
    23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33
  ];

  return linhasNovas.map(function(linha) {
    const antiga = historico[String(linha[0] || '').trim()];
    if (!antiga) return linha;

    colunasComFallback.forEach(function(indice) {
      if (!temValorAuvo_(linha[indice]) && temValorAuvo_(antiga[indice])) {
        linha[indice] = antiga[indice];
      }
    });

    return linha;
  });
}

/**
 * O backup cria a base historica. Valores nao vazios da BASE atual prevalecem,
 * pois podem ter sido corrigidos manualmente ou por outra integracao.
 */
function criarMapaHistoricoAuvo_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const mapa = {};
  const nomes = [AUVO.ABA_BACKUP, AUVO.ABA_CLIENTES];

  nomes.forEach(function(nome, indiceFonte) {
    const aba = planilha.getSheetByName(nome);
    if (!aba || aba.getLastRow() < 2) return;

    validarCabecalhosBaseAuvo_(aba);

    const quantidade = aba.getLastRow() - 1;
    const linhas = aba
      .getRange(2, 1, quantidade, AUVO.TOTAL_COLUNAS)
      .getValues();

    linhas.forEach(function(linha) {
      const chave = String(linha[0] || '').trim();
      if (!chave) return;

      if (!mapa[chave]) {
        mapa[chave] = linha.slice();
        return;
      }

      // A BASE atual e a fonte mais recente e substitui somente com nao vazios.
      if (indiceFonte > 0) {
        linha.forEach(function(valor, indice) {
          if (temValorAuvo_(valor)) {
            mapa[chave][indice] = valor;
          }
        });
      }
    });
  });

  return mapa;
}

function validarFonteHistoricoAuvo_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const backup = planilha.getSheetByName(AUVO.ABA_BACKUP);

  if (!backup || backup.getLastRow() < 2) {
    throw new Error(
      'Crie a aba "' + AUVO.ABA_BACKUP +
      '" com a exportação anterior antes de atualizar.'
    );
  }

  validarCabecalhosBaseAuvo_(backup);

  const quantidade = backup.getLastRow() - 1;
  const amostra = backup
    .getRange(2, 1, quantidade, AUVO.TOTAL_COLUNAS)
    .getValues();

  const indicesHistoricos = [10, 11, 20, 22, 26, 27, 32, 33];
  const possuiHistorico = amostra.some(function(linha) {
    return indicesHistoricos.some(function(indice) {
      return temValorAuvo_(linha[indice]);
    });
  });

  if (!possuiHistorico) {
    throw new Error(
      'A aba BACKUP_BASE_AUVO existe, mas não contém os campos históricos esperados.'
    );
  }
}

function resolverIdsAuvo_(ids, mapa) {
  const lista = normalizarListaAuvo_(ids);
  if (!lista.length) return '';

  const nomes = lista.map(function(id) {
    return mapa[String(id)] || '';
  });

  // Se algum ID nao foi resolvido, devolve vazio para permitir o fallback.
  if (nomes.some(function(nome) { return !temValorAuvo_(nome); })) return '';
  return nomes.join(';');
}

function resolverIdAuvo_(id, mapa) {
  if (!temValorAuvo_(id)) return '';
  return mapa[String(id)] || '';
}

function juntarListaAuvo_(valor) {
  return normalizarListaAuvo_(valor)
    .map(function(item) {
      if (item && typeof item === 'object') {
        return (
          item.value ||
          item.description ||
          item.name ||
          item.email ||
          item.phoneNumber ||
          ''
        );
      }
      return item;
    })
    .filter(temValorAuvo_)
    .join(';');
}

function normalizarListaAuvo_(valor) {
  if (!temValorAuvo_(valor)) return [];
  return Array.isArray(valor) ? valor : [valor];
}

function normalizarStatusAuvo_(ativo) {
  if (ativo === true || String(ativo).toLowerCase() === 'true') return 'Ativo';
  if (ativo === false || String(ativo).toLowerCase() === 'false') return 'Inativo';
  return valorAuvo_(ativo);
}

function converterDataAuvo_(valor) {
  if (!temValorAuvo_(valor)) return '';
  const data = new Date(valor);
  return isNaN(data.getTime()) ? valor : data;
}

function valorAuvo_(valor) {
  return valor === null || valor === undefined ? '' : valor;
}

function temValorAuvo_(valor) {
  if (Array.isArray(valor)) return valor.length > 0;
  return valor !== null && valor !== undefined && String(valor).trim() !== '';
}

function gravarBaseAuvo_(linhas) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(AUVO.ABA_CLIENTES);

  if (!aba) {
    throw new Error('A aba "' + AUVO.ABA_CLIENTES + '" não foi encontrada.');
  }

  validarCabecalhosBaseAuvo_(aba);

  const totalNecessario = linhas.length + 1;
  if (aba.getMaxRows() < totalNecessario) {
    aba.insertRowsAfter(
      aba.getMaxRows(),
      totalNecessario - aba.getMaxRows()
    );
  }

  const quantidadeExistente = Math.max(aba.getLastRow() - 1, 1);
  aba
    .getRange(2, 1, quantidadeExistente, AUVO.TOTAL_COLUNAS)
    .clearContent();

  aba
    .getRange(2, 1, linhas.length, AUVO.TOTAL_COLUNAS)
    .setValues(linhas);

  // Codigos, documentos, telefones, e-mails, inscricoes e CEPs como texto.
  aba.getRange(2, 1, linhas.length, 5).setNumberFormat('@');
  aba.getRange(2, 8, linhas.length, 3).setNumberFormat('@');
  aba.getRange(2, 23, linhas.length, 12).setNumberFormat('@');

  // Ultima visita e data de cadastro.
  aba.getRange(2, 21, linhas.length, 2)
    .setNumberFormat('dd/MM/yyyy HH:mm');

  aba.setFrozenRows(1);
  SpreadsheetApp.flush();
}

function validarCabecalhosBaseAuvo_(aba) {
  const atuais = aba
    .getRange(1, 1, 1, AUVO.TOTAL_COLUNAS)
    .getDisplayValues()[0];

  for (let coluna = 0; coluna < CABECALHOS_BASE_AUVO_.length; coluna++) {
    if (
      String(atuais[coluna]).trim() !==
      String(CABECALHOS_BASE_AUVO_[coluna]).trim()
    ) {
      throw new Error(
        'Cabeçalho diferente na coluna ' + (coluna + 1) +
        '. Esperado: "' + CABECALHOS_BASE_AUVO_[coluna] +
        '". Encontrado: "' + atuais[coluna] + '".'
      );
    }
  }
}

function autenticarAuvo_() {
  const propriedades = PropertiesService.getScriptProperties();
  const apiKey = propriedades.getProperty('AUVO_API_KEY');
  const apiToken = propriedades.getProperty('AUVO_API_TOKEN');

  if (!apiKey || !apiToken) {
    throw new Error(
      'Credenciais não configuradas. Use o menu Integração AUVO.'
    );
  }

  const url =
    AUVO.BASE_URL +
    '/login?apiKey=' + encodeURIComponent(apiKey) +
    '&apiToken=' + encodeURIComponent(apiToken);

  const resposta = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Accept: 'application/json' },
    muteHttpExceptions: true
  });

  validarHttpAuvo_(resposta, 'autenticação');

  const json = JSON.parse(resposta.getContentText());
  const token =
    json.result && (json.result.accessToken || json.result.token);

  if (!token) {
    throw new Error(
      'O AUVO respondeu, mas não retornou o token de acesso.'
    );
  }

  return token;
}

function validarHttpAuvo_(resposta, operacao) {
  const codigo = resposta.getResponseCode();
  if (codigo >= 200 && codigo < 300) return;

  let mensagem = resposta.getContentText();
  if (mensagem.length > 500) mensagem = mensagem.substring(0, 500);

  throw new Error(
    'Erro na ' + operacao + '. HTTP ' + codigo + ': ' + mensagem
  );
}

function registrarLogAuvo_(processo, status, mensagem) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(AUVO.ABA_LOG);

  if (!aba) {
    aba = planilha.insertSheet(AUVO.ABA_LOG);
    aba.appendRow(['Data e hora', 'Processo', 'Status', 'Mensagem']);
    aba.setFrozenRows(1);
  }

  aba.appendRow([new Date(), processo, status, mensagem]);
}

function formatarDuracaoAuvo_(milissegundos) {
  const totalSegundos = Math.max(0, Math.round(milissegundos / 1000));
  const minutos = Math.floor(totalSegundos / 60);
  const segundos = totalSegundos % 60;
  return minutos + 'min ' + segundos + 's';
}

function testarCatalogosAuvo() {
  const token = autenticarAuvo_();
  const cache = CacheService.getScriptCache();

  ['USUARIOS', 'GRUPOS', 'SEGMENTOS', 'EQUIPES'].forEach(function(chave) {
    cache.remove('AUVO_CATALOGO_' + chave);
  });

  const catalogos = carregarCatalogosAuvo_(token);
  const resumo = [
    'Versão: ' + AUVO.VERSAO,
    'Usuários: ' + Object.keys(catalogos.usuarios).length,
    'Grupos: ' + Object.keys(catalogos.grupos).length,
    'Segmentos: ' + Object.keys(catalogos.segmentos).length,
    'Equipes: ' + Object.keys(catalogos.equipes).length
  ].join(' | ');

  const obrigatoriosOk =
    Object.keys(catalogos.usuarios).length > 0 &&
    Object.keys(catalogos.grupos).length > 0;

  registrarLogAuvo_(
    'CATÁLOGOS',
    obrigatoriosOk ? 'SUCESSO' : 'INCOMPLETO',
    resumo
  );
  SpreadsheetApp.getActiveSpreadsheet().toast(
    (obrigatoriosOk ? 'Catálogos validados. ' : 'Teste incompleto. ') + resumo,
    'Integração AUVO',
    12
  );

  if (!obrigatoriosOk) {
    throw new Error(
      'Teste incompleto: usuários e grupos precisam retornar valores maiores que zero. ' +
      resumo
    );
  }
}

function instalarAtualizacaoDiaria() {
  removerAtualizacaoDiaria_(false);

  ScriptApp.newTrigger('atualizarBaseAuvo')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();

  SpreadsheetApp.getUi().alert(
    'Atualização diária instalada para ocorrer entre 5h e 6h.'
  );
}

function removerAtualizacaoDiaria() {
  removerAtualizacaoDiaria_(true);
}

function removerAtualizacaoDiaria_(mostrarAviso) {
  ScriptApp.getProjectTriggers().forEach(function(gatilho) {
    if (gatilho.getHandlerFunction() === 'atualizarBaseAuvo') {
      ScriptApp.deleteTrigger(gatilho);
    }
  });

  if (mostrarAviso) {
    SpreadsheetApp.getUi().alert('Atualização diária removida.');
  }
}

/**
 * CONFIGURACAO DOS ALERTAS
 *
 * E-mail:
 * - Aceita varios enderecos separados por virgula ou ponto e virgula.
 *
 * WhatsApp:
 * - Usa exclusivamente a API oficial WhatsApp Cloud API da Meta.
 * - Os numeros devem incluir DDI e DDD, somente digitos (ex.: 5562999999999).
 * - Para envio automatico fora da janela de atendimento, crie e aprove um
 *   template de utilidade com dois parametros no corpo:
 *     {{1}} = quantidade de clientes sem grupo
 *     {{2}} = link da planilha
 * - Nome sugerido: alerta_clientes_sem_grupo
 * - Texto sugerido:
 *   "Alerta AUVO: existem {{1}} clientes sem grupo. Consulte: {{2}}"
 */
function configurarDestinatariosAlertasAuvo() {
  const ui = SpreadsheetApp.getUi();
  const propriedades = PropertiesService.getScriptProperties();

  const respostaEmails = ui.prompt(
    'Destinatários por e-mail',
    'Informe os e-mails separados por vírgula ou ponto e vírgula:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respostaEmails.getSelectedButton() !== ui.Button.OK) return;

  const emails = normalizarEmailsAlertaAuvo_(
    respostaEmails.getResponseText()
  );
  if (!emails.length) {
    ui.alert('Informe pelo menos um e-mail válido.');
    return;
  }

  const respostaTelefones = ui.prompt(
    'Destinatários por WhatsApp',
    'Informe os números com DDI e DDD, separados por vírgula ou ponto e vírgula. Ex.: 5562999999999',
    ui.ButtonSet.OK_CANCEL
  );
  if (respostaTelefones.getSelectedButton() !== ui.Button.OK) return;

  const telefones = normalizarTelefonesAlertaAuvo_(
    respostaTelefones.getResponseText()
  );
  if (!telefones.length) {
    ui.alert('Informe pelo menos um número válido com DDI e DDD.');
    return;
  }

  propriedades.setProperties({
    ALERTA_SEM_GRUPO_EMAILS: emails.join(','),
    ALERTA_SEM_GRUPO_WHATSAPP: telefones.join(',')
  });

  registrarLogAuvo_(
    'ALERTA CONFIGURAÇÃO',
    'SUCESSO',
    emails.length + ' e-mail(s) e ' + telefones.length +
      ' WhatsApp(s) configurados.'
  );
  ui.alert('Destinatários salvos com sucesso.');
}

function autorizarEnvioEmailAuvo() {
  const escopoEmail = 'https://www.googleapis.com/auth/script.send_mail';

  // Interrompe a execução e abre a tela de consentimento se o novo escopo
  // ainda não tiver sido concedido. Depois de autorizar, execute novamente.
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [escopoEmail]);

  const cota = MailApp.getRemainingDailyQuota();
  const mensagem = 'Envio de e-mail autorizado. Cota restante hoje: ' + cota +
    ' destinatário(s).';

  registrarLogAuvo_('AUTORIZAÇÃO E-MAIL', 'SUCESSO', mensagem);
  SpreadsheetApp.getUi().alert(
    'Autorização concluída',
    mensagem,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function configurarWhatsAppMetaAuvo() {
  const ui = SpreadsheetApp.getUi();
  const propriedades = PropertiesService.getScriptProperties();

  const respostaId = ui.prompt(
    'WhatsApp oficial - Meta',
    'Cole o ID do número de telefone (Phone Number ID):',
    ui.ButtonSet.OK_CANCEL
  );
  if (respostaId.getSelectedButton() !== ui.Button.OK) return;

  const phoneNumberId = respostaId.getResponseText().trim();
  if (!/^\d+$/.test(phoneNumberId)) {
    ui.alert('O Phone Number ID deve conter somente números.');
    return;
  }

  const respostaToken = ui.prompt(
    'WhatsApp oficial - Meta',
    'Cole o token permanente do usuário do sistema:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respostaToken.getSelectedButton() !== ui.Button.OK) return;

  const token = respostaToken.getResponseText().trim();
  if (!token) {
    ui.alert('O token da Meta é obrigatório.');
    return;
  }

  const respostaTemplate = ui.prompt(
    'Template aprovado',
    'Informe o nome do template. Sugestão: alerta_clientes_sem_grupo',
    ui.ButtonSet.OK_CANCEL
  );
  if (respostaTemplate.getSelectedButton() !== ui.Button.OK) return;

  const template = respostaTemplate.getResponseText().trim();
  if (!/^[a-z0-9_]+$/.test(template)) {
    ui.alert('Use somente letras minúsculas, números e underscore no nome do template.');
    return;
  }

  const respostaIdioma = ui.prompt(
    'Idioma do template',
    'Informe o código do idioma aprovado. Para português do Brasil: pt_BR',
    ui.ButtonSet.OK_CANCEL
  );
  if (respostaIdioma.getSelectedButton() !== ui.Button.OK) return;

  const idioma = respostaIdioma.getResponseText().trim() || 'pt_BR';

  const respostaVersao = ui.prompt(
    'Versão da Graph API',
    'Informe a versão mostrada no painel da Meta. Ex.: v23.0',
    ui.ButtonSet.OK_CANCEL
  );
  if (respostaVersao.getSelectedButton() !== ui.Button.OK) return;

  const versao = respostaVersao.getResponseText().trim();
  if (!/^v\d+\.\d+$/.test(versao)) {
    ui.alert('Versão inválida. Use o formato v23.0, por exemplo.');
    return;
  }

  propriedades.setProperties({
    META_WHATSAPP_PHONE_NUMBER_ID: phoneNumberId,
    META_WHATSAPP_TOKEN: token,
    META_WHATSAPP_TEMPLATE: template,
    META_WHATSAPP_TEMPLATE_LANGUAGE: idioma,
    META_GRAPH_API_VERSION: versao
  });

  registrarLogAuvo_(
    'WHATSAPP CONFIGURAÇÃO',
    'SUCESSO',
    'WhatsApp oficial configurado. Template: ' + template +
      ' | Idioma: ' + idioma + ' | Graph API: ' + versao + '.'
  );
  ui.alert('WhatsApp oficial configurado com sucesso.');
}

function testarAlertaClientesSemGrupoAuvo() {
  const propriedades = PropertiesService.getScriptProperties();
  const emails = normalizarEmailsAlertaAuvo_(
    propriedades.getProperty('ALERTA_SEM_GRUPO_EMAILS') || ''
  );

  if (emails.length) {
    ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
      'https://www.googleapis.com/auth/script.send_mail'
    ]);
  }

  const linhas = obterLinhasAtuaisBaseAuvo_();
  const resultado = dispararAlertaClientesSemGrupoAuvo_(linhas, true);

  SpreadsheetApp.getUi().alert(
    'Teste concluído',
    resultado,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function dispararAlertaClientesSemGrupoAuvo_(linhas, testeManual) {
  const clientes = identificarClientesSemGrupoAuvo_(linhas);

  if (!clientes.length) {
    const mensagem = 'Nenhum cliente sem grupo foi encontrado.';
    registrarLogAuvo_('ALERTA SEM GRUPO', 'SEM PENDÊNCIAS', mensagem);
    return mensagem;
  }

  const propriedades = PropertiesService.getScriptProperties();
  const emails = normalizarEmailsAlertaAuvo_(
    propriedades.getProperty('ALERTA_SEM_GRUPO_EMAILS') || ''
  );
  const telefones = normalizarTelefonesAlertaAuvo_(
    propriedades.getProperty('ALERTA_SEM_GRUPO_WHATSAPP') || ''
  );

  const resultados = [];

  if (emails.length) {
    try {
      enviarEmailClientesSemGrupoAuvo_(clientes, emails, testeManual);
      registrarLogAuvo_(
        'ALERTA E-MAIL',
        'SUCESSO',
        clientes.length + ' cliente(s) enviados para ' + emails.length +
          ' destinatário(s).'
      );
      resultados.push('E-mail: enviado');
    } catch (erroEmail) {
      registrarLogAuvo_('ALERTA E-MAIL', 'ERRO', erroEmail.message);
      resultados.push('E-mail: erro - ' + erroEmail.message);
    }
  } else {
    registrarLogAuvo_(
      'ALERTA E-MAIL',
      'NÃO CONFIGURADO',
      'Cadastre os destinatários no menu da integração.'
    );
    resultados.push('E-mail: não configurado');
  }

  if (telefones.length && whatsappMetaConfiguradoAuvo_()) {
    try {
      const quantidadeEnviada = enviarWhatsAppClientesSemGrupoAuvo_(
        clientes,
        telefones
      );
      registrarLogAuvo_(
        'ALERTA WHATSAPP',
        'SUCESSO',
        clientes.length + ' cliente(s) informados a ' + quantidadeEnviada +
          ' destinatário(s).'
      );
      resultados.push('WhatsApp: enviado para ' + quantidadeEnviada);
    } catch (erroWhatsApp) {
      registrarLogAuvo_('ALERTA WHATSAPP', 'ERRO', erroWhatsApp.message);
      resultados.push('WhatsApp: erro - ' + erroWhatsApp.message);
    }
  } else if (!telefones.length) {
    registrarLogAuvo_(
      'ALERTA WHATSAPP',
      'NÃO CONFIGURADO',
      'Cadastre os números no menu da integração.'
    );
    resultados.push('WhatsApp: não configurado');
  } else {
    registrarLogAuvo_(
      'ALERTA WHATSAPP',
      'PENDENTE',
      'Destinatários salvos; falta configurar a API oficial da Meta na opção 6.'
    );
    resultados.push('WhatsApp: pendente de configuração');
  }

  const houveSucesso = resultados.some(function(item) {
    return item.indexOf('enviado') >= 0;
  });
  registrarLogAuvo_(
    'ALERTA SEM GRUPO',
    houveSucesso ? 'CONCLUÍDO' : 'INCOMPLETO',
    clientes.length + ' cliente(s) sem grupo. ' + resultados.join(' | ')
  );

  return clientes.length + ' cliente(s) sem grupo. ' + resultados.join(' | ');
}

function obterLinhasAtuaisBaseAuvo_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    AUVO.ABA_CLIENTES
  );
  if (!aba || aba.getLastRow() < 2) {
    throw new Error('A BASE_AUVO não possui clientes para testar.');
  }

  validarCabecalhosBaseAuvo_(aba);
  return aba.getRange(
    2,
    1,
    aba.getLastRow() - 1,
    AUVO.TOTAL_COLUNAS
  ).getDisplayValues();
}

function identificarClientesSemGrupoAuvo_(linhas) {
  const indiceCodigo = obterIndiceCabecalhoAuvo_('Código');
  const indiceCliente = obterIndiceCabecalhoAuvo_('cliente');
  const indiceDocumento = obterIndiceCabecalhoAuvo_('CPF ou CNPJ');
  const indiceResponsavel = obterIndiceCabecalhoAuvo_('Usuário Auvo responsável');
  const indiceGrupo = obterIndiceCabecalhoAuvo_('Grupos');
  const indiceStatus = obterIndiceCabecalhoAuvo_('Status');

  return linhas
    .filter(function(linha) {
      return temValorAuvo_(linha[indiceCodigo]) &&
        !temValorAuvo_(linha[indiceGrupo]);
    })
    .map(function(linha) {
      return {
        codigo: String(linha[indiceCodigo] || '').trim(),
        cliente: String(linha[indiceCliente] || '').trim(),
        documento: String(linha[indiceDocumento] || '').trim(),
        responsavel: String(linha[indiceResponsavel] || '').trim(),
        status: String(linha[indiceStatus] || '').trim()
      };
    })
    .sort(function(a, b) {
      return a.cliente.localeCompare(b.cliente);
    });
}

function obterIndiceCabecalhoAuvo_(nome) {
  const indice = CABECALHOS_BASE_AUVO_.indexOf(nome);
  if (indice < 0) {
    throw new Error('Cabeçalho não mapeado: ' + nome);
  }
  return indice;
}

function enviarEmailClientesSemGrupoAuvo_(clientes, emails, testeManual) {
  const autorizacao = ScriptApp.getAuthorizationInfo(
    ScriptApp.AuthMode.FULL,
    ['https://www.googleapis.com/auth/script.send_mail']
  );
  if (autorizacao.getAuthorizationStatus() ===
      ScriptApp.AuthorizationStatus.REQUIRED) {
    throw new Error(
      'Envio de e-mail ainda não autorizado. Execute a opção 5A do menu e aceite as permissões.'
    );
  }

  if (MailApp.getRemainingDailyQuota() < emails.length) {
    throw new Error('Cota diária de e-mail insuficiente para os destinatários.');
  }

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const link = planilha.getUrl() + '#gid=' +
    planilha.getSheetByName(AUVO.ABA_CLIENTES).getSheetId();
  const prefixo = testeManual ? '[TESTE] ' : '';
  const assunto = prefixo + 'OLFATI | ' + clientes.length +
    ' cliente(s) sem grupo no AUVO';

  const linhasHtml = clientes.map(function(cliente) {
    return '<tr>' +
      '<td style="padding:7px;border-bottom:1px solid #ddd">' +
        escaparHtmlAuvo_(cliente.codigo) + '</td>' +
      '<td style="padding:7px;border-bottom:1px solid #ddd">' +
        escaparHtmlAuvo_(cliente.cliente) + '</td>' +
      '<td style="padding:7px;border-bottom:1px solid #ddd">' +
        escaparHtmlAuvo_(cliente.documento || '-') + '</td>' +
      '<td style="padding:7px;border-bottom:1px solid #ddd">' +
        escaparHtmlAuvo_(cliente.responsavel || '-') + '</td>' +
      '<td style="padding:7px;border-bottom:1px solid #ddd">' +
        escaparHtmlAuvo_(cliente.status || '-') + '</td>' +
      '</tr>';
  }).join('');

  const html =
    '<div style="font-family:Arial,sans-serif;color:#222">' +
    '<h2 style="color:#a83b2d">Clientes sem grupo no AUVO</h2>' +
    '<p>Após a atualização da BASE_AUVO, foram identificados <strong>' +
      clientes.length + '</strong> cliente(s) sem grupo.</p>' +
    '<table style="border-collapse:collapse;width:100%;max-width:950px">' +
    '<thead><tr style="background:#f2e4df;text-align:left">' +
    '<th style="padding:8px">Código</th>' +
    '<th style="padding:8px">Cliente</th>' +
    '<th style="padding:8px">CPF/CNPJ</th>' +
    '<th style="padding:8px">Responsável</th>' +
    '<th style="padding:8px">Status</th>' +
    '</tr></thead><tbody>' + linhasHtml + '</tbody></table>' +
    '<p style="margin-top:18px"><a href="' + escaparHtmlAuvo_(link) +
      '">Abrir BASE_AUVO</a></p>' +
    '<p style="color:#777;font-size:12px">Mensagem automática da integração AUVO.</p>' +
    '</div>';

  const texto =
    prefixo + 'Foram identificados ' + clientes.length +
    ' cliente(s) sem grupo no AUVO. Consulte a BASE_AUVO: ' + link;

  MailApp.sendEmail({
    to: emails.join(','),
    subject: assunto,
    body: texto,
    htmlBody: html,
    name: 'Integração AUVO - OLFATI'
  });
}

function enviarWhatsAppClientesSemGrupoAuvo_(clientes, telefones) {
  const propriedades = PropertiesService.getScriptProperties();
  const phoneNumberId = propriedades.getProperty(
    'META_WHATSAPP_PHONE_NUMBER_ID'
  );
  const token = propriedades.getProperty('META_WHATSAPP_TOKEN');
  const template = propriedades.getProperty('META_WHATSAPP_TEMPLATE');
  const idioma = propriedades.getProperty(
    'META_WHATSAPP_TEMPLATE_LANGUAGE'
  );
  const versao = propriedades.getProperty('META_GRAPH_API_VERSION');

  if (!phoneNumberId || !token || !template || !idioma || !versao) {
    throw new Error(
      'WhatsApp oficial não configurado. Use a opção 6 do menu.'
    );
  }

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const link = planilha.getUrl() + '#gid=' +
    planilha.getSheetByName(AUVO.ABA_CLIENTES).getSheetId();
  const url = 'https://graph.facebook.com/' + versao + '/' +
    phoneNumberId + '/messages';
  let enviados = 0;
  const erros = [];

  telefones.forEach(function(telefone) {
    const payload = {
      messaging_product: 'whatsapp',
      recipient_type: 'individual',
      to: telefone,
      type: 'template',
      template: {
        name: template,
        language: { code: idioma },
        components: [{
          type: 'body',
          parameters: [
            { type: 'text', text: String(clientes.length) },
            { type: 'text', text: link }
          ]
        }]
      }
    };

    const resposta = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const codigo = resposta.getResponseCode();
    if (codigo >= 200 && codigo < 300) {
      enviados++;
      return;
    }

    let detalhe = resposta.getContentText();
    if (detalhe.length > 300) detalhe = detalhe.substring(0, 300);
    erros.push(mascararTelefoneAuvo_(telefone) + ': HTTP ' + codigo +
      ' - ' + detalhe);
  });

  if (erros.length) {
    throw new Error(
      enviados + ' enviado(s), ' + erros.length +
      ' falha(s). ' + erros.join(' | ')
    );
  }

  return enviados;
}

function whatsappMetaConfiguradoAuvo_() {
  const propriedades = PropertiesService.getScriptProperties();
  return Boolean(
    propriedades.getProperty('META_WHATSAPP_PHONE_NUMBER_ID') &&
    propriedades.getProperty('META_WHATSAPP_TOKEN') &&
    propriedades.getProperty('META_WHATSAPP_TEMPLATE') &&
    propriedades.getProperty('META_WHATSAPP_TEMPLATE_LANGUAGE') &&
    propriedades.getProperty('META_GRAPH_API_VERSION')
  );
}

function normalizarEmailsAlertaAuvo_(texto) {
  const vistos = {};
  return String(texto || '')
    .split(/[;,\s]+/)
    .map(function(item) { return item.trim().toLowerCase(); })
    .filter(function(email) {
      const valido = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
      if (!valido || vistos[email]) return false;
      vistos[email] = true;
      return true;
    });
}

function normalizarTelefonesAlertaAuvo_(texto) {
  const vistos = {};
  return String(texto || '')
    .split(/[;,\n]+/)
    .map(function(item) { return item.replace(/\D/g, ''); })
    .filter(function(telefone) {
      const valido = /^\d{10,15}$/.test(telefone);
      if (!valido || vistos[telefone]) return false;
      vistos[telefone] = true;
      return true;
    });
}

function escaparHtmlAuvo_(valor) {
  return String(valor === null || valor === undefined ? '' : valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function mascararTelefoneAuvo_(telefone) {
  const texto = String(telefone || '');
  return texto.length <= 4 ? '****' :
    texto.substring(0, 2) + '*****' + texto.substring(texto.length - 4);
}

/**
 * Diagnostico somente de leitura para descobrir a estrutura real que a conta
 * AUVO devolve para questionarios e respostas de uma tarefa concluida.
 * Nao altera BASE_QUESTIONARIO nem BASE_AUVO.
 */
function diagnosticarQuestionariosAuvo() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const token = autenticarAuvo_();
  const codigoTarefa = obterTarefaAmostraQuestionarioAuvo_();
  const linhas = [['Origem', 'HTTP', 'Caminho', 'Tipo', 'Amostra']];
  let questionariosOk = false;
  let tarefaOk = false;

  registrarLogAuvo_(
    'QUESTIONÁRIOS DIAGNÓSTICO',
    'INÍCIO',
    'Versão: ' + AUVO.VERSAO + ' | Tarefa de amostra: ' + codigoTarefa
  );

  const consultaQuestionarios = consultarJsonDiagnosticoAuvo_(
    token,
    '/questionnaires/?page=1&pageSize=100&order=asc'
  );
  adicionarResultadoDiagnosticoAuvo_(
    linhas,
    'LISTA_QUESTIONARIOS',
    consultaQuestionarios
  );
  questionariosOk = consultaQuestionarios.ok;

  const rotasTarefa = [
    '/tasks/' + encodeURIComponent(codigoTarefa),
    '/tasks/' + encodeURIComponent(codigoTarefa) + '/'
  ];
  for (let indice = 0; indice < rotasTarefa.length; indice++) {
    const consultaTarefa = consultarJsonDiagnosticoAuvo_(
      token,
      rotasTarefa[indice]
    );
    adicionarResultadoDiagnosticoAuvo_(
      linhas,
      'TAREFA_' + codigoTarefa + '_' + (indice + 1),
      consultaTarefa
    );
    if (consultaTarefa.ok) {
      tarefaOk = true;
      break;
    }
  }

  let aba = planilha.getSheetByName(AUVO.ABA_DIAGNOSTICO_QUESTIONARIOS);
  if (!aba) aba = planilha.insertSheet(AUVO.ABA_DIAGNOSTICO_QUESTIONARIOS);
  aba.clearContents();
  aba.getRange(1, 1, linhas.length, linhas[0].length).setValues(linhas);
  aba.setFrozenRows(1);
  aba.getRange(1, 1, 1, 5)
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF');
  aba.setColumnWidth(1, 210);
  aba.setColumnWidth(2, 70);
  aba.setColumnWidth(3, 420);
  aba.setColumnWidth(4, 110);
  aba.setColumnWidth(5, 500);
  if (linhas.length > 1) {
    aba.getRange(2, 1, linhas.length - 1, 5).setWrap(true);
  }

  const status = questionariosOk && tarefaOk ? 'SUCESSO' : 'INCOMPLETO';
  const mensagem =
    'Definições: ' + (questionariosOk ? 'OK' : 'falha') +
    ' | Tarefa ' + codigoTarefa + ': ' + (tarefaOk ? 'OK' : 'falha') +
    ' | Linhas do diagnóstico: ' + (linhas.length - 1) +
    '. A BASE_QUESTIONARIO não foi alterada.';

  registrarLogAuvo_('QUESTIONÁRIOS DIAGNÓSTICO', status, mensagem);
  planilha.toast(mensagem, 'Integração AUVO', 15);
}

function obterTarefaAmostraQuestionarioAuvo_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(AUVO.ABA_QUESTIONARIOS);
  if (!aba) {
    throw new Error('A aba ' + AUVO.ABA_QUESTIONARIOS + ' não foi encontrada.');
  }

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) {
    throw new Error('A BASE_QUESTIONARIO não possui tarefas para o diagnóstico.');
  }

  const valores = aba.getRange(2, 1, ultimaLinha - 1, 20).getValues();
  for (let indice = 0; indice < valores.length; indice++) {
    const codigo = valores[indice][0];
    const dataResposta = valores[indice][19];
    if (codigo !== '' && codigo !== null && dataResposta !== '' && dataResposta !== null) {
      return String(codigo).replace(/\.0$/, '').trim();
    }
  }

  throw new Error(
    'Nenhuma tarefa com data de resposta foi encontrada na BASE_QUESTIONARIO.'
  );
}

function consultarJsonDiagnosticoAuvo_(token, endpoint) {
  const resposta = UrlFetchApp.fetch(AUVO.BASE_URL + endpoint, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  });

  const http = resposta.getResponseCode();
  const texto = resposta.getContentText();
  let json = null;
  let erroJson = '';

  try {
    json = texto ? JSON.parse(texto) : null;
  } catch (erro) {
    erroJson = 'Resposta não é JSON: ' + erro.message;
  }

  return {
    ok: http >= 200 && http < 300 && json !== null,
    http: http,
    endpoint: endpoint,
    json: json,
    erro: erroJson || (http >= 200 && http < 300 ? '' : texto.substring(0, 500))
  };
}

function adicionarResultadoDiagnosticoAuvo_(linhas, origem, consulta) {
  linhas.push([
    origem,
    consulta.http,
    consulta.endpoint,
    consulta.ok ? 'RESPOSTA' : 'ERRO',
    consulta.ok ? 'JSON recebido' : String(consulta.erro || 'Falha sem detalhe')
  ]);

  if (!consulta.ok) return;

  percorrerJsonDiagnosticoAuvo_(
    consulta.json,
    '$',
    origem,
    consulta.http,
    linhas,
    0
  );
}

function percorrerJsonDiagnosticoAuvo_(valor, caminho, origem, http, linhas, nivel) {
  if (linhas.length >= 1500 || nivel > 8) return;

  if (Array.isArray(valor)) {
    linhas.push([origem, http, caminho, 'ARRAY', valor.length + ' item(ns)']);
    const limite = Math.min(valor.length, 5);
    for (let indice = 0; indice < limite; indice++) {
      percorrerJsonDiagnosticoAuvo_(
        valor[indice],
        caminho + '[' + indice + ']',
        origem,
        http,
        linhas,
        nivel + 1
      );
    }
    if (valor.length > limite) {
      linhas.push([
        origem,
        http,
        caminho,
        'RESUMO',
        (valor.length - limite) + ' item(ns) adicionais não exibidos'
      ]);
    }
    return;
  }

  if (valor !== null && typeof valor === 'object') {
    const chaves = Object.keys(valor);
    linhas.push([origem, http, caminho, 'OBJETO', chaves.length + ' campo(s)']);
    chaves.forEach(function(chave) {
      if (linhas.length >= 1500) return;
      percorrerJsonDiagnosticoAuvo_(
        valor[chave],
        caminho + '.' + chave,
        origem,
        http,
        linhas,
        nivel + 1
      );
    });
    return;
  }

  linhas.push([
    origem,
    http,
    caminho,
    valor === null ? 'NULO' : typeof valor,
    resumirValorDiagnosticoAuvo_(valor)
  ]);
}

function resumirValorDiagnosticoAuvo_(valor) {
  if (valor === null || valor === undefined) return '';
  let texto = String(valor).replace(/[\r\n]+/g, ' ').trim();
  if (/^https?:\/\//i.test(texto)) {
    texto = '[URL] ' + texto.substring(0, 160);
  }
  return texto.length > 300 ? texto.substring(0, 300) + '…' : texto;
}

/**
 * Sincroniza os questionarios cadastrados no AUVO com uma aba de configuracao.
 * Decisoes ja tomadas pela operacao sao preservadas. Questionarios novos ficam
 * como NAO, exceto os relatorios 03 e 04 na primeira criacao da configuracao.
 */
function sincronizarConfiguracaoQuestionariosAuvo() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Outra execução da integração AUVO está em andamento.');
  }

  try {
    const token = autenticarAuvo_();
    const endpoints = [
      '/questionnaires/?page=1&pageSize=500&order=asc',
      '/questionnaires/?Page=1&PageSize=500&Order=asc'
    ];
    let entidades = [];
    const tentativas = [];

    for (let indice = 0; indice < endpoints.length; indice++) {
      const consulta = consultarJsonDiagnosticoAuvo_(token, endpoints[indice]);
      const lista = consulta.ok ? extrairListaCatalogoAuvo_(consulta.json) : [];
      tentativas.push(
        endpoints[indice] + ': HTTP ' + consulta.http + ', ' + lista.length + ' item(ns)'
      );
      if (lista.length) {
        entidades = lista;
        break;
      }
    }

    const catalogo = {};
    entidades.forEach(function(entidade) {
      const id = extrairIdCatalogoAuvo_(entidade);
      const nome = extrairNomeCatalogoAuvo_(entidade);
      if (temValorAuvo_(id) && temValorAuvo_(nome)) {
        catalogo[String(id).trim()] = String(nome).trim();
      }
    });

    const idsCatalogo = Object.keys(catalogo);
    if (!idsCatalogo.length) {
      throw new Error(
        'Nenhum questionário foi reconhecido no AUVO. Tentativas: ' +
        tentativas.join(' | ')
      );
    }

    const aba = obterOuCriarAbaConfigQuestionariosAuvo_();
    const existentes = criarMapaConfigQuestionariosAuvo_(aba);
    const agora = new Date();
    const linhas = [];

    idsCatalogo
      .sort(function(a, b) {
        return catalogo[a].localeCompare(catalogo[b], 'pt-BR', { numeric: true });
      })
      .forEach(function(id) {
        const anterior = existentes[id];
        const importar = anterior
          ? anterior.importar
          : (questionarioPadraoContabilizadoAuvo_(catalogo[id]) ? 'SIM' : 'NÃO');
        const produtividade = anterior
          ? anterior.produtividade
          : (questionarioPadraoProdutividadeAuvo_(catalogo[id]) ? 'SIM' : 'NÃO');

        linhas.push([
          id,
          catalogo[id],
          normalizarOpcaoQuestionarioAuvo_(importar),
          normalizarOpcaoQuestionarioAuvo_(produtividade),
          anterior ? anterior.inicio : '',
          anterior ? anterior.fim : '',
          'SIM',
          agora,
          ''
        ]);
        delete existentes[id];
      });

    Object.keys(existentes)
      .sort(function(a, b) {
        return existentes[a].nome.localeCompare(existentes[b].nome, 'pt-BR', { numeric: true });
      })
      .forEach(function(id) {
        const anterior = existentes[id];
        linhas.push([
          id,
          anterior.nome,
          normalizarOpcaoQuestionarioAuvo_(anterior.importar),
          normalizarOpcaoQuestionarioAuvo_(anterior.produtividade),
          anterior.inicio,
          anterior.fim,
          'NÃO',
          agora,
          ''
        ]);
      });

    const ultimaLinhaAnterior = Math.max(aba.getLastRow(), 2);
    aba.getRange(2, 1, ultimaLinhaAnterior - 1, 9).clearContent();
    if (linhas.length) {
      aba.getRange(2, 1, linhas.length, 9).setValues(linhas);
    }
    formatarAbaConfigQuestionariosAuvo_(aba, linhas.length);

    const resultado = validarConfiguracaoQuestionariosAuvo_(false);
    const mensagem =
      idsCatalogo.length + ' questionário(s) encontrado(s) no AUVO. ' +
      resultado.ativos + ' selecionado(s) para contabilização. ' +
      resultado.agendados + ' agendado(s) para data futura.';

    registrarLogAuvo_('QUESTIONÁRIOS CONFIGURAÇÃO', 'SUCESSO', mensagem);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      mensagem,
      'Integração AUVO',
      15
    );
  } catch (erro) {
    registrarLogAuvo_(
      'QUESTIONÁRIOS CONFIGURAÇÃO',
      'ERRO',
      String(erro && erro.message ? erro.message : erro)
    );
    throw erro;
  } finally {
    lock.releaseLock();
  }
}

function obterOuCriarAbaConfigQuestionariosAuvo_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(AUVO.ABA_CONFIG_QUESTIONARIOS);
  if (!aba) aba = planilha.insertSheet(AUVO.ABA_CONFIG_QUESTIONARIOS);
  return aba;
}

function criarMapaConfigQuestionariosAuvo_(aba) {
  const mapa = {};
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return mapa;

  const formatoNovo = String(aba.getRange(1, 4).getDisplayValue()).trim() ===
    'Contabilizar na produtividade?';
  const largura = formatoNovo ? 9 : 8;
  const valores = aba.getRange(2, 1, ultimaLinha - 1, largura).getValues();
  valores.forEach(function(linha) {
    const id = String(linha[0] || '').trim();
    if (!id || mapa[id]) return;
    mapa[id] = {
      id: id,
      nome: String(linha[1] || '').trim(),
      importar: normalizarOpcaoQuestionarioAuvo_(linha[2]),
      produtividade: formatoNovo
        ? normalizarOpcaoQuestionarioAuvo_(linha[3])
        : (questionarioPadraoProdutividadeAuvo_(linha[1]) ? 'SIM' : 'NÃO'),
      inicio: linha[formatoNovo ? 4 : 3] || '',
      fim: linha[formatoNovo ? 5 : 4] || '',
      encontrado: normalizarOpcaoQuestionarioAuvo_(linha[formatoNovo ? 6 : 5])
    };
  });
  return mapa;
}

function formatarAbaConfigQuestionariosAuvo_(aba, quantidadeLinhas) {
  const totalLinhas = Math.max(quantidadeLinhas + 1, 2);
  aba.getRange(1, 1, 1, 9).setValues([[
    'ID do questionário', 'Nome do questionário', 'Importar para a base?',
    'Contabilizar na produtividade?', 'Vigência inicial', 'Vigência final',
    'Encontrado no AUVO?', 'Última sincronização', 'Situação'
  ]]);
  const cabecalho = aba.getRange(1, 1, 1, 9);
  cabecalho
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  cabecalho.setNotes([[
    'Identificador interno do AUVO. Não editar.',
    'Nome sincronizado do AUVO. Não editar.',
    'Selecione SIM para trazer as respostas deste questionário para a BASE_QUESTIONARIO.',
    'Selecione SIM somente quando cada resposta representar atendimento da produtividade técnica.',
    'Primeiro dia em que as respostas passam a ser contabilizadas. Vazio = sem limite inicial.',
    'Último dia em que as respostas são contabilizadas. Vazio = sem limite final.',
    'Indica se o questionário ainda existe no catálogo retornado pelo AUVO.',
    'Data e hora da última consulta ao catálogo do AUVO.',
    'Resultado da validação da regra de contabilização.'
  ]]);

  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 150);
  aba.setColumnWidth(2, 430);
  aba.setColumnWidth(3, 145);
  aba.setColumnWidth(4, 185);
  aba.setColumnWidth(5, 135);
  aba.setColumnWidth(6, 135);
  aba.setColumnWidth(7, 155);
  aba.setColumnWidth(8, 175);
  aba.setColumnWidth(9, 260);

  if (quantidadeLinhas > 0) {
    const validacaoSimNao = SpreadsheetApp.newDataValidation()
      .requireValueInList(['SIM', 'NÃO'], true)
      .setAllowInvalid(false)
      .build();
    aba.getRange(2, 3, quantidadeLinhas, 2).setDataValidation(validacaoSimNao);
    aba.getRange(2, 5, quantidadeLinhas, 2).setNumberFormat('dd/MM/yyyy');
    aba.getRange(2, 8, quantidadeLinhas, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    aba.getRange(2, 1, quantidadeLinhas, 9).setVerticalAlignment('middle');
    aba.getRange(2, 2, quantidadeLinhas, 1).setWrap(true);
    aba.getRange(2, 9, quantidadeLinhas, 1).setWrap(true);
  }

  const filtro = aba.getFilter();
  if (filtro) filtro.remove();
  aba.getRange(1, 1, totalLinhas, 9).createFilter();
}

function questionarioPadraoContabilizadoAuvo_(nome) {
  const texto = String(nome || '').trim();
  return /^0?3(?:\.|\s|-)/.test(texto) || /^0?4(?:\.|\s|-)/.test(texto);
}

function questionarioPadraoProdutividadeAuvo_(nome) {
  return /^0?4(?:\.|\s|-)/.test(String(nome || '').trim());
}

function normalizarOpcaoQuestionarioAuvo_(valor) {
  const texto = String(valor || '').trim().toUpperCase();
  return texto === 'SIM' ? 'SIM' : 'NÃO';
}

function validarConfiguracaoQuestionariosAuvo() {
  return validarConfiguracaoQuestionariosAuvo_(true);
}

function validarConfiguracaoQuestionariosAuvo_(mostrarAviso) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(AUVO.ABA_CONFIG_QUESTIONARIOS);
  if (!aba || aba.getLastRow() < 2) {
    throw new Error(
      'A configuração está vazia. Execute primeiro “9. Sincronizar e selecionar questionários”.'
    );
  }

  const quantidade = aba.getLastRow() - 1;
  const valores = aba.getRange(2, 1, quantidade, 9).getValues();
  const ids = {};
  const erros = [];
  let ativos = 0;
  let agendados = 0;
  let encerrados = 0;
  const hoje = inicioDoDiaQuestionarioAuvo_(new Date());
  const situacoes = [];

  valores.forEach(function(linha, indice) {
    const numeroLinha = indice + 2;
    const id = String(linha[0] || '').trim();
    const nome = String(linha[1] || '').trim();
    const importarOriginal = String(linha[2] || '').trim().toUpperCase();
    const importar = normalizarOpcaoQuestionarioAuvo_(linha[2]);
    const produtividadeOriginal = String(linha[3] || '').trim().toUpperCase();
    const produtividade = normalizarOpcaoQuestionarioAuvo_(linha[3]);
    const inicio = normalizarDataConfigQuestionarioAuvo_(linha[4]);
    const fim = normalizarDataConfigQuestionarioAuvo_(linha[5]);
    const encontrado = normalizarOpcaoQuestionarioAuvo_(linha[6]);
    let situacao = 'NÃO CONTABILIZADO';

    if (!id) erros.push('Linha ' + numeroLinha + ': ID vazio.');
    if (!nome) erros.push('Linha ' + numeroLinha + ': nome vazio.');
    if (id && ids[id]) erros.push('Linha ' + numeroLinha + ': ID duplicado ' + id + '.');
    if (id) ids[id] = true;
    if (importarOriginal !== 'SIM' && importarOriginal !== 'NÃO') {
      erros.push('Linha ' + numeroLinha + ': use somente SIM ou NÃO em Importar para a base?.');
    }
    if (produtividadeOriginal !== 'SIM' && produtividadeOriginal !== 'NÃO') {
      erros.push('Linha ' + numeroLinha + ': use somente SIM ou NÃO em Contabilizar na produtividade?.');
    }
    if (produtividade === 'SIM' && importar !== 'SIM') {
      erros.push('Linha ' + numeroLinha + ': produtividade SIM exige importação SIM.');
    }
    if (linha[4] && !inicio) erros.push('Linha ' + numeroLinha + ': vigência inicial inválida.');
    if (linha[5] && !fim) erros.push('Linha ' + numeroLinha + ': vigência final inválida.');
    if (inicio && fim && inicio.getTime() > fim.getTime()) {
      erros.push('Linha ' + numeroLinha + ': vigência inicial é posterior à final.');
    }

    if (encontrado !== 'SIM') {
      situacao = 'NÃO LOCALIZADO NO AUVO';
      if (importar === 'SIM') {
        erros.push('Linha ' + numeroLinha + ': questionário selecionado não foi localizado no AUVO.');
      }
    } else if (importar === 'SIM') {
      if (inicio && inicio.getTime() > hoje.getTime()) {
        situacao = 'AGENDADO PARA ' + formatarDataQuestionarioAuvo_(inicio);
        agendados++;
      } else if (fim && fim.getTime() < hoje.getTime()) {
        situacao = 'ENCERRADO EM ' + formatarDataQuestionarioAuvo_(fim);
        encerrados++;
      } else {
        situacao = 'ATIVO';
        ativos++;
      }
    }
    situacoes.push([situacao]);
  });

  aba.getRange(2, 9, quantidade, 1).setValues(situacoes);
  colorirSituacoesQuestionariosAuvo_(aba, situacoes);

  if (ativos + agendados + encerrados === 0) {
    erros.push('Nenhum questionário foi selecionado para contabilização.');
  }

  if (erros.length) {
    const mensagemErro = erros.slice(0, 12).join(' | ') +
      (erros.length > 12 ? ' | +' + (erros.length - 12) + ' erro(s)' : '');
    registrarLogAuvo_('QUESTIONÁRIOS CONFIGURAÇÃO', 'ERRO', mensagemErro);
    if (mostrarAviso) {
      SpreadsheetApp.getUi().alert('Configuração inválida:\n\n' + mensagemErro);
    }
    throw new Error(mensagemErro);
  }

  const resumo =
    'Configuração válida. Ativos hoje: ' + ativos +
    ' | Agendados: ' + agendados +
    ' | Encerrados: ' + encerrados + '.';
  registrarLogAuvo_('QUESTIONÁRIOS CONFIGURAÇÃO', 'SUCESSO', resumo);
  if (mostrarAviso) SpreadsheetApp.getUi().alert(resumo);
  return { ativos: ativos, agendados: agendados, encerrados: encerrados };
}

function colorirSituacoesQuestionariosAuvo_(aba, situacoes) {
  if (!situacoes.length) return;
  const cores = situacoes.map(function(item) {
    const texto = String(item[0] || '');
    if (texto === 'ATIVO') return ['#D9EAD3'];
    if (texto.indexOf('AGENDADO') === 0) return ['#FFF2CC'];
    if (texto.indexOf('ENCERRADO') === 0) return ['#D9D9D9'];
    if (texto.indexOf('NÃO LOCALIZADO') === 0) return ['#F4CCCC'];
    return ['#FFFFFF'];
  });
  aba.getRange(2, 9, cores.length, 1).setBackgrounds(cores);
}

function normalizarDataConfigQuestionarioAuvo_(valor) {
  if (!valor) return null;
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return isNaN(valor.getTime()) ? null : inicioDoDiaQuestionarioAuvo_(valor);
  }

  const texto = String(valor).trim();
  let partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (partes) {
    const dataBr = new Date(Number(partes[3]), Number(partes[2]) - 1, Number(partes[1]));
    if (
      isNaN(dataBr.getTime()) ||
      dataBr.getFullYear() !== Number(partes[3]) ||
      dataBr.getMonth() !== Number(partes[2]) - 1 ||
      dataBr.getDate() !== Number(partes[1])
    ) return null;
    return inicioDoDiaQuestionarioAuvo_(dataBr);
  }

  partes = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (partes) {
    const dataIso = new Date(Number(partes[1]), Number(partes[2]) - 1, Number(partes[3]));
    if (
      isNaN(dataIso.getTime()) ||
      dataIso.getFullYear() !== Number(partes[1]) ||
      dataIso.getMonth() !== Number(partes[2]) - 1 ||
      dataIso.getDate() !== Number(partes[3])
    ) return null;
    return inicioDoDiaQuestionarioAuvo_(dataIso);
  }
  return null;
}

function inicioDoDiaQuestionarioAuvo_(data) {
  return new Date(data.getFullYear(), data.getMonth(), data.getDate());
}

function formatarDataQuestionarioAuvo_(data) {
  return Utilities.formatDate(
    data,
    AUVO.FUSO_HORARIO,
    'dd/MM/yyyy'
  );
}

/**
 * Contrato usado pela carga mensal e pela atualizacao diaria. Retorna
 * somente as regras selecionadas que se sobrepoem ao periodo solicitado.
 */
function obterQuestionariosContabilizadosAuvo_(dataInicial, dataFinal) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName(AUVO.ABA_CONFIG_QUESTIONARIOS);
  if (!aba || aba.getLastRow() < 2) {
    throw new Error('A configuração de questionários ainda não foi criada.');
  }

  const inicioPeriodo = normalizarDataConfigQuestionarioAuvo_(dataInicial);
  const fimPeriodo = normalizarDataConfigQuestionarioAuvo_(dataFinal);
  if (!inicioPeriodo || !fimPeriodo || inicioPeriodo.getTime() > fimPeriodo.getTime()) {
    throw new Error('Período inválido para contabilização dos questionários.');
  }

  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 9).getValues();
  return valores.filter(function(linha) {
    if (normalizarOpcaoQuestionarioAuvo_(linha[2]) !== 'SIM') return false;
    if (normalizarOpcaoQuestionarioAuvo_(linha[6]) !== 'SIM') return false;
    const inicioRegra = normalizarDataConfigQuestionarioAuvo_(linha[4]);
    const fimRegra = normalizarDataConfigQuestionarioAuvo_(linha[5]);
    return (!inicioRegra || inicioRegra.getTime() <= fimPeriodo.getTime()) &&
      (!fimRegra || fimRegra.getTime() >= inicioPeriodo.getTime());
  }).map(function(linha) {
    return {
      id: String(linha[0]).trim(),
      nome: String(linha[1]).trim(),
      produtividade: normalizarOpcaoQuestionarioAuvo_(linha[3]) === 'SIM',
      inicio: normalizarDataConfigQuestionarioAuvo_(linha[4]),
      fim: normalizarDataConfigQuestionarioAuvo_(linha[5])
    };
  });
}

/**
 * Filtro por atendimento: mesmo que a regra se sobreponha ao mes, cada
 * resposta somente entra se sua data estiver dentro da vigencia configurada.
 */
function questionarioDeveSerContabilizadoAuvo_(idQuestionario, dataAtendimento, regras) {
  const id = String(idQuestionario || '').trim();
  const data = normalizarDataConfigQuestionarioAuvo_(dataAtendimento);
  if (!id || !data) return false;

  return (regras || []).some(function(regra) {
    if (String(regra.id) !== id) return false;
    return (!regra.inicio || regra.inicio.getTime() <= data.getTime()) &&
      (!regra.fim || regra.fim.getTime() >= data.getTime());
  });
}

/* ========================================================================== *
 * CARGA MENSAL E INCREMENTAL DA BASE_QUESTIONARIO
 * ========================================================================== */

function configurarPeriodoQuestionariosAuvo() {
  const ui = SpreadsheetApp.getUi();
  const inicioResposta = ui.prompt(
    'Período dos questionários',
    'Informe a data inicial no formato DD/MM/AAAA:',
    ui.ButtonSet.OK_CANCEL
  );
  if (inicioResposta.getSelectedButton() !== ui.Button.OK) return;

  const fimResposta = ui.prompt(
    'Período dos questionários',
    'Informe a data final no formato DD/MM/AAAA:',
    ui.ButtonSet.OK_CANCEL
  );
  if (fimResposta.getSelectedButton() !== ui.Button.OK) return;

  const inicio = normalizarDataConfigQuestionarioAuvo_(inicioResposta.getResponseText());
  const fim = normalizarDataConfigQuestionarioAuvo_(fimResposta.getResponseText());
  validarPeriodoMensalQuestionariosAuvo_(inicio, fim);

  PropertiesService.getScriptProperties().setProperties({
    QUESTIONARIOS_PERIODO_INICIO: formatarDataIsoQuestionarioAuvo_(inicio),
    QUESTIONARIOS_PERIODO_FIM: formatarDataIsoQuestionarioAuvo_(fim)
  });

  const mensagem = 'Período salvo: ' + formatarDataQuestionarioAuvo_(inicio) +
    ' a ' + formatarDataQuestionarioAuvo_(fim) + '.';
  registrarLogAuvo_('QUESTIONÁRIOS PERÍODO', 'SUCESSO', mensagem);
  ui.alert(mensagem);
}

function validarPeriodoMensalQuestionariosAuvo_(inicio, fim) {
  if (!inicio || !fim || inicio.getTime() > fim.getTime()) {
    throw new Error('Período inválido. Use datas no formato DD/MM/AAAA.');
  }
  if (inicio.getFullYear() !== fim.getFullYear() || inicio.getMonth() !== fim.getMonth()) {
    throw new Error('A carga mensal deve começar e terminar dentro do mesmo mês.');
  }
  const dias = Math.round((fim.getTime() - inicio.getTime()) / 86400000) + 1;
  if (dias > 31) throw new Error('O período não pode ultrapassar 31 dias.');
}

function obterPeriodoConfiguradoQuestionariosAuvo_() {
  const propriedades = PropertiesService.getScriptProperties();
  const inicio = normalizarDataConfigQuestionarioAuvo_(
    propriedades.getProperty('QUESTIONARIOS_PERIODO_INICIO')
  );
  const fim = normalizarDataConfigQuestionarioAuvo_(
    propriedades.getProperty('QUESTIONARIOS_PERIODO_FIM')
  );
  validarPeriodoMensalQuestionariosAuvo_(inicio, fim);
  return { inicio: inicio, fim: fim };
}

function formatarDataIsoQuestionarioAuvo_(data) {
  return Utilities.formatDate(
    data,
    AUVO.FUSO_HORARIO,
    'yyyy-MM-dd'
  );
}

function testarRotaTarefasQuestionariosAuvo() {
  const periodo = obterPeriodoConfiguradoQuestionariosAuvo_();
  const token = autenticarAuvo_();
  const hoje = inicioDoDiaQuestionarioAuvo_(new Date());
  const fimEfetivo = periodo.fim.getTime() > hoje.getTime() ? hoje : periodo.fim;
  if (periodo.inicio.getTime() > fimEfetivo.getTime()) {
    throw new Error('O período configurado ainda não começou.');
  }
  const estilo = descobrirEstiloRotaTarefasAuvo_(
    token, periodo.inicio, fimEfetivo, true
  );
  const amostra = consultarPaginaTarefasAuvo_(
    token, estilo, periodo.inicio, fimEfetivo, 1
  );
  const mensagem = 'Rota validada: ' + estilo + ' | Primeira página: ' +
    amostra.itens.length + ' tarefa(s) | Total informado: ' +
    (amostra.total === null ? 'não informado' : amostra.total) +
    ' | Período testado: ' + formatarDataQuestionarioAuvo_(periodo.inicio) +
    ' a ' + formatarDataQuestionarioAuvo_(fimEfetivo) + '.';
  registrarLogAuvo_('QUESTIONÁRIOS ROTA TAREFAS', 'SUCESSO', mensagem);
  SpreadsheetApp.getUi().alert(mensagem);
}

function montarEndpointTarefasAuvo_(estilo, inicio, fim, pagina) {
  const de = formatarDataIsoQuestionarioAuvo_(inicio);
  const ate = formatarDataIsoQuestionarioAuvo_(fim);
  const basePaginacao = 'page=' + pagina + '&pageSize=100&order=asc';
  let filtro;

  // Fallback compatível com o contrato genérico documentado pela AUVO.
  // A API lista em ordem decrescente e o script aplica o período localmente.
  if (estilo === 'SEM_FILTRO_LOCAL_MINUSCULO') {
    return '/tasks/?paramFilter=' + encodeURIComponent('{}') +
      '&page=' + pagina + '&pageSize=100&order=desc';
  }
  if (estilo === 'SEM_FILTRO_LOCAL_MAIUSCULO') {
    return '/tasks/?ParamFilter=' + encodeURIComponent('{}') +
      '&Page=' + pagina + '&PageSize=100&Order=Desc';
  }

  // Formato oficial documentado pela AUVO para a listagem de tarefas.
  // Nesta rota, endpoint, parametros, propriedades e valores de ordenacao
  // sao sensiveis a maiusculas/minusculas.
  if (estilo === 'OFICIAL_START_END') {
    filtro = {
      StartDate: de + 'T00:00:00',
      EndDate: ate + 'T23:59:59'
    };
    return '/Tasks?ParamFilter=' + encodeURIComponent(JSON.stringify(filtro)) +
      '&Page=' + pagina + '&PageSize=100&Order=Asc';
  }

  if (estilo === 'PARAM_TASK_DATE_EQUAL') {
    filtro = { taskDate: de };
    return '/tasks/?paramFilter=' + encodeURIComponent(JSON.stringify(filtro)) + '&' + basePaginacao;
  }

  if (estilo === 'PARAM_TASK_DATE_FROM_TO') {
    filtro = { taskDateFrom: de + ' 00:00:00', taskDateTo: ate + ' 23:59:59' };
    return '/tasks/?paramFilter=' + encodeURIComponent(JSON.stringify(filtro)) + '&' + basePaginacao;
  }
  if (estilo === 'PARAM_START_END_DATE') {
    filtro = { startDate: de + ' 00:00:00', endDate: ate + ' 23:59:59' };
    return '/tasks/?paramFilter=' + encodeURIComponent(JSON.stringify(filtro)) + '&' + basePaginacao;
  }
  if (estilo === 'PARAM_INITIAL_FINAL_DATE') {
    filtro = { initialDate: de + ' 00:00:00', finalDate: ate + ' 23:59:59' };
    return '/tasks/?paramFilter=' + encodeURIComponent(JSON.stringify(filtro)) + '&' + basePaginacao;
  }
  if (estilo === 'PARAM_DATE_START_END') {
    filtro = { dateStart: de + ' 00:00:00', dateEnd: ate + ' 23:59:59' };
    return '/tasks/?paramFilter=' + encodeURIComponent(JSON.stringify(filtro)) + '&' + basePaginacao;
  }
  if (estilo === 'PARAM_TASK_DATE_START_END') {
    filtro = { taskDateStart: de + ' 00:00:00', taskDateEnd: ate + ' 23:59:59' };
    return '/tasks/?paramFilter=' + encodeURIComponent(JSON.stringify(filtro)) + '&' + basePaginacao;
  }
  if (estilo === 'PARAM_DATE_FROM_TO') {
    filtro = { dateFrom: de + ' 00:00:00', dateTo: ate + ' 23:59:59' };
    return '/tasks/?paramFilter=' + encodeURIComponent(JSON.stringify(filtro)) + '&' + basePaginacao;
  }
  if (estilo === 'QUERY_START_END_DATE') {
    return '/tasks/?startDate=' + encodeURIComponent(de) + '&endDate=' +
      encodeURIComponent(ate) + '&' + basePaginacao;
  }
  if (estilo === 'QUERY_TASK_DATE_FROM_TO') {
    return '/tasks/?taskDateFrom=' + encodeURIComponent(de) + '&taskDateTo=' +
      encodeURIComponent(ate) + '&' + basePaginacao;
  }
  throw new Error('Estilo de rota de tarefas desconhecido: ' + estilo);
}

function descobrirEstiloRotaTarefasAuvo_(token, inicio, fim, forcarTeste) {
  const propriedades = PropertiesService.getScriptProperties();
  const salvo = propriedades.getProperty('AUVO_TAREFAS_FILTRO_ESTILO');
  if (salvo && !forcarTeste) return salvo;

  const estilos = [
    'OFICIAL_START_END',
    'PARAM_TASK_DATE_EQUAL', 'PARAM_TASK_DATE_FROM_TO', 'PARAM_TASK_DATE_START_END',
    'PARAM_START_END_DATE',
    'PARAM_INITIAL_FINAL_DATE', 'PARAM_DATE_START_END',
    'PARAM_DATE_FROM_TO',
    'QUERY_START_END_DATE', 'QUERY_TASK_DATE_FROM_TO',
    'SEM_FILTRO_LOCAL_MINUSCULO', 'SEM_FILTRO_LOCAL_MAIUSCULO'
  ];
  const tentativas = [];

  for (let indice = 0; indice < estilos.length; indice++) {
    try {
      const consulta = consultarPaginaTarefasAuvo_(
        token, estilos[indice], inicio, fim, 1, true
      );
      if (!consulta.itens.length) {
        tentativas.push(estilos[indice] + ': resposta vazia, não validável');
        continue;
      }
      const filtroLocal = estiloUsaFiltroLocalTarefasAuvo_(estilos[indice]);
      const fora = filtroLocal ? [] : consulta.itens.filter(function(tarefa) {
        const data = obterDataTarefaAuvo_(tarefa);
        return !data || data.getTime() < inicio.getTime() || data.getTime() > fim.getTime();
      });
      if (fora.length) {
        tentativas.push(estilos[indice] + ': filtro ignorado (' + fora.length + ' fora do período)');
        continue;
      }
      propriedades.setProperty('AUVO_TAREFAS_FILTRO_ESTILO', estilos[indice]);
      return estilos[indice];
    } catch (erro) {
      tentativas.push(estilos[indice] + ': ' + erro.message);
    }
  }
  throw new Error(
    'Nenhuma variação da rota de tarefas foi validada. ' + tentativas.join(' | ')
  );
}

function consultarPaginaTarefasAuvo_(token, estilo, inicio, fim, pagina, silencioso) {
  const endpoint = montarEndpointTarefasAuvo_(estilo, inicio, fim, pagina);
  const resposta = UrlFetchApp.fetch(AUVO.BASE_URL + endpoint, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json',
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  const http = resposta.getResponseCode();
  if (http < 200 || http >= 300) {
    throw new Error('HTTP ' + http + ' em ' + endpoint + ': ' +
      resposta.getContentText().substring(0, 300));
  }
  let json;
  try {
    json = JSON.parse(resposta.getContentText());
  } catch (erroJson) {
    throw new Error('Resposta inválida da rota de tarefas: ' + erroJson.message);
  }
  const itensBrutos = extrairListaCatalogoAuvo_(json);
  const filtroLocal = estiloUsaFiltroLocalTarefasAuvo_(estilo);
  const itens = filtroLocal ? itensBrutos.filter(function(tarefa) {
    const data = obterDataTarefaAuvo_(tarefa);
    return data && data.getTime() >= inicio.getTime() && data.getTime() <= fim.getTime();
  }) : itensBrutos;
  const atingiuDataAnterior = filtroLocal && itensBrutos.some(function(tarefa) {
    const data = obterDataTarefaAuvo_(tarefa);
    return data && data.getTime() < inicio.getTime();
  });
  const paginacao = extrairPaginacaoCatalogoAuvo_(json);
  const total = paginacao && temValorAuvo_(paginacao.totalItems)
    ? Number(paginacao.totalItems) : null;

  if (!silencioso) {
    itens.forEach(function(tarefa) {
      const data = obterDataTarefaAuvo_(tarefa);
      if (!data || data.getTime() < inicio.getTime() || data.getTime() > fim.getTime()) {
        throw new Error('A API devolveu tarefa fora do período; a carga foi bloqueada.');
      }
    });
  }
  return {
    itens: itens,
    total: isNaN(total) ? null : total,
    endpoint: endpoint,
    quantidadeBruta: itensBrutos.length,
    atingiuDataAnterior: atingiuDataAnterior
  };
}

function estiloUsaFiltroLocalTarefasAuvo_(estilo) {
  return estilo === 'SEM_FILTRO_LOCAL_MINUSCULO' ||
    estilo === 'SEM_FILTRO_LOCAL_MAIUSCULO';
}

function listarTarefasPeriodoQuestionariosAuvo_(token, inicio, fim) {
  const estilo = descobrirEstiloRotaTarefasAuvo_(token, inicio, fim, false);
  const tarefas = [];
  const vistos = {};
  for (let pagina = 1; pagina <= 100; pagina++) {
    const consulta = consultarPaginaTarefasAuvo_(token, estilo, inicio, fim, pagina, false);
    consulta.itens.forEach(function(tarefa) {
      const id = String(tarefa.taskID || tarefa.taskId || tarefa.id || '').trim();
      if (!id || vistos[id]) return;
      vistos[id] = true;
      tarefas.push(tarefa);
    });
    if (estiloUsaFiltroLocalTarefasAuvo_(estilo)) {
      if (!consulta.quantidadeBruta || consulta.quantidadeBruta < 100 || consulta.atingiuDataAnterior) break;
    } else {
      if (!consulta.itens.length || consulta.itens.length < 100) break;
      if (consulta.total !== null && tarefas.length >= consulta.total) break;
    }
    if (pagina === 100) throw new Error('A rota de tarefas excedeu 100 páginas.');
  }
  return completarDetalhesTarefasQuestionariosAuvo_(token, tarefas);
}

/**
 * CORRECAO (bug #4, adicionado em 2026-08-24 apos falha real observada em producao):
 * o estilo de rota de tarefas descoberto em descobrirEstiloRotaTarefasAuvo_ fica em
 * cache nas Script Properties (AUVO_TAREFAS_FILTRO_ESTILO) e, depois de validado uma
 * vez, nunca mais e testado de novo. Em producao a rota OFICIAL_START_END, validada em
 * 07/08, voltou a funcionar para o dia 01/08 mas devolveu HTTP 404 no dia seguinte
 * (02/08) na mesma execucao — uma falha transitoria da API do AUVO bastava para abortar
 * a carga inteira do mes, exigindo reinicio manual. Esta funcao tenta o dia uma vez com
 * o estilo em cache e, se falhar, descarta o cache, redescobre o estilo correto e tenta
 * esse mesmo dia mais uma vez antes de desistir.
 */
function listarTarefasComRetentativaQuestionariosAuvo_(token, dia) {
  try {
    return listarTarefasPeriodoQuestionariosAuvo_(token, dia, dia);
  } catch (erro) {
    registrarLogAuvo_(
      'QUESTIONÁRIOS DIA',
      'AVISO',
      formatarDataQuestionarioAuvo_(dia) + ': primeira tentativa falhou (' +
      erro.message + '). Redescobrindo a rota de tarefas e tentando novamente.'
    );
    PropertiesService.getScriptProperties().deleteProperty('AUVO_TAREFAS_FILTRO_ESTILO');
    descobrirEstiloRotaTarefasAuvo_(token, dia, dia, true);
    return listarTarefasPeriodoQuestionariosAuvo_(token, dia, dia);
  }
}

function completarDetalhesTarefasQuestionariosAuvo_(token, tarefas) {
  const pendentes = tarefas.filter(function(tarefa) {
    return !Array.isArray(tarefa.questionnaires);
  });
  if (!pendentes.length) return tarefas;

  const mapa = {};
  tarefas.forEach(function(tarefa) {
    const id = String(tarefa.taskID || tarefa.taskId || tarefa.id || '').trim();
    if (id) mapa[id] = tarefa;
  });

  for (let inicio = 0; inicio < pendentes.length; inicio += 40) {
    const lote = pendentes.slice(inicio, inicio + 40);
    const requisicoes = lote.map(function(tarefa) {
      const id = String(tarefa.taskID || tarefa.taskId || tarefa.id || '').trim();
      return {
        url: AUVO.BASE_URL + '/tasks/' + encodeURIComponent(id),
        method: 'get',
        headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
        muteHttpExceptions: true
      };
    });
    const respostas = UrlFetchApp.fetchAll(requisicoes);
    respostas.forEach(function(resposta, indice) {
      const http = resposta.getResponseCode();
      const id = String(lote[indice].taskID || lote[indice].taskId || lote[indice].id || '').trim();
      if (http < 200 || http >= 300) {
        throw new Error('Falha ao detalhar tarefa ' + id + '. HTTP ' + http + '.');
      }
      const json = JSON.parse(resposta.getContentText());
      mapa[id] = json && json.result ? json.result : json;
    });
  }
  return Object.keys(mapa).map(function(id) { return mapa[id]; });
}

function carregarMesQuestionariosAuvo() {
  const periodo = obterPeriodoConfiguradoQuestionariosAuvo_();
  const hoje = inicioDoDiaQuestionarioAuvo_(new Date());
  const fimEfetivo = periodo.fim.getTime() > hoje.getTime() ? hoje : periodo.fim;
  if (periodo.inicio.getTime() > fimEfetivo.getTime()) {
    throw new Error('O período configurado ainda não começou.');
  }
  iniciarCargaQuestionariosAuvo_(periodo.inicio, fimEfetivo, 'MENSAL');
}

function obterDiaAmostraRotaTarefasAuvo_(inicio, fim) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUVO.ABA_QUESTIONARIOS);
  let melhor = null;
  if (aba && aba.getLastRow() > 1) {
    aba.getRange(2, 25, aba.getLastRow() - 1, 1).getValues().forEach(function(linha) {
      const data = converterDataCelulaQuestionarioAuvo_(linha[0]);
      if (!data || data.getTime() < inicio.getTime() || data.getTime() > fim.getTime()) return;
      if (!melhor || data.getTime() > melhor.getTime()) melhor = data;
    });
  }
  if (melhor) return melhor;
  const hoje = inicioDoDiaQuestionarioAuvo_(new Date());
  if (hoje.getTime() >= inicio.getTime() && hoje.getTime() <= fim.getTime()) return hoje;
  return fim;
}

function atualizarQuestionariosRecentesAuvo() {
  const hoje = inicioDoDiaQuestionarioAuvo_(new Date());
  const ontem = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 1);
  iniciarCargaQuestionariosAuvo_(ontem, hoje, 'DIÁRIA');
}

/**
 * CORRECAO (bug #1 da revisao 2026-08-24): antes desta versao, uma carga
 * mensal em andamento (13.) podia ser interrompida e ter seu estado
 * sobrescrito se o gatilho diario das 23h (14.) disparasse no meio do
 * processo — a aba temporaria era limpa e as propriedades de controle eram
 * trocadas por baixo do processo em andamento, sem aviso. Agora, assim como
 * atualizarBaseAuvo() ja fazia para AUVO_ATUALIZACAO_ATIVA, verificamos
 * QUESTIONARIOS_CARGA_ATIVA antes de iniciar e desistimos se ja houver uma
 * carga em andamento.
 */
function iniciarCargaQuestionariosAuvo_(inicio, fim, tipo) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Outra integração AUVO está em andamento.');

  try {
    const propriedades = PropertiesService.getScriptProperties();

    if (propriedades.getProperty('QUESTIONARIOS_CARGA_ATIVA') === 'SIM') {
      registrarLogAuvo_(
        'QUESTIONÁRIOS CARGA',
        'AVISO',
        'Nova carga (' + tipo + ') ignorada: já existe uma carga de questionários em andamento.'
      );
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Já existe uma carga de questionários em andamento. Aguarde terminar (ou cancele) antes de iniciar outra.',
        'Integração AUVO',
        10
      );
      return;
    }

    validarConfiguracaoQuestionariosAuvo_(false);
    const regras = obterQuestionariosContabilizadosAuvo_(inicio, fim);
    if (!regras.length) throw new Error('Nenhum questionário possui vigência no período.');
    const token = autenticarAuvo_();
    const diaAmostra = obterDiaAmostraRotaTarefasAuvo_(inicio, fim);
    descobrirEstiloRotaTarefasAuvo_(token, diaAmostra, diaAmostra, false);
    prepararTemporarioQuestionariosAuvo_();
    removerGatilhosContinuacaoQuestionariosAuvo_();
    propriedades.setProperties({
      QUESTIONARIOS_CARGA_ATIVA: 'SIM',
      QUESTIONARIOS_CARGA_TIPO: tipo,
      QUESTIONARIOS_CARGA_INICIO: formatarDataIsoQuestionarioAuvo_(inicio),
      QUESTIONARIOS_CARGA_FIM: formatarDataIsoQuestionarioAuvo_(fim),
      QUESTIONARIOS_CARGA_DIA: formatarDataIsoQuestionarioAuvo_(inicio)
    });
    registrarLogAuvo_(
      'QUESTIONÁRIOS CARGA', 'INÍCIO', tipo + ' | ' +
      formatarDataQuestionarioAuvo_(inicio) + ' a ' + formatarDataQuestionarioAuvo_(fim)
    );
  } finally {
    lock.releaseLock();
  }
  processarCargaQuestionariosAuvo_();
}

function continuarCargaQuestionariosAuvo() {
  processarCargaQuestionariosAuvo_();
}

/**
 * CORRECAO (bug #2 da revisao 2026-08-24): quando o lock nao era adquirido,
 * a funcao apenas retornava sem reagendar a continuacao — uma disputa de
 * lock passageira (por exemplo, colidindo com a atualizacao de clientes)
 * travava silenciosamente a cadeia de continuacao de uma carga em andamento
 * ate alguem perceber e reexecutar manualmente. Agora reagenda a
 * continuacao, no mesmo padrao usado por processarPaginasAuvo_() para a
 * BASE_AUVO.
 */
function processarCargaQuestionariosAuvo_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    agendarContinuacaoQuestionariosAuvo_();
    return;
  }
  const inicioExecucao = Date.now();
  try {
    const propriedades = PropertiesService.getScriptProperties();
    if (propriedades.getProperty('QUESTIONARIOS_CARGA_ATIVA') !== 'SIM') return;
    const inicioPeriodo = normalizarDataConfigQuestionarioAuvo_(
      propriedades.getProperty('QUESTIONARIOS_CARGA_INICIO')
    );
    const fimPeriodo = normalizarDataConfigQuestionarioAuvo_(
      propriedades.getProperty('QUESTIONARIOS_CARGA_FIM')
    );
    let dia = normalizarDataConfigQuestionarioAuvo_(
      propriedades.getProperty('QUESTIONARIOS_CARGA_DIA')
    );
    const regras = obterQuestionariosContabilizadosAuvo_(inicioPeriodo, fimPeriodo);
    const token = autenticarAuvo_();

    while (dia.getTime() <= fimPeriodo.getTime() && Date.now() - inicioExecucao < 210000) {
      const tarefas = listarTarefasComRetentativaQuestionariosAuvo_(token, dia);
      const objetos = transformarTarefasQuestionariosAuvo_(tarefas, regras);
      gravarObjetosTemporariosQuestionariosAuvo_(dia, objetos);
      registrarLogAuvo_(
        'QUESTIONÁRIOS DIA', 'SUCESSO', formatarDataQuestionarioAuvo_(dia) +
        ': ' + tarefas.length + ' tarefa(s), ' + objetos.length + ' linha(s).'
      );
      dia = new Date(dia.getFullYear(), dia.getMonth(), dia.getDate() + 1);
      propriedades.setProperty('QUESTIONARIOS_CARGA_DIA', formatarDataIsoQuestionarioAuvo_(dia));
    }

    if (dia.getTime() <= fimPeriodo.getTime()) {
      agendarContinuacaoQuestionariosAuvo_();
      return;
    }
    finalizarCargaQuestionariosAuvo_(inicioPeriodo, fimPeriodo);
  } catch (erro) {
    registrarLogAuvo_('QUESTIONÁRIOS CARGA', 'ERRO', erro.message);
    encerrarEstadoCargaQuestionariosAuvo_();
    throw erro;
  } finally {
    lock.releaseLock();
  }
}

function prepararTemporarioQuestionariosAuvo_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(AUVO.ABA_TEMP_QUESTIONARIOS);
  if (!aba) aba = planilha.insertSheet(AUVO.ABA_TEMP_QUESTIONARIOS);
  aba.clearContents();
  aba.getRange(1, 1, 1, 2).setValues([['Data da tarefa', 'Objeto JSON']]);
  aba.hideSheet();
}

function gravarObjetosTemporariosQuestionariosAuvo_(dia, objetos) {
  if (!objetos.length) return;
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUVO.ABA_TEMP_QUESTIONARIOS);
  const linhas = objetos.map(function(objeto) {
    return [formatarDataIsoQuestionarioAuvo_(dia), JSON.stringify(objeto)];
  });
  aba.getRange(aba.getLastRow() + 1, 1, linhas.length, 2).setValues(linhas);
}

function transformarTarefasQuestionariosAuvo_(tarefas, regras) {
  const regrasPorId = {};
  regras.forEach(function(regra) { regrasPorId[String(regra.id)] = regra; });
  const documentos = criarMapaDocumentosClientesAuvo_();
  const objetos = [];

  tarefas.forEach(function(tarefa) {
    const dataTarefa = obterDataTarefaAuvo_(tarefa);
    if (!dataTarefa || !Array.isArray(tarefa.questionnaires)) return;
    tarefa.questionnaires.forEach(function(questionario) {
      const id = String(questionario.questionnaireId || questionario.id || '').trim();
      const regra = regrasPorId[id];
      if (!Array.isArray(questionario.answers) || !questionario.answers.length) return;
      if (!regra || !questionarioDeveSerContabilizadoAuvo_(id, dataTarefa, regras)) return;
      objetos.push(montarObjetoQuestionarioAuvo_(tarefa, questionario, regra, documentos));
    });
  });
  return objetos;
}

function montarObjetoQuestionarioAuvo_(tarefa, questionario, regra, documentos) {
  const respostas = Array.isArray(questionario.answers) ? questionario.answers : [];
  const agrupadas = {};
  respostas.forEach(function(resposta) {
    const id = String(resposta.questionId || '').trim();
    if (!id) return;
    if (!agrupadas[id]) {
      agrupadas[id] = { id: id, descricao: resposta.questionDescription || '', valores: [], datas: [], replyIds: [] };
    }
    const valor = valorAuvo_(resposta.reply);
    if (temValorAuvo_(valor) && agrupadas[id].valores.indexOf(String(valor)) < 0) {
      agrupadas[id].valores.push(String(valor));
    }
    if (temValorAuvo_(resposta.replyDate)) agrupadas[id].datas.push(String(resposta.replyDate));
    if (temValorAuvo_(resposta.replyId)) agrupadas[id].replyIds.push(String(resposta.replyId));
  });

  const legado = new Array(15).fill('');
  const extras = {};
  Object.keys(agrupadas).forEach(function(id) {
    const item = agrupadas[id];
    const valor = item.valores.join(';');
    const indiceLegado = obterIndiceRespostaLegadaQuestionarioAuvo_(item.descricao);
    if (indiceLegado >= 0) legado[indiceLegado] = valor;
    else extras[id] = { descricao: item.descricao, valor: valor };
  });

  const dataRespostaTexto = obterMaiorDataRespostaQuestionarioAuvo_(agrupadas);
  const dataTarefa = obterDataTarefaAuvo_(tarefa);
  const idTarefa = String(tarefa.taskID || tarefa.taskId || tarefa.id || '').trim();
  const idQuestionario = String(questionario.questionnaireId || questionario.id || '').trim();
  const clienteId = String(tarefa.customerId || '').trim();
  const assinatura = criarAssinaturaInstanciaQuestionarioAuvo_(questionario, agrupadas);
  const chave = idTarefa + '|' + idQuestionario + '|' + assinatura;
  const documento = documentos[clienteId] || '';

  const fixa = new Array(43).fill('');
  fixa[0] = idTarefa;
  fixa[1] = valorAuvo_(tarefa.userToName);
  fixa[2] = valorAuvo_(tarefa.customerDescription);
  fixa[3] = regra.produtividade ? clienteId : '';
  for (let i = 0; i < 15; i++) fixa[4 + i] = legado[i];
  fixa[19] = dataRespostaTexto || '';
  fixa[20] = dataRespostaTexto || '';
  fixa[21] = valorAuvo_(questionario.questionnaireEquipamentId);
  fixa[23] = valorAuvo_(tarefa.externalId);
  fixa[24] = formatarDataIsoQuestionarioAuvo_(dataTarefa);
  fixa[25] = documento;
  fixa[37] = idQuestionario;
  fixa[38] = valorAuvo_(questionario.questionnaireDescription || regra.nome);
  fixa[39] = chave;
  fixa[40] = regra.produtividade ? 'SIM' : 'NÃO';
  fixa[41] = clienteId;
  fixa[42] = assinatura;
  return { fixa: fixa, extras: extras };
}

function obterIndiceRespostaLegadaQuestionarioAuvo_(descricao) {
  const texto = normalizarTextoQuestionarioAuvo_(descricao);
  if (/modelo.*equipamento/.test(texto)) return 0;
  if (/codigo.*patrimonio/.test(texto)) return 1;
  if (/foto.*patrimonio|qr.*code/.test(texto)) return 2;
  if (/foto.*ambiente.*antes|equipamento.*antes.*manutencao/.test(texto)) return 3;
  if (/local.*instalacao|local.*manutencao/.test(texto)) return 4;
  if (/foto.*refil.*antes/.test(texto)) return 5;
  if (/volume.*refil.*antes/.test(texto)) return 6;
  if (/volume.*fragrancia.*acima|motivo.*consumo/.test(texto)) return 7;
  if (/servicos?.*executad/.test(texto)) return 8;
  if (/foto.*pilha/.test(texto)) return 9;
  if (/foto.*configuracao|potencia.*programacao/.test(texto)) return 10;
  if (/situacao.*trava|trava.*equipamento/.test(texto)) return 11;
  if (/frasco.*lote|lote.*fragrancia/.test(texto)) return 12;
  if (/refil.*depois|limpeza.*equipamento/.test(texto)) return 13;
  if (/observa/.test(texto)) return 14;
  return -1;
}

function normalizarTextoQuestionarioAuvo_(valor) {
  return String(valor || '').normalize('NFD').replace(/\p{Mn}/gu, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function obterMaiorDataRespostaQuestionarioAuvo_(agrupadas) {
  let maior = '';
  Object.keys(agrupadas).forEach(function(id) {
    agrupadas[id].datas.forEach(function(data) {
      if (!maior || String(data) > maior) maior = String(data);
    });
  });
  return maior;
}

function criarAssinaturaInstanciaQuestionarioAuvo_(questionario, agrupadas) {
  const equipamentoId = String(questionario.questionnaireEquipamentId || '').trim();
  if (equipamentoId && equipamentoId !== '0') return 'EQUIP:' + equipamentoId;

  const ids = Object.keys(agrupadas);
  for (let indice = 0; indice < ids.length; indice++) {
    const item = agrupadas[ids[indice]];
    const texto = normalizarTextoQuestionarioAuvo_(item.descricao);
    if (/codigo.*patrimonio|numero.*serie|identificador.*equipamento/.test(texto) && item.valores.length) {
      return 'IDENT:' + normalizarTextoQuestionarioAuvo_(item.valores[0]);
    }
  }
  const replyIds = [];
  ids.forEach(function(id) { replyIds.push.apply(replyIds, agrupadas[id].replyIds); });
  replyIds.sort(function(a, b) { return Number(a) - Number(b); });
  return replyIds.length ? 'REPLY:' + replyIds[0] :
    'SEMRESPOSTA:' + String(questionario.questionnaireDescription || 'questionario');
}

function criarMapaDocumentosClientesAuvo_() {
  const mapa = {};
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUVO.ABA_CLIENTES);
  if (!aba || aba.getLastRow() < 2) return mapa;
  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, AUVO.TOTAL_COLUNAS).getDisplayValues();
  const indiceCodigo = obterIndiceCabecalhoAuvo_('Código');
  const indiceDocumento = obterIndiceCabecalhoAuvo_('CPF ou CNPJ');
  valores.forEach(function(linha) {
    const id = String(linha[indiceCodigo] || '').trim();
    if (id) mapa[id] = String(linha[indiceDocumento] || '').trim();
  });
  return mapa;
}

function obterDataTarefaAuvo_(tarefa) {
  const valor = tarefa && (tarefa.taskDate || tarefa.date || tarefa.scheduledDate);
  if (!valor) return null;
  const texto = String(valor).substring(0, 10);
  return normalizarDataConfigQuestionarioAuvo_(texto);
}

function finalizarCargaQuestionariosAuvo_(inicio, fim) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const temporaria = planilha.getSheetByName(AUVO.ABA_TEMP_QUESTIONARIOS);
  const objetos = [];
  if (temporaria && temporaria.getLastRow() > 1) {
    temporaria.getRange(2, 2, temporaria.getLastRow() - 1, 1).getValues()
      .forEach(function(linha) { if (linha[0]) objetos.push(JSON.parse(linha[0])); });
  }
  gravarBaseQuestionariosAuvo_(objetos, inicio, fim);
  registrarLogAuvo_(
    'QUESTIONÁRIOS CARGA', 'SUCESSO', objetos.length + ' linha(s) gravadas no período ' +
    formatarDataQuestionarioAuvo_(inicio) + ' a ' + formatarDataQuestionarioAuvo_(fim) + '.'
  );
  encerrarEstadoCargaQuestionariosAuvo_();
  planilha.toast(objetos.length + ' questionário(s) atualizado(s).', 'Integração AUVO', 15);
}

function gravarBaseQuestionariosAuvo_(objetos, inicio, fim) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(AUVO.ABA_QUESTIONARIOS);
  if (!aba) aba = planilha.insertSheet(AUVO.ABA_QUESTIONARIOS);
  criarBackupInicialBaseQuestionariosAuvo_(aba);

  const ultimaColunaExistente = Math.max(aba.getLastColumn(), 37);
  const cabecalhosExistentes = aba.getRange(1, 1, 1, ultimaColunaExistente).getDisplayValues()[0];
  const cabecalhos = CABECALHOS_BASE_QUESTIONARIO_.slice();
  cabecalhosExistentes.slice(43).forEach(function(cabecalho) {
    if (cabecalho && cabecalhos.indexOf(cabecalho) < 0) cabecalhos.push(cabecalho);
  });
  objetos.forEach(function(objeto) {
    Object.keys(objeto.extras || {}).forEach(function(id) {
      const titulo = 'Q_' + id + ' | ' + objeto.extras[id].descricao;
      if (cabecalhos.indexOf(titulo) < 0) cabecalhos.push(titulo);
    });
  });

  const indiceCabecalhoExistente = {};
  cabecalhosExistentes.forEach(function(item, indice) { if (item) indiceCabecalhoExistente[item] = indice; });
  const mantidas = [];
  if (aba.getLastRow() > 1) {
    const existentes = aba.getRange(2, 1, aba.getLastRow() - 1, ultimaColunaExistente).getValues();
    existentes.forEach(function(linha) {
      const data = converterDataCelulaQuestionarioAuvo_(linha[24]);
      if (data && data.getTime() >= inicio.getTime() && data.getTime() <= fim.getTime()) return;
      if (!linha.some(function(valor) { return valor !== '' && valor !== null; })) return;
      const nova = new Array(cabecalhos.length).fill('');
      cabecalhos.forEach(function(cabecalho, indice) {
        if (Object.prototype.hasOwnProperty.call(indiceCabecalhoExistente, cabecalho)) {
          nova[indice] = linha[indiceCabecalhoExistente[cabecalho]];
        } else if (indice < linha.length && indice < 43) {
          nova[indice] = linha[indice];
        }
      });
      migrarLinhaLegadaQuestionarioAuvo_(nova);
      mantidas.push(nova);
    });
  }

  const importadas = objetos.map(function(objeto) {
    const linha = new Array(cabecalhos.length).fill('');
    objeto.fixa.forEach(function(valor, indice) {
      linha[indice] = converterValorImportadoQuestionarioAuvo_(indice, valor);
    });
    Object.keys(objeto.extras || {}).forEach(function(id) {
      const titulo = 'Q_' + id + ' | ' + objeto.extras[id].descricao;
      linha[cabecalhos.indexOf(titulo)] = objeto.extras[id].valor;
    });
    return linha;
  });
  const todas = mantidas.concat(importadas);
  todas.sort(function(a, b) {
    const da = converterDataCelulaQuestionarioAuvo_(a[24]);
    const db = converterDataCelulaQuestionarioAuvo_(b[24]);
    const diferenca = (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
    return diferenca || String(a[0]).localeCompare(String(b[0]), 'pt-BR', { numeric: true }) ||
      String(a[39]).localeCompare(String(b[39]));
  });

  const linhasNecessarias = Math.max(todas.length + 1, 2);
  const colunasNecessarias = cabecalhos.length;
  if (aba.getMaxRows() < linhasNecessarias) {
    aba.insertRowsAfter(aba.getMaxRows(), linhasNecessarias - aba.getMaxRows());
  }
  if (aba.getMaxColumns() < colunasNecessarias) {
    aba.insertColumnsAfter(aba.getMaxColumns(), colunasNecessarias - aba.getMaxColumns());
  }
  const limparLinhas = Math.max(aba.getLastRow(), linhasNecessarias) - 1;
  const limparColunas = Math.max(aba.getLastColumn(), colunasNecessarias);
  if (limparLinhas > 0) aba.getRange(2, 1, limparLinhas, limparColunas).clearContent();
  aba.getRange(1, 1, 1, colunasNecessarias).setValues([cabecalhos]);
  if (todas.length) aba.getRange(2, 1, todas.length, colunasNecessarias).setValues(todas);
  aplicarFormulasBaseQuestionariosAuvo_(aba, todas.length);
  formatarBaseQuestionariosAuvo_(aba, todas.length, colunasNecessarias);
}

function migrarLinhaLegadaQuestionarioAuvo_(linha) {
  if (!temValorAuvo_(linha[0]) || temValorAuvo_(linha[37])) return;
  linha[37] = '318045';
  linha[38] = '04. Relatório de Atendimento v.2';
  linha[40] = 'SIM';
  linha[41] = linha[3];
  const patrimonio = normalizarTextoQuestionarioAuvo_(linha[5]);
  linha[42] = patrimonio ? 'IDENT:' + patrimonio : 'LEGADO:' + String(linha[0]);
  linha[39] = String(linha[0]) + '|318045|' + linha[42];
}

function converterValorImportadoQuestionarioAuvo_(indice, valor) {
  if ((indice === 19 || indice === 20) && valor) {
    const dataResposta = new Date(String(valor).replace(' ', 'T'));
    return isNaN(dataResposta.getTime()) ? valor : dataResposta;
  }
  if (indice === 24 && valor) {
    const dataTarefa = normalizarDataConfigQuestionarioAuvo_(valor);
    return dataTarefa || valor;
  }
  return valor;
}

function converterDataCelulaQuestionarioAuvo_(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return inicioDoDiaQuestionarioAuvo_(valor);
  }
  return normalizarDataConfigQuestionarioAuvo_(valor);
}

/**
 * CORRECAO (bug #3 da revisao 2026-08-24): a formula da coluna "Chave Tec
 * Data" referenciava RC26 ("CPF/CNPJ do Cliente", texto), quando deveria
 * referenciar RC25 ("Data da tarefa") — aplicar TEXT(...,"dd/mm/yyyy") sobre
 * um CPF/CNPJ nao produz uma chave tecnico+data valida.
 */
function aplicarFormulasBaseQuestionariosAuvo_(aba, quantidade) {
  if (!quantidade) return;
  aba.getRange(2, 33, quantidade, 5).setFormulasR1C1(
    Array.from({ length: quantidade }, function() {
      return [
        '=IF(RC41<>"SIM","EXCLUIR",IFERROR(IF(OR(ISNUMBER(SEARCH("BRASILIA",UPPER(IFERROR(INDEX(Grupos!C7,MATCH(TEXT(RC4,"0"),Grupos!C1,0)),"")&" "&IFERROR(INDEX(Grupos!C8,MATCH(TEXT(RC4,"0"),Grupos!C1,0)),"")))),ISNUMBER(SEARCH("CANCELADOS",UPPER(IFERROR(INDEX(Grupos!C7,MATCH(TEXT(RC4,"0"),Grupos!C1,0)),"")&" "&IFERROR(INDEX(Grupos!C8,MATCH(TEXT(RC4,"0"),Grupos!C1,0)),""))))),"EXCLUIR","OK"),"OK"))',
        '=IF(RC33<>"OK","",TEXT(RC4,"0"))',
        '=IF(RC33<>"OK","",IF(COUNTIFS(R2C2:RC2,RC2,R2C4:RC4,RC4,R2C33:RC33,"OK")=1,RC2&"|"&TEXT(RC4,"0"),""))',
        '=IF(RC33<>"OK","",IF(COUNTIFS(R2C2:RC2,RC2,R2C25:RC25,RC25,R2C33:RC33,"OK")=1,RC2&"|"&TEXT(RC25,"dd/mm/yyyy"),""))',
        '=IF(RC33<>"OK","",IF(COUNTIFS(R2C4:RC4,RC4,R2C33:RC33,"OK")=1,TEXT(RC4,"0"),""))'
      ];
    })
  );
}

function formatarBaseQuestionariosAuvo_(aba, quantidade, colunas) {
  aba.setFrozenRows(1);
  aba.getRange(1, 1, 1, colunas).setFontWeight('bold').setBackground('#1F4E78')
    .setFontColor('#FFFFFF').setWrap(true);
  if (quantidade) {
    aba.getRange(2, 20, quantidade, 1).setNumberFormat('dd/MM/yyyy');
    aba.getRange(2, 21, quantidade, 1).setNumberFormat('HH:mm:ss');
    aba.getRange(2, 25, quantidade, 1).setNumberFormat('dd/MM/yyyy');
  }
  aba.setColumnWidth(1, 110);
  aba.setColumnWidth(2, 210);
  aba.setColumnWidth(3, 300);
  aba.setColumnWidth(38, 130);
  aba.setColumnWidth(39, 300);
  aba.setColumnWidth(40, 320);
  aba.hideColumns(38, 6);
  const filtro = aba.getFilter();
  if (filtro) filtro.remove();
  aba.getRange(1, 1, Math.max(quantidade + 1, 2), colunas).createFilter();
}

function criarBackupInicialBaseQuestionariosAuvo_(aba) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  if (planilha.getSheetByName('BACKUP_BASE_QUESTIONARIO')) return;
  const copia = aba.copyTo(planilha);
  copia.setName('BACKUP_BASE_QUESTIONARIO');
  copia.hideSheet();
}

function agendarContinuacaoQuestionariosAuvo_() {
  removerGatilhosContinuacaoQuestionariosAuvo_();
  ScriptApp.newTrigger('continuarCargaQuestionariosAuvo').timeBased().after(60000).create();
}

function removerGatilhosContinuacaoQuestionariosAuvo_() {
  ScriptApp.getProjectTriggers().forEach(function(gatilho) {
    if (gatilho.getHandlerFunction() === 'continuarCargaQuestionariosAuvo') {
      ScriptApp.deleteTrigger(gatilho);
    }
  });
}

function encerrarEstadoCargaQuestionariosAuvo_() {
  removerGatilhosContinuacaoQuestionariosAuvo_();
  PropertiesService.getScriptProperties().deleteProperty('QUESTIONARIOS_CARGA_ATIVA');
  PropertiesService.getScriptProperties().deleteProperty('QUESTIONARIOS_CARGA_TIPO');
  PropertiesService.getScriptProperties().deleteProperty('QUESTIONARIOS_CARGA_INICIO');
  PropertiesService.getScriptProperties().deleteProperty('QUESTIONARIOS_CARGA_FIM');
  PropertiesService.getScriptProperties().deleteProperty('QUESTIONARIOS_CARGA_DIA');
}

function instalarAtualizacaoDiariaQuestionariosAuvo() {
  removerAtualizacaoDiariaQuestionariosAuvo_(false);
  ScriptApp.newTrigger('atualizarQuestionariosRecentesAuvo').timeBased().everyDays(1).atHour(23).create();
  registrarLogAuvo_('QUESTIONÁRIOS GATILHO', 'SUCESSO', 'Atualização diária instalada para 23h.');
  SpreadsheetApp.getUi().alert('Atualização diária dos questionários instalada para 23h.');
}

function removerAtualizacaoDiariaQuestionariosAuvo() {
  removerAtualizacaoDiariaQuestionariosAuvo_(true);
}

function removerAtualizacaoDiariaQuestionariosAuvo_(mostrarAviso) {
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(function(gatilho) {
    if (gatilho.getHandlerFunction() === 'atualizarQuestionariosRecentesAuvo') {
      ScriptApp.deleteTrigger(gatilho);
      removidos++;
    }
  });
  if (mostrarAviso) SpreadsheetApp.getUi().alert(removidos + ' gatilho(s) removido(s).');
}
