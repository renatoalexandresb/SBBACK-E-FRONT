/**
 * SERVE O HTML PRINCIPAL DA APLICAÇÃO
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Formulario');
  template.modoConsultaGestores = !!(e && e.parameter && String(e.parameter.view || '').toLowerCase() === 'calendario');
  return template.evaluate()
    .setTitle(template.modoConsultaGestores ? 'Calendário de Experiência - Consulta' : 'ERP RH - Módulo de Gestão')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * GERA O LINK DE CONSULTA SOMENTE LEITURA PARA OS GESTORES.
 * O link usa a mesma implantação do Web App e abre diretamente o modo calendário.
 */
function obterLinkConsultaGestores() {
  var url = ScriptApp.getService().getUrl();
  if (!url) throw new Error('Não foi possível obter a URL da implantação do Web App.');
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + 'view=calendario';
}

/**
 * DADOS MÍNIMOS PARA A PÁGINA DE CONSULTA DOS GESTORES.
 * Não retorna CPF, e-mail, celular ou demais dados administrativos.
 */
function buscarCalendarioConsulta() {
  try {
    var resposta = buscarAvaliacaoExperiencia();
    if (!resposta || !resposta.sucesso) return resposta;
    var dadosSeguros = (resposta.dados || []).map(function(item) {
      return {
        nome: item.nome || '',
        cargo: item.cargo || '',
        loja: item.loja || 'Sem Loja',
        supervisor: item.supervisor || 'Sem Supervisor',
        dataAdmissao: item.dataAdmissao || '',
        dataAdmissaoISO: item.dataAdmissaoISO || '',
        venc15: item.venc15 || '', venc15ISO: item.venc15ISO || '', status15: item.status15 || '',
        venc30: item.venc30 || '', venc30ISO: item.venc30ISO || '', status30: item.status30 || '',
        venc45: item.venc45 || '', venc45ISO: item.venc45ISO || '', status45: item.status45 || '',
        venc60: item.venc60 || '', venc60ISO: item.venc60ISO || '', status60: item.status60 || '',
        venc75: item.venc75 || '', venc75ISO: item.venc75ISO || '', status75: item.status75 || '',
        venc90: item.venc90 || '', venc90ISO: item.venc90ISO || '', status90: item.status90 || ''
      };
    });
    return { sucesso: true, dados: dadosSeguros };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao carregar calendário de consulta: ' + e.message };
  }
}

/**
 * MOTIVOS DE DESLIGAMENTO PERMITIDOS
 */
var MOTIVOS_DESLIGAMENTO_PERMITIDOS = [
  "Sem justa causa",
  "Por justa causa",
  "Pedido de demissão",
  "Demissão consensual",
  "Rescisão indireta"
];

/**
 * LOCALIZA A ABA PRINCIPAL DA BASE DE COLABORADORES.
 * A base atual usa o nome "Colaboradores". Mantemos "Funcionários" como
 * compatibilidade com versões antigas, sem depender da aba ativa.
 */
function obterAbaColaboradores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Colaboradores");
  if (aba) return aba;
  aba = ss.getSheetByName("Funcionários");
  if (aba) return aba;
  throw new Error('A aba principal "Colaboradores" não foi encontrada.');
}

function normalizarMotivoDesligamento(valor) {
  var texto = String(valor == null ? "" : valor).trim();
  if (!texto) return "";
  var normalizado = normalizarTexto(texto);
  for (var i = 0; i < MOTIVOS_DESLIGAMENTO_PERMITIDOS.length; i++) {
    if (normalizarTexto(MOTIVOS_DESLIGAMENTO_PERMITIDOS[i]) === normalizado) {
      return MOTIVOS_DESLIGAMENTO_PERMITIDOS[i];
    }
  }
  return "";
}

function pareceTelefone(valor) {
  var numeros = String(valor == null ? "" : valor).replace(/\D/g, "");
  return numeros.length >= 10 && numeros.length <= 13;
}

/**
 * CORRIGE REGISTROS LEGADOS EM QUE UM TELEFONE FOI GRAVADO EM P.
 * E continua sendo a coluna oficial de Celular e P a de Motivo.
 */
function corrigirInconsistenciasContatoDesligamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = obterAbaColaboradores();
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return 0;

  var dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
  var corrigidos = 0;

  for (var i = 0; i < dados.length; i++) {
    var celular = String(dados[i][4] == null ? "" : dados[i][4]).trim();
    var motivo = String(dados[i][15] == null ? "" : dados[i][15]).trim();
    if (!motivo) continue;

    var motivoValido = normalizarMotivoDesligamento(motivo);
    if (motivoValido) {
      if (motivo !== motivoValido) {
        aba.getRange(i + 2, 16).setValue(motivoValido);
        corrigidos++;
      }
      continue;
    }

    // Caso clássico do erro: telefone em P. Se E estiver vazio, recupera o telefone.
    if (pareceTelefone(motivo)) {
      if (!celular) {
        aba.getRange(i + 2, 5).setValue(motivo);
      }
      aba.getRange(i + 2, 16).clearContent();
      corrigidos++;
    }
  }

  if (corrigidos) SpreadsheetApp.flush();
  return corrigidos;
}

/**
 * SALVA UM COLABORADOR INDIVIDUALMENTE (CADASTRO MANUAL)
 */
function salvarFuncionario(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaColaboradores();
    
    // Formata datas
    var dtNasc = dados.dataNascimento ? formatarDataISO(dados.dataNascimento) : "";
    var dtContr = dados.dataContratacao ? formatarDataISO(dados.dataContratacao) : "";
    var dtDeslig = dados.dataDesligamento ? formatarDataISO(dados.dataDesligamento) : "";
    var motivoDesligamento = normalizarMotivoDesligamento(dados.tipoDesligamento);

    aba.appendRow([dados.nome,dados.email,dados.cpf,dados.sexo,dados.celular,dtNasc,dtContr,dados.formacao,dados.cargo,dados.tags,dados.unidadeGeografica,dados.unidadeNegocio,dados.grupos,dados.lideres,dtDeslig,motivoDesligamento,dados.bloquear,"Pendente","Pendente"]);
    registrarMovimentacao({cpf:dados.cpf,nome:dados.nome,dataMovimentacao:dtContr||new Date(),tipoMovimentacao:"Admissão",cargoAnterior:"",cargoNovo:dados.cargo,lojaAnterior:"",lojaNova:dados.unidadeNegocio,supervisorAnterior:"",supervisorNovo:dados.lideres,origem:"Cadastro Manual",observacao:"Registro inicial do colaborador"});

    return { sucesso: true, mensagem: "Colaborador cadastrado com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao salvar no servidor: " + e.message };
  }
}

/**
 * ATUALIZA UM REGISTRO EXISTENTE (MODAL DE EDIÇÃO)
 */
function atualizarFuncionario(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaColaboradores();
    var linha = parseInt(dados.linhaPlanilha);

    if (!linha || linha < 2) throw new Error("Linha inválida para atualização.");

    var dtNasc = dados.dataNascimento ? formatarDataISO(dados.dataNascimento) : "";
    var dtContr = dados.dataContratacao ? formatarDataISO(dados.dataContratacao) : "";
    var dtDeslig = dados.dataDesligamento ? formatarDataISO(dados.dataDesligamento) : "";
    var motivoDesligamento = normalizarMotivoDesligamento(dados.tipoDesligamento);
    var celularAtualizado = String(dados.celular == null ? "" : dados.celular).trim();
    // Proteção contra o erro legado: telefone nunca deve ser gravado em P.
    if (!motivoDesligamento && pareceTelefone(dados.tipoDesligamento)) {
      if (!celularAtualizado) celularAtualizado = String(dados.tipoDesligamento).trim();
      motivoDesligamento = "";
    }
    var anterior = aba.getRange(linha,1,1,19).getValues()[0];
    var cargoAnterior=anterior[8]||"", lojaAnterior=anterior[11]||"", supervisorAnterior=anterior[13]||"";

    aba.getRange(linha, 1, 1, 17).setValues([[dados.nome,dados.email,dados.cpf,dados.sexo,celularAtualizado,dtNasc,dtContr,dados.formacao,dados.cargo,dados.tags,dados.unidadeGeografica,dados.unidadeNegocio,dados.grupos,dados.lideres,dtDeslig,motivoDesligamento,dados.bloquear]]);

    var mudouCargo=normalizarTexto(cargoAnterior)!==normalizarTexto(dados.cargo);
    var mudouLoja=normalizarTexto(lojaAnterior)!==normalizarTexto(dados.unidadeNegocio);
    var mudouSupervisor=normalizarTexto(supervisorAnterior)!==normalizarTexto(dados.lideres);
    if(mudouCargo||mudouLoja||mudouSupervisor){
      if(dados.tipoMovimentacao){
        registrarMovimentacao({cpf:dados.cpf||anterior[2],nome:dados.nome||anterior[0],dataMovimentacao:dados.dataMovimentacao||new Date(),tipoMovimentacao:dados.tipoMovimentacao,cargoAnterior:cargoAnterior,cargoNovo:dados.cargo,lojaAnterior:lojaAnterior,lojaNova:dados.unidadeNegocio,supervisorAnterior:supervisorAnterior,supervisorNovo:dados.lideres,origem:"Edição Manual",observacao:dados.observacaoMovimentacao||""});
      }
    }
    return { sucesso: true, mensagem: "Dados atualizados com sucesso!", movimentoDetectado:(mudouCargo||mudouLoja||mudouSupervisor) };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao atualizar: " + e.message };
  }
}

/**
 * BUSCA TODOS OS COLABORADORES ATIVOS PARA A TABELA DE CONSULTA
 */
function buscarQuadroAtivo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaColaboradores();
    var dados = aba.getDataRange().getValues();
    
    if (dados.length <= 1) return { sucesso: true, dados: [] };

    var lista = [];
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var estaBloqueado = String(linha[16]).trim().toLowerCase() === "sim";
      var temDesligamento = linha[14] !== "" && linha[14] !== null;

      if (!estaBloqueado && !temDesligamento) {
        lista.push({
          linhaPlanilha: i + 1,
          nome: linha[0] || "",
          email: linha[1] || "",
          cpf: linha[2] || "",
          sexo: linha[3] || "",
          celular: linha[4] || "",
          dataNascimento: formatarDataExibicao(linha[5]),
          dataNascimentoISO: obterDataISO(linha[5]),
          dataContratacao: formatarDataExibicao(linha[6]),
          dataContratacaoISO: obterDataISO(linha[6]),
          formacao: linha[7] || "",
          cargo: linha[8] || "",
          tags: linha[9] || "",
          unidadeGeografica: linha[10] || "",
          unidadeNegocio: linha[11] || "",
          grupos: linha[12] || "",
          lideres: linha[13] || "",
          dataDesligamentoISO: obterDataISO(linha[14]),
          tipoDesligamento: linha[15] || "",
          bloquear: linha[16] || "Não"
        });
      }
    }

    return { sucesso: true, dados: lista };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao carregar dados: " + e.message };
  }
}

/**
 * BUSCA OS COLABORADORES ATIVOS PARA MONTAGEM DO ORGANOGRAMA.
 * A hierarquia usa o campo N (Líderes / Supervisor) da aba Colaboradores.
 * Quando o líder informado corresponde a um único nome ativo, o backend
 * devolve o CPF desse superior para o frontend montar a árvore.
 */
function buscarOrganograma() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaColaboradores();
    var dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return { sucesso: true, dados: [] };

    var ativos = [];
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var bloqueado = String(linha[16] == null ? "" : linha[16]).trim().toLowerCase() === "sim";
      var desligado = linha[14] !== "" && linha[14] !== null;
      if (bloqueado || desligado) continue;
      var cpf = linha[2] || "";
      ativos.push({
        linhaPlanilha: i + 1,
        nome: linha[0] || "",
        cpf: cpf,
        cpfChave: normalizarCPF(cpf),
        cargo: linha[8] || "",
        unidadeGeografica: linha[10] || "",
        unidadeNegocio: linha[11] || "",
        lideres: linha[13] || "",
        dataContratacao: formatarDataExibicao(linha[6])
      });
    }

    var porNome = {};
    ativos.forEach(function(p) {
      var nome = normalizarTexto(p.nome);
      if (!nome) return;
      if (!porNome[nome]) porNome[nome] = [];
      porNome[nome].push(p);
    });

    ativos.forEach(function(p) {
      var partes = String(p.lideres || "")
        .split(/\s*(?:\/|,|;|\||\s+e\s+)\s*/i)
        .map(function(v){ return normalizarTexto(v); })
        .filter(Boolean);
      var candidatos = [];
      partes.forEach(function(nome) {
        var encontrados = porNome[nome] || [];
        if (encontrados.length === 1 && normalizarCPF(encontrados[0].cpf) !== normalizarCPF(p.cpf)) candidatos.push(encontrados[0]);
      });
      if (candidatos.length) {
        p.superiorCpf = normalizarCPF(candidatos[0].cpf);
        p.superiorNome = candidatos[0].nome;
      } else {
        p.superiorCpf = "";
        p.superiorNome = "";
      }
    });

    ativos.sort(function(a,b){ return String(a.nome || "").localeCompare(String(b.nome || ""), "pt-BR", {sensitivity:"base"}); });
    return { sucesso: true, dados: ativos, total: ativos.length };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao carregar organograma: " + e.message };
  }
}

/**
 * BUSCA DADOS DE EXPERIÊNCIA (REGRA: DIA 1 +89 DIAS PARA O 90º DIA)
 */
function garantirColunasStatusExperiencia(aba) {
  var cabecalhos = [
    {col:18, nome:"Status Avaliação 45"},
    {col:19, nome:"Status Avaliação 90"},
    {col:20, nome:"Status Avaliação 15"},
    {col:21, nome:"Status Avaliação 30"},
    {col:22, nome:"Status Avaliação 60"},
    {col:23, nome:"Status Avaliação 75"}
  ];
  cabecalhos.forEach(function(item) {
    var celula = aba.getRange(1, item.col);
    if (!String(celula.getValue() || "").trim()) celula.setValue(item.nome);
  });
}

/**
 * BUSCA DADOS DE EXPERIÊNCIA.
 * Integração/admissão é o DIA 1:
 * 15=+14, 30=+29, 45=+44, 60=+59, 75=+74, 90=+89.
 */
function buscarAvaliacaoExperiencia() {
  try {
    var aba = obterAbaColaboradores();
    garantirColunasStatusExperiencia(aba);
    var dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return { sucesso: true, dados: [] };

    var hoje = new Date();
    hoje.setHours(0,0,0,0);
    var limite90DiasAtras = new Date(hoje);
    limite90DiasAtras.setDate(hoje.getDate() - 90);

    var resultado = [];
    var periodos = [
      {tipo:"15", dias:14, indiceStatus:19},
      {tipo:"30", dias:29, indiceStatus:20},
      {tipo:"45", dias:44, indiceStatus:17},
      {tipo:"60", dias:59, indiceStatus:21},
      {tipo:"75", dias:74, indiceStatus:22},
      {tipo:"90", dias:89, indiceStatus:18}
    ];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var dtAdmissao = linha[6] instanceof Date ? linha[6] : new Date(linha[6]);
      if (isNaN(dtAdmissao.getTime()) || dtAdmissao < limite90DiasAtras || dtAdmissao > hoje) continue;

      var item = {
        linhaPlanilha: i + 1,
        nome: linha[0] || "",
        cargo: linha[8] || "",
        gerencia: linha[10] || "Sem Gerência",
        loja: linha[11] || "Sem Loja",
        supervisor: linha[13] || "Sem Supervisor",
        dataAdmissao: formatarDataExibicao(dtAdmissao),
        dataAdmissaoISO: obterDataISO(dtAdmissao)
      };

      periodos.forEach(function(p) {
        var venc = new Date(dtAdmissao);
        venc.setDate(venc.getDate() + p.dias);
        var salvo = String(linha[p.indiceStatus] || "").trim().toLowerCase();
        var concluida = ["feito","realizado","concluído","concluido","sim"].indexOf(salvo) >= 0;
        var statusSalvo = normalizarStatusAvaliacao(salvo);
        var status = statusSalvo || (concluida ? "Feito" : (venc < hoje ? "Vencido" : "A Vencer"));
        item["venc" + p.tipo] = formatarDataExibicao(venc);
        item["venc" + p.tipo + "ISO"] = obterDataISO(venc);
        item["status" + p.tipo] = status;
        item["avaliacao" + p.tipo + "Concluida"] = concluida;
      });

      resultado.push(item);
    }
    return { sucesso: true, dados: resultado };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao buscar avaliações: " + e.message };
  }
}

/**
 * ALTERNA O STATUS DA AVALIAÇÃO AO CLICAR NO LÁPIS
 */
function alternarStatusAvaliacao(linhaPlanilha, tipoEtiqueta, novoStatus) {
  try {
    var aba = obterAbaColaboradores();
    var linha = parseInt(linhaPlanilha, 10);
    if (!linha || linha < 2) throw new Error("Linha da planilha inválida.");

    var tipo = String(tipoEtiqueta);
    var mapaColunas = {"15":20,"30":21,"45":18,"60":22,"75":23,"90":19};
    if (!mapaColunas[tipo]) throw new Error("Tipo de avaliação inválido.");

    var status = String(novoStatus || "").trim();
    if (["A Vencer", "Vencido", "Feito"].indexOf(status) < 0) {
      throw new Error("Status inválido. Use A Vencer, Vencido ou Feito.");
    }

    aba.getRange(linha, mapaColunas[tipo]).setValue(status);
    return { sucesso: true, statusSalvo: status };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao atualizar status da avaliação: " + e.message };
  }
}

/**
 * IMPORTAÇÃO DE DADOS EM MASSA (EXCEL/CSV)
 */
function salvarEmMassa(listaDados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaColaboradores();

    var matriz = listaDados.map(function(item) {
      return [
        item["Nome"] || item["nome"] || "",
        item["Email"] || item["email"] || "",
        item["CPF"] || item["cpf"] || "",
        item["Sexo"] || item["sexo"] || "",
        item["Celular"] || item["celular"] || "",
        item["Data Nascimento"] ? formatarDataISO(item["Data Nascimento"]) : "",
        item["Data Contratacao"] ? formatarDataISO(item["Data Contratacao"]) : "",
        item["Formacao"] || item["formacao"] || "",
        item["Cargo"] || item["cargo"] || "",
        item["Tags"] || item["tags"] || "",
        item["Unidade Geografica"] || item["unidadeGeografica"] || "",
        item["Unidade Negocio"] || item["unidadeNegocio"] || "",
        item["Grupos"] || item["grupos"] || "",
        item["Lideres"] || item["lideres"] || "",
        item["Data Desligamento"] ? formatarDataISO(item["Data Desligamento"]) : "",
        item["Tipo Desligamento"] || item["tipoDesligamento"] || "",
        item["Bloquear"] || item["bloquear"] || "Não",
        "Pendente",
        "Pendente"
      ];
    });

    if (matriz.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, matriz.length, matriz[0].length).setValues(matriz);
      listaDados.forEach(function(item){ registrarMovimentacao({cpf:item["CPF"]||item["cpf"]||"",nome:item["Nome"]||item["nome"]||"",dataMovimentacao:item["Data Contratacao"]?formatarDataISO(item["Data Contratacao"]):new Date(),tipoMovimentacao:"Admissão",cargoAnterior:"",cargoNovo:item["Cargo"]||item["cargo"]||"",lojaAnterior:"",lojaNova:item["Unidade Negocio"]||item["unidadeNegocio"]||"",supervisorAnterior:"",supervisorNovo:item["Lideres"]||item["lideres"]||"",origem:"Carga de Admissões",observacao:"Registro inicial do colaborador"}); });
    }

    return { sucesso: true, mensagem: matriz.length + " registros importados com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao processar carga em massa: " + e.message };
  }
}

/**
 * NORMALIZA STATUS MANUAL DA AVALIAÇÃO
 * Retorna vazio para valores antigos que não representam um status explícito.
 */
function normalizarStatusAvaliacao(status) {
  var valor = String(status || "").trim().toLowerCase();

  if (valor === "a vencer") return "A Vencer";
  if (valor === "vencido") return "Vencido";
  if (valor === "feito") return "Feito";

  return "";
}


/**
 * GERA O MODELO CSV PARA NOVAS ADMISSÕES.
 * O arquivo contém somente os campos que devem ser preenchidos
 * no cadastro de um novo colaborador.
 */
function gerarModeloAdmissaoCSV() {
  var cabecalho = [
    "Nome",
    "Email",
    "CPF",
    "Sexo",
    "Celular",
    "Data Nascimento",
    "Data Contratacao",
    "Formacao",
    "Cargo",
    "Tags",
    "Unidade Geografica",
    "Unidade Negocio",
    "Grupos",
    "Lideres"
  ];

  return gerarCSV(cabecalho, []);
}

/**
 * GERA O MODELO CSV PARA DEMISSÕES.
 * Somente Nome e CPF são necessários no arquivo.
 */
function gerarModeloDemissaoCSV() {
  return gerarCSV([
    "Nome Completo",
    "CPF",
    "Data de Desligamento",
    "Motivo do Desligamento"
  ], []);
}

/**
 * GERA O MODELO EXCEL DE DEMISSÃO COM LISTA SUSPENSA NA COLUNA D.
 *
 * A planilha gerada possui:
 * - A: Nome Completo
 * - B: CPF
 * - C: Data de Desligamento
 * - D: Motivo do Desligamento
 *
 * A coluna D recebe uma validação de dados com os cinco motivos permitidos.
 * O arquivo é exportado em XLSX para preservar a lista suspensa.
 */
function gerarModeloDemissaoXLSX() {
  try {
    var motivosPermitidos = [
      "Sem justa causa",
      "Por justa causa",
      "Pedido de demissão",
      "Demissão consensual",
      "Rescisão indireta"
    ];

    // Gera o XLSX diretamente no Apps Script, sem chamadas externas.
    // Isso evita qualquer necessidade de autorização para requisições HTTP externas.
    var nomeArquivo = "modelo_demissao.xlsx";
    var cabecalhos = [
      "Nome Completo",
      "CPF",
      "Data de Desligamento",
      "Motivo do Desligamento"
    ];

    function xmlEscapar(valor) {
      return String(valor == null ? "" : valor)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
    }

    function colunaParaNumero(coluna) {
      var n = 0;
      for (var i = 0; i < coluna.length; i++) {
        n = n * 26 + coluna.charCodeAt(i) - 64;
      }
      return n;
    }

    function numeroParaColuna(n) {
      var resultado = "";
      while (n > 0) {
        var resto = (n - 1) % 26;
        resultado = String.fromCharCode(65 + resto) + resultado;
        n = Math.floor((n - 1) / 26);
      }
      return resultado;
    }

    // XML da linha de cabeçalho.
    var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="20"/>' +
      '<cols>' +
        '<col min="1" max="1" width="34" customWidth="1"/>' +
        '<col min="2" max="2" width="22" customWidth="1"/>' +
        '<col min="3" max="3" width="22" customWidth="1"/>' +
        '<col min="4" max="4" width="34" customWidth="1"/>' +
      '</cols>' +
      '<sheetData><row r="1" ht="24" customHeight="1">';

    for (var c = 0; c < cabecalhos.length; c++) {
      var ref = numeroParaColuna(c + 1) + "1";
      sheetXml += '<c r="' + ref + '" s="1" t="inlineStr"><is><t>' +
        xmlEscapar(cabecalhos[c]) + '</t></is></c>';
    }

    sheetXml += '</row></sheetData>' +
      '<autoFilter ref="A1:D1000"/>' +
      '<dataValidations count="1">' +
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorStyle="stop" ' +
          'errorTitle="Motivo inválido" error="Selecione um motivo da lista." ' +
          'promptTitle="Motivo do Desligamento" prompt="Selecione um dos motivos disponíveis." sqref="D2:D1000">' +
          '<formula1>"' + xmlEscapar(motivosPermitidos.join(",")) + '"</formula1>' +
        '</dataValidation>' +
      '</dataValidations>' +
      '<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>' +
      '</worksheet>';

    var workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="1"/>' +
      '<workbookPr defaultThemeVersion="124226"/>' +
      '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16000" windowHeight="9000"/></bookViews>' +
      '<sheets><sheet name="Demissões" sheetId="1" r:id="rId1"/></sheets>' +
      '</workbook>';

    var workbookRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>';

    var rootRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';

    var contentTypesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>';

    var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>' +
      '<fonts count="2"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>';

    var blobs = [
      Utilities.newBlob(contentTypesXml, "application/xml", "[Content_Types].xml"),
      Utilities.newBlob(rootRelsXml, "application/vnd.openxmlformats-package.relationships+xml", "_rels/.rels"),
      Utilities.newBlob(workbookXml, "application/xml", "xl/workbook.xml"),
      Utilities.newBlob(workbookRelsXml, "application/xml", "xl/_rels/workbook.xml.rels"),
      Utilities.newBlob(stylesXml, "application/xml", "xl/styles.xml"),
      Utilities.newBlob(sheetXml, "application/xml", "xl/worksheets/sheet1.xml")
    ];

    var xlsxBlob = Utilities.zip(blobs, nomeArquivo);
    return Utilities.base64Encode(xlsxBlob.getBytes());

  } catch (e) {
    throw new Error("Erro ao gerar o modelo Excel de demissão: " + e.message);
  }
}

/**
 * ATUALIZA DEMISSÕES EM MASSA.
 *
 * Para cada CPF encontrado na base:
 * - Data de Desligamento (coluna O) = data do processamento;
 * - Bloquear (coluna Q) = Sim.
 *
 * O nome é usado como conferência quando houver mais de um registro
 * com o mesmo CPF.
 */
/**
 * MODELO EXCEL DE MOVIMENTAÇÕES COM LISTA SUSPENSA.
 */
function gerarModeloMovimentacoesXLSX() {
  try {
    var tipos=["Promoção","Transferência de Loja","Transferência de Unidade","Alteração de Cargo","Alteração de Supervisor","Promoção + Transferência","Outro"];
    var cab=["Nome Completo","CPF","Loja / Unidade Nova","Cargo Novo","Supervisor Novo","Data de Movimentação","Tipo de Movimentação","Observação"];
    function xe(v){return String(v==null?"":v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;");}
    function cn(n){var r="";while(n){var x=(n-1)%26;r=String.fromCharCode(65+x)+r;n=Math.floor((n-1)/26);}return r;}
    function c(ref,v,st){return '<c r="'+ref+'" t="inlineStr"'+(st?' s="'+st+'"':'')+'><is><t>'+xe(v)+'</t></is></c>';}
    var sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/><cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="18" customWidth="1"/><col min="3" max="5" width="24" customWidth="1"/><col min="6" max="6" width="20" customWidth="1"/><col min="7" max="7" width="28" customWidth="1"/><col min="8" max="8" width="35" customWidth="1"/></cols><sheetData><row r="1">'+cab.map(function(h,i){return c(cn(i+1)+"1",h,1);}).join('')+'</row></sheetData><autoFilter ref="A1:H1000"/><dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="G2:G1000"><formula1>"'+tipos.join(',')+'"</formula1></dataValidation></dataValidations><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
    var styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs></styleSheet>';
    var workbook='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Movimentações" sheetId="1" r:id="rId1"/></sheets></workbook>';
    var rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    var root='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
    var zip=Utilities.zip([Utilities.newBlob(ct,'application/xml','[Content_Types].xml'),Utilities.newBlob(root,'application/xml','_rels/.rels'),Utilities.newBlob(workbook,'application/xml','xl/workbook.xml'),Utilities.newBlob(rels,'application/xml','xl/_rels/workbook.xml.rels'),Utilities.newBlob(sheet,'application/xml','xl/worksheets/sheet1.xml'),Utilities.newBlob(styles,'application/xml','xl/styles.xml')],"modelo_movimentacoes.xlsx");
    return {sucesso:true,nomeArquivo:"modelo_movimentacoes.xlsx",mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",base64:Utilities.base64Encode(zip.getBytes())};
  } catch(e){return {sucesso:false,mensagem:"Erro ao gerar modelo de movimentações: "+e.message};}
}

/** PROCESSA MOVIMENTAÇÕES EM MASSA. */
function processarMovimentacoesEmMassa(listaDados) {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet(), aba=obterAbaColaboradores(), dados=aba.getDataRange().getValues();
    var tipos=["Promoção","Transferência de Loja","Transferência de Unidade","Alteração de Cargo","Alteração de Supervisor","Promoção + Transferência","Outro"], ok=0, ign=0, erros=[];
    (listaDados||[]).forEach(function(item,idx){try{
      var cpf=normalizarCPF(item["CPF"]||item["cpf"]||""), nome=normalizarTexto(item["Nome Completo"]||item["Nome"]||item["nome"]||"");
      if(!cpf)throw new Error("CPF não informado."); if(!nome)throw new Error("Nome Completo não informado.");
      var cs=[]; for(var i=1;i<dados.length;i++)if(normalizarCPF(dados[i][2])===cpf)cs.push({linha:i+1,row:dados[i]});
      if(!cs.length)throw new Error("CPF não encontrado na base.");
      if(cs.length>1){var por=cs.filter(function(x){return normalizarTexto(x.row[0])===nome;});if(por.length!==1)throw new Error("CPF duplicado e o nome não foi suficiente para identificar o colaborador.");cs=por;}
      var c=cs[0], row=c.row, loja=item["Loja / Unidade Nova"]||item["Loja Nova"]||item["Unidade Nova"]||"", cargo=item["Cargo Novo"]||"", sup=item["Supervisor Novo"]||"", dt=item["Data de Movimentação"]||item["Data Movimentacao"]||"", tipo=String(item["Tipo de Movimentação"]||item["Tipo Movimentacao"]||"").trim(), obs=item["Observação"]||item["Observacao"]||"";
      if(!dt)throw new Error("Data de Movimentação não informada."); if(tipos.indexOf(tipo)<0)throw new Error("Tipo de Movimentação inválido."); var data=converterDataFlexivel(dt); if(!data)throw new Error("Data de Movimentação inválida.");
      var ca=row[8]||"", la=row[11]||"", sa=row[13]||"", cf=cargo||ca, lf=loja||la, sf=sup||sa;
      if(normalizarTexto(ca)===normalizarTexto(cf)&&normalizarTexto(la)===normalizarTexto(lf)&&normalizarTexto(sa)===normalizarTexto(sf)){ign++;return;}
      aba.getRange(c.linha,9).setValue(cf); aba.getRange(c.linha,12).setValue(lf); aba.getRange(c.linha,14).setValue(sf);
      registrarMovimentacao({cpf:cpf,nome:row[0]||nome,dataMovimentacao:data,tipoMovimentacao:tipo,cargoAnterior:ca,cargoNovo:cf,lojaAnterior:la,lojaNova:lf,supervisorAnterior:sa,supervisorNovo:sf,origem:"Carga de Movimentações",observacao:obs}); ok++;
    }catch(e){erros.push("Linha "+(idx+2)+": "+e.message);}});
    return {sucesso:erros.length===0,mensagem:ok+" movimentação(ões) processada(s)."+(ign?" "+ign+" sem alteração foram ignoradas.":"")+(erros.length?" Erros: "+erros.join(" | "):"")};
  }catch(e){return {sucesso:false,mensagem:"Erro ao processar movimentações: "+e.message};}
}

function garantirAbaHistoricoMovimentacoes(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), aba=ss.getSheetByName("Histórico Funcionários");
  if(!aba){aba=ss.insertSheet("Histórico Funcionários");aba.getRange(1,1,1,14).setValues([["ID Movimento","CPF","Nome","Data Movimentação","Data Registro","Tipo de Movimentação","Cargo Anterior","Cargo Novo","Loja / Unidade Anterior","Loja / Unidade Nova","Supervisor Anterior","Supervisor Novo","Origem","Observação"]]);aba.setFrozenRows(1);aba.getRange(1,1,1,14).setFontWeight("bold");}
  return aba;
}

function registrarMovimentacao(dados){
  var aba=garantirAbaHistoricoMovimentacoes(), id="MOV-"+Utilities.getUuid().substring(0,8).toUpperCase(), dt=dados.dataMovimentacao?formatarDataISO(dados.dataMovimentacao):new Date();
  aba.appendRow([id,dados.cpf||"",dados.nome||"",dt,new Date(),dados.tipoMovimentacao||"Outro",dados.cargoAnterior||"",dados.cargoNovo||"",dados.lojaAnterior||"",dados.lojaNova||"",dados.supervisorAnterior||"",dados.supervisorNovo||"",dados.origem||"",dados.observacao||""]);
  return id;
}

function buscarHistoricoFuncionario(cpf){
  try{var chave=normalizarCPF(cpf||"");if(!chave)throw new Error("CPF não informado.");var aba=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Histórico Funcionários");if(!aba)return {sucesso:true,dados:[],resumo:null};var v=aba.getDataRange().getValues(),lista=[];for(var i=1;i<v.length;i++)if(normalizarCPF(v[i][1])===chave)lista.push({idMovimento:v[i][0]||"",cpf:v[i][1]||"",nome:v[i][2]||"",dataMovimentacao:formatarDataExibicao(converterDataFlexivel(v[i][3])),dataMovimentacaoISO:obterDataISO(converterDataFlexivel(v[i][3])),dataRegistro:formatarDataExibicao(converterDataFlexivel(v[i][4])),tipoMovimentacao:v[i][5]||"",cargoAnterior:v[i][6]||"",cargoNovo:v[i][7]||"",lojaAnterior:v[i][8]||"",lojaNova:v[i][9]||"",supervisorAnterior:v[i][10]||"",supervisorNovo:v[i][11]||"",origem:v[i][12]||"",observacao:v[i][13]||""});lista.sort(function(a,b){return String(b.dataMovimentacaoISO).localeCompare(String(a.dataMovimentacaoISO));});return {sucesso:true,dados:lista};}catch(e){return {sucesso:false,mensagem:"Erro ao buscar histórico: "+e.message};}
}

function processarDemissoesEmMassa(listaDados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaColaboradores();
    if (!listaDados || !listaDados.length) throw new Error("O arquivo não contém registros.");

    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) throw new Error("Não existem colaboradores cadastrados na base.");

    // Para demissões, não é necessário carregar as 19 colunas da base.
    // Lemos apenas Nome + CPF para localizar os colaboradores e depois
    // alteramos somente O:Q das linhas encontradas.
    var baseAC = aba.getRange(2, 1, ultimaLinha - 1, 3).getValues();
    var porCpf = {};
    for (var i = 0; i < baseAC.length; i++) {
      var cpfBase = normalizarCPF(baseAC[i][2]);
      if (!cpfBase) continue;
      if (!porCpf[cpfBase]) porCpf[cpfBase] = [];
      porCpf[cpfBase].push(i + 2); // número real da linha na planilha
    }

    var motivosPermitidos = MOTIVOS_DESLIGAMENTO_PERMITIDOS.slice();
    var atualizados = 0;
    var naoEncontrados = [];
    var invalidos = [];
    var alteracoes = {};

    listaDados.forEach(function(item, idx) {
      try {
        var nomeArquivo = String(
          item["Nome Completo"] || item["Nome"] || item["nomeCompleto"] || item["nome"] || ""
        ).trim();
        var cpfArquivo = normalizarCPF(item["CPF"] || item["cpf"] || "");
        var dataArquivo = item["Data de Desligamento"] || item["Data Desligamento"] || item["dataDesligamento"] || item["Data de desligamento"] || "";
        var motivoArquivo = String(
          item["Motivo do Desligamento"] || item["Motivo Desligamento"] || item["motivoDoDesligamento"] || item["motivoDesligamento"] || item["Tipo Desligamento"] || item["tipoDesligamento"] || ""
        ).trim();

        if (!cpfArquivo) {
          invalidos.push((nomeArquivo || "(sem nome)") + " - CPF não informado");
          return;
        }
        if (!dataArquivo) {
          invalidos.push((nomeArquivo || "(sem nome)") + " - Data de Desligamento não informada");
          return;
        }

        var dataDesligamento;
        try {
          dataDesligamento = obterDataDesligamentoMassa(dataArquivo);
        } catch (erroData) {
          invalidos.push((nomeArquivo || "(sem nome)") + " - " + erroData.message);
          return;
        }

        var motivoCanonico = normalizarMotivoDesligamento(motivoArquivo);
        if (motivosPermitidos.indexOf(motivoCanonico) < 0) {
          invalidos.push((nomeArquivo || "(sem nome)") + " - Motivo inválido. Use: " + motivosPermitidos.join(", "));
          return;
        }

        var candidatos = porCpf[cpfArquivo] || [];
        if (!candidatos.length) {
          naoEncontrados.push((nomeArquivo || "(sem nome)") + " - CPF " + cpfArquivo + " não encontrado");
          return;
        }

        var linhaEscolhida = candidatos[0];
        if (candidatos.length > 1) {
          if (!nomeArquivo) {
            naoEncontrados.push("CPF " + cpfArquivo + " duplicado na base; informe o Nome Completo.");
            return;
          }
          var nomeNormalizado = normalizarTexto(nomeArquivo);
          var encontradosPorNome = candidatos.filter(function(linha) {
            return normalizarTexto(baseAC[linha - 2][0]) === nomeNormalizado;
          });
          if (encontradosPorNome.length !== 1) {
            naoEncontrados.push((nomeArquivo || "(sem nome)") + " - CPF duplicado na base e nome não permitiu identificar uma única pessoa");
            return;
          }
          linhaEscolhida = encontradosPorNome[0];
        }

        alteracoes[linhaEscolhida] = {
          data: dataDesligamento,
          motivo: motivoCanonico
        };
      } catch (e) {
        invalidos.push("Linha " + (idx + 2) + ": " + e.message);
      }
    });

    var linhas = Object.keys(alteracoes).map(Number).sort(function(a, b) { return a - b; });
    if (!linhas.length) {
      return {
        sucesso: false,
        mensagem: "Nenhuma demissão foi aplicada. " +
          (invalidos.length ? invalidos.length + " registro(s) inválido(s). " : "") +
          (naoEncontrados.length ? naoEncontrados.length + " registro(s) não localizado(s)." : ""),
        atualizados: 0,
        naoEncontrados: naoEncontrados,
        invalidos: invalidos
      };
    }

    // Lê somente O:Q das linhas que realmente serão alteradas.
    // As alterações são agrupadas por linhas consecutivas para reduzir
    // drasticamente a quantidade de operações no Spreadsheet Service.
    var blocos = [];
    var blocoInicio = linhas[0];
    var blocoFim = linhas[0];
    for (var j = 1; j < linhas.length; j++) {
      if (linhas[j] === blocoFim + 1) {
        blocoFim = linhas[j];
      } else {
        blocos.push([blocoInicio, blocoFim]);
        blocoInicio = blocoFim = linhas[j];
      }
    }
    blocos.push([blocoInicio, blocoFim]);

    blocos.forEach(function(bloco) {
      var inicio = bloco[0];
      var quantidade = bloco[1] - bloco[0] + 1;
      var atual = aba.getRange(inicio, 15, quantidade, 3).getValues();
      for (var k = 0; k < quantidade; k++) {
        var linhaReal = inicio + k;
        if (!alteracoes[linhaReal]) continue;
        atual[k][0] = alteracoes[linhaReal].data;   // O - Data de Desligamento
        atual[k][1] = alteracoes[linhaReal].motivo; // P - Motivo do Desligamento
        atual[k][2] = "Sim";                        // Q - Bloquear
        atualizados++;
      }
      aba.getRange(inicio, 15, quantidade, 3).setValues(atual);
    });

    SpreadsheetApp.flush();

    var mensagem = atualizados + " demissão(ões) processada(s) com sucesso.";
    if (naoEncontrados.length) {
      mensagem += " " + naoEncontrados.length + " registro(s) não localizado(s): " + naoEncontrados.slice(0, 10).join(" | ");
      if (naoEncontrados.length > 10) mensagem += " | ...";
    }
    if (invalidos.length) {
      mensagem += " " + invalidos.length + " registro(s) inválido(s): " + invalidos.slice(0, 10).join(" | ");
      if (invalidos.length > 10) mensagem += " | ...";
    }

    return {
      sucesso: atualizados > 0 && invalidos.length === 0 && naoEncontrados.length === 0,
      mensagem: mensagem,
      atualizados: atualizados,
      naoEncontrados: naoEncontrados,
      invalidos: invalidos
    };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao processar demissões: " + e.message };
  }
}

/**
 * Converte a data recebida pelo XLSX/CSV para um objeto Date seguro.
 * Aceita:
 * - objeto Date
 * - yyyy-mm-dd
 * - dd/mm/yyyy
 * - dd-mm-yyyy
 * - número serial de Excel/Sheets
 */
function obterDataDesligamentoMassa(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    var data = new Date(valor);
    data.setHours(0, 0, 0, 0);
    return data;
  }

  if (typeof valor === "number" && isFinite(valor)) {
    // Excel serial date: 1899-12-30 é a base usada pelo Sheets/Excel.
    var serialDate = new Date(Date.UTC(1899, 11, 30) + Math.round(valor * 86400000));
    if (!isNaN(serialDate.getTime())) {
      return new Date(
        serialDate.getUTCFullYear(),
        serialDate.getUTCMonth(),
        serialDate.getUTCDate()
      );
    }
  }

  var texto = String(valor == null ? "" : valor).trim();

  if (!texto) {
    throw new Error("Data de Desligamento não informada");
  }

  // Remove horário quando vier junto da data.
  texto = texto.split(/[ T]/)[0];

  var matchISO = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (matchISO) {
    var anoISO = parseInt(matchISO[1], 10);
    var mesISO = parseInt(matchISO[2], 10) - 1;
    var diaISO = parseInt(matchISO[3], 10);
    var dataISO = new Date(anoISO, mesISO, diaISO);
    if (
      dataISO.getFullYear() === anoISO &&
      dataISO.getMonth() === mesISO &&
      dataISO.getDate() === diaISO
    ) {
      dataISO.setHours(0, 0, 0, 0);
      return dataISO;
    }
  }

  var matchBR = texto.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (matchBR) {
    var diaBR = parseInt(matchBR[1], 10);
    var mesBR = parseInt(matchBR[2], 10) - 1;
    var anoBR = parseInt(matchBR[3], 10);
    var dataBR = new Date(anoBR, mesBR, diaBR);
    if (
      dataBR.getFullYear() === anoBR &&
      dataBR.getMonth() === mesBR &&
      dataBR.getDate() === diaBR
    ) {
      dataBR.setHours(0, 0, 0, 0);
      return dataBR;
    }
  }

  var dataGenerica = new Date(texto);
  if (!isNaN(dataGenerica.getTime())) {
    dataGenerica.setHours(0, 0, 0, 0);
    return dataGenerica;
  }

  throw new Error("Data de Desligamento inválida: " + texto);
}

function normalizarCPF(valor) {
  return String(valor == null ? "" : valor).replace(/\D/g, "");
}

function normalizarTexto(valor) {
  return String(valor == null ? "" : valor)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function gerarCSV(cabecalho, linhas) {
  var todasLinhas = [cabecalho].concat(linhas || []);

  return todasLinhas.map(function(linha) {
    return linha.map(function(valor) {
      var texto = String(valor == null ? "" : valor);
      texto = texto.replace(/"/g, '""');
      return '"' + texto + '"';
    }).join(";");
  }).join("\r\n");
}


/**
 * BUSCA ANIVERSARIANTES DO DIA E ANIVERSARIANTES DE EMPRESA.
 * Considera somente colaboradores ativos:
 * - não bloqueados
 * - sem Data de Desligamento
 */


/** MODELO EXCEL PARA CARGA/RECONSTRUÇÃO DE HISTÓRICO. */
function gerarModeloHistoricoXLSX() {
  try {
    var cab=["Mês de Referência","Nome Completo","CPF","Data de Contratação","Cargo","Unidade Geográfica","Unidade de Negócio / Loja","Líderes / Supervisor","Data de Desligamento","Tipo de Desligamento","Email","Celular"];
    function xe(v){return String(v==null?"":v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\"/g,"&quot;").replace(/'/g,"&apos;");}
    function cn(n){var r="";while(n){var x=(n-1)%26;r=String.fromCharCode(65+x)+r;n=Math.floor((n-1)/26);}return r;}
    function c(ref,v,st){return '<c r="'+ref+'" t="inlineStr"'+(st?' s="'+st+'"':'')+'><is><t>'+xe(v)+'</t></is></c>';}
    var widths=[18,30,18,20,28,26,28,28,22,26,32,18];
    var cols=widths.map(function(w,i){return '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>';}).join('');
    var sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/><cols>'+cols+'</cols><sheetData><row r="1">'+cab.map(function(h,i){return c(cn(i+1)+"1",h,1);}).join('')+'</row></sheetData><autoFilter ref="A1:L10000"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
    var styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs></styleSheet>';
    var workbook='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Histórico" sheetId="1" r:id="rId1"/></sheets></workbook>';
    var rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    var root='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
    var zip=Utilities.zip([Utilities.newBlob(ct,'application/xml','[Content_Types].xml'),Utilities.newBlob(root,'application/xml','_rels/.rels'),Utilities.newBlob(workbook,'application/xml','xl/workbook.xml'),Utilities.newBlob(rels,'application/xml','xl/_rels/workbook.xml.rels'),Utilities.newBlob(sheet,'application/xml','xl/worksheets/sheet1.xml'),Utilities.newBlob(styles,'application/xml','xl/styles.xml')],"modelo_historico.xlsx");
    return {sucesso:true,nomeArquivo:"modelo_historico.xlsx",mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",base64:Utilities.base64Encode(zip.getBytes())};
  } catch(e){return {sucesso:false,mensagem:"Erro ao gerar modelo de histórico: "+e.message};}
}

function obterMesReferenciaHistorico(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) return new Date(valor.getFullYear(),valor.getMonth(),1);
  var s=String(valor==null?"":valor).trim();
  if(!s) return null;
  var m=s.match(/^(0?[1-9]|1[0-2])\s*[\\/-]\s*(\\d{4})$/);
  if(m) return new Date(Number(m[2]),Number(m[1])-1,1);
  m=s.match(/^(\\d{4})\s*[\\/-]\s*(0?[1-9]|1[0-2])$/);
  if(m) return new Date(Number(m[1]),Number(m[2])-1,1);
  var d=converterDataFlexivel(s); if(d) return new Date(d.getFullYear(),d.getMonth(),1);
  return null;
}

function normalizarItemHistorico(item) {
  var mes=obterMesReferenciaHistorico(item["Mês de Referência"]||item["Mes de Referencia"]||item["Mês"]||item["Mes"]||item["Referencia"]||"");
  var cpf=normalizarCPF(item["CPF"]||item["cpf"]||"");
  var nome=String(item["Nome Completo"]||item["Nome"]||item["nome"]||"").trim();
  return {
    mes:mes,
    mesISO:mes?Utilities.formatDate(mes,Session.getScriptTimeZone()||"America/Sao_Paulo","yyyy-MM"):"",
    nome:nome,
    cpf:cpf,
    dataContratacao:item["Data de Contratação"]||item["Data Contratacao"]||item["Data de Contratação"]||"",
    cargo:String(item["Cargo"]||item["Cargo Novo"]||"").trim(),
    unidadeGeografica:String(item["Unidade Geográfica"]||item["Unidade Geografica"]||item["Gerência"]||item["Gerencia"]||"").trim(),
    loja:String(item["Unidade de Negócio / Loja"]||item["Unidade de Negocio / Loja"]||item["Unidade de Negócio"]||item["Unidade Negocio"]||item["Loja / Unidade"]||item["Loja"]||"").trim(),
    supervisor:String(item["Líderes / Supervisor"]||item["Lideres / Supervisor"]||item["Líderes"]||item["Lideres"]||item["Supervisor"]||"").trim(),
    dataDesligamento:item["Data de Desligamento"]||item["Data Desligamento"]||"",
    tipoDesligamento:String(item["Tipo de Desligamento"]||item["Tipo Desligamento"]||"").trim(),
    email:String(item["Email"]||item["E-mail"]||item["email"]||"").trim(),
    celular:String(item["Celular"]||item["Telefone"]||item["celular"]||"").trim()
  };
}

function classificarMovimentacaoHistorica(anterior, atual) {
  var cargo=normalizarTexto(anterior.cargo)!==normalizarTexto(atual.cargo);
  var loja=normalizarTexto(anterior.loja)!==normalizarTexto(atual.loja);
  var unidade=normalizarTexto(anterior.unidadeGeografica)!==normalizarTexto(atual.unidadeGeografica);
  var sup=normalizarTexto(anterior.supervisor)!==normalizarTexto(atual.supervisor);
  var campos=[];
  if(cargo) campos.push("cargo"); if(loja) campos.push("loja"); if(unidade) campos.push("unidade"); if(sup) campos.push("supervisor");
  var tipo="";
  if(cargo && loja && !unidade && !sup) tipo="Promoção + Transferência";
  else if(loja && !cargo && !unidade && !sup) tipo="Transferência de Loja";
  else if(unidade && !cargo && !loja && !sup) tipo="Transferência de Unidade";
  else if(cargo && !loja && !unidade && !sup) tipo="Alteração de Cargo";
  else if(sup && !cargo && !loja && !unidade) tipo="Alteração de Supervisor";
  else if(cargo && loja && (unidade || sup)) tipo="Outro";
  else if(campos.length) tipo="Outro";
  return {tipo:tipo,cargo:cargo,loja:loja,unidade:unidade,supervisor:sup};
}

/** Analisa uma ou mais bases históricas sem alterar a base atual. */
function analisarHistoricoEmMassa(listaDados) {
  try {
    var itens=(listaDados||[]).map(normalizarItemHistorico);
    var erros=[], grupos={};
    itens.forEach(function(it,idx){
      if(!it.mes) { erros.push("Linha "+(idx+2)+": Mês de Referência inválido ou não informado."); return; }
      if(!it.cpf) { erros.push("Linha "+(idx+2)+": CPF não informado."); return; }
      if(!it.nome) { erros.push("Linha "+(idx+2)+": Nome Completo não informado."); return; }
      if(!grupos[it.mesISO]) grupos[it.mesISO]=[];
      grupos[it.mesISO].push(it);
    });
    if(erros.length) return {sucesso:false,mensagem:erros.slice(0,20).join(" | ")+(erros.length>20?" | ...":""),erros:erros};
    var meses=Object.keys(grupos).sort(), bases=[], movimentos=[], desaparecidos=[], vistosInicial=0;
    meses.forEach(function(mesISO,idx){
      var atual=grupos[mesISO], mapaAtual={};
      atual.forEach(function(it){mapaAtual[it.cpf]=it;});
      bases.push({mes:mesISO,total:atual.length});
      if(idx===0){
        vistosInicial += atual.length;
        return;
      }
      var anterior=grupos[meses[idx-1]], mapaAnterior={};
      anterior.forEach(function(it){mapaAnterior[it.cpf]=it;});
      Object.keys(mapaAnterior).forEach(function(cpf){
        if(!mapaAtual[cpf]) desaparecidos.push({cpf:cpf,nome:mapaAnterior[cpf].nome,ultimaBase:meses[idx-1],baseSeguinte:mesISO,anterior:mapaAnterior[cpf],dataDesligamento:"",tipoDesligamento:"Não informado",selecionado:true});
      });
      Object.keys(mapaAtual).forEach(function(cpf){
        if(!mapaAnterior[cpf]) return;
        var a=mapaAnterior[cpf], b=mapaAtual[cpf], cl=classificarMovimentacaoHistorica(a,b);
        if(cl.tipo){
          var obs=[];
          if(cl.unidade) obs.push("Unidade Geográfica: "+(a.unidadeGeografica||"Não informado")+" → "+(b.unidadeGeografica||"Não informado"));
          movimentos.push({cpf:cpf,nome:b.nome||a.nome,dataMovimentacaoISO:mesISO+"-01",tipoMovimentacao:cl.tipo,cargoAnterior:a.cargo,cargoNovo:b.cargo,lojaAnterior:a.loja,lojaNova:b.loja,supervisorAnterior:a.supervisor,supervisorNovo:b.supervisor,origem:"Reconstrução Histórica",observacao:obs.join(" | ")});
        }
      });
    });
    return {sucesso:true,bases:bases,desaparecidos:desaparecidos,movimentos:movimentos,totais:{bases:meses.length,registros:itens.length,movimentos:movimentos.length,desaparecidos:desaparecidos.length,primeiraBaseRegistros:vistosInicial},dadosOriginais:itens};
  } catch(e){return {sucesso:false,mensagem:"Erro ao analisar histórico: "+e.message};}
}

/** Confirma a reconstrução histórica e grava os eventos/possíveis desligamentos. */
function confirmarHistoricoEmMassa(payload) {
  try {
    payload=payload||{};
    var ss=SpreadsheetApp.getActiveSpreadsheet(), aba=obterAbaColaboradores();
    var historico=garantirAbaHistoricoMovimentacoes(), movimentos=payload.movimentos||[], desaparecidos=payload.desaparecidos||[], okMov=0, okDes=0, ignorados=0, erros=[];
    function jaExiste(cpf,tipo,dataISO){
      var vals=historico.getDataRange().getValues(), chave=normalizarCPF(cpf), alvo=String(dataISO||"").slice(0,10);
      for(var i=1;i<vals.length;i++) if(normalizarCPF(vals[i][1])===chave && String(vals[i][5]||"")===String(tipo||"") && obterDataISO(converterDataFlexivel(vals[i][3]))===alvo) return true;
      return false;
    }
    movimentos.forEach(function(m){
      try{
        if(!m.cpf||!m.tipoMovimentacao) return;
        if(jaExiste(m.cpf,m.tipoMovimentacao,m.dataMovimentacaoISO)){ignorados++;return;}
        var dt=converterDataFlexivel(m.dataMovimentacaoISO)||new Date();
        registrarMovimentacao({cpf:m.cpf,nome:m.nome,dataMovimentacao:dt,tipoMovimentacao:m.tipoMovimentacao,cargoAnterior:m.cargoAnterior,cargoNovo:m.cargoNovo,lojaAnterior:m.lojaAnterior,lojaNova:m.lojaNova,supervisorAnterior:m.supervisorAnterior,supervisorNovo:m.supervisorNovo,origem:"Reconstrução Histórica",observacao:m.observacao||""}); okMov++;
      }catch(e){erros.push("Movimentação "+m.cpf+": "+e.message);}
    });
    desaparecidos.forEach(function(d){
      try{
        if(!d.selecionado) return;
        var cpf=normalizarCPF(d.cpf); if(!cpf) return;
        var vals=aba.getDataRange().getValues(), matches=[];
        for(var i=1;i<vals.length;i++) if(normalizarCPF(vals[i][2])===cpf) matches.push({linha:i+1,row:vals[i]});
        if(!matches.length) { erros.push("Desligamento "+d.nome+": CPF não encontrado na base atual."); return; }
        var alvo=matches[0];
        if(matches.length>1){var nomeN=normalizarTexto(d.nome);var ex=matches.filter(function(x){return normalizarTexto(x.row[0])===nomeN;});if(ex.length!==1){erros.push("Desligamento "+d.nome+": CPF duplicado e nome não foi suficiente.");return;}alvo=ex[0];}
        var dt="", tipo=String(d.tipoDesligamento||"Não informado").trim()||"Não informado";
        if(d.dataDesligamento){var parsed=converterDataFlexivel(d.dataDesligamento);dt=parsed?formatarDataISO(parsed):"Não informado";} else dt="Não informado";
        aba.getRange(alvo.linha,15).setValue(dt); aba.getRange(alvo.linha,16).setValue(tipo); aba.getRange(alvo.linha,17).setValue("Sim");
        var ref=d.baseSeguinte||d.ultimaBase||"";
        var obs="Desligamento identificado por ausência na base seguinte ("+ref+"). Data de desligamento: "+dt+".";
        if(!jaExiste(cpf,"Desligamento",dt!=="Não informado"?dt:"")) registrarMovimentacao({cpf:cpf,nome:alvo.row[0]||d.nome,dataMovimentacao:dt!=="Não informado"?(converterDataFlexivel(dt)||new Date()):new Date(),tipoMovimentacao:"Desligamento",cargoAnterior:alvo.row[8]||"",cargoNovo:"",lojaAnterior:alvo.row[11]||"",lojaNova:"",supervisorAnterior:alvo.row[13]||"",supervisorNovo:"",origem:"Reconstrução Histórica",observacao:obs});
        okDes++;
      }catch(e){erros.push("Desligamento "+d.nome+": "+e.message);}
    });
    return {sucesso:erros.length===0,mensagem:okMov+" movimentação(ões) históricas registrada(s) e "+okDes+" desligamento(s) confirmado(s)."+(ignorados?" "+ignorados+" evento(s) já existente(s) foram ignorados.":"")+(erros.length?" Erros: "+erros.join(" | "):"")};
  }catch(e){return {sucesso:false,mensagem:"Erro ao confirmar reconstrução histórica: "+e.message};}
}

/**
 * CONSULTA DESLIGAMENTOS POR MÊS/ANO.
 * Usa a Data de Desligamento da coluna O como referência e retorna os dados
 * atuais do colaborador junto com o motivo registrado na coluna P.
 */
function buscarDemitidosDoMes(mes, ano) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaColaboradores();
    var dados = aba.getDataRange().getValues();

    var hoje = new Date();
    var mesNumero = Number(mes);
    var anoNumero = Number(ano);
    if (!mesNumero || mesNumero < 1 || mesNumero > 12) mesNumero = hoje.getMonth() + 1;
    if (!anoNumero || anoNumero < 1900 || anoNumero > 2500) anoNumero = hoje.getFullYear();

    var lista = [];
    if (dados.length <= 1) {
      return { sucesso: true, mes: mesNumero, ano: anoNumero, dados: [], total: 0 };
    }

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var dtDesligamento = converterDataFlexivel(linha[14]);
      if (!dtDesligamento) continue;

      if (dtDesligamento.getMonth() + 1 !== mesNumero || dtDesligamento.getFullYear() !== anoNumero) continue;

      lista.push({
        linhaPlanilha: i + 1,
        nome: linha[0] || "",
        email: linha[1] || "",
        cpf: linha[2] || "",
        sexo: linha[3] || "",
        celular: linha[4] || "",
        dataNascimento: formatarDataExibicao(converterDataFlexivel(linha[5])),
        dataContratacao: formatarDataExibicao(converterDataFlexivel(linha[6])),
        cargo: linha[8] || "",
        loja: linha[11] || "Sem Loja",
        supervisor: linha[13] || "Sem Supervisor",
        dataDesligamento: formatarDataExibicao(dtDesligamento),
        dataDesligamentoISO: obterDataISO(dtDesligamento),
        motivoDesligamento: linha[15] || "Não informado",
        bloquear: linha[16] || ""
      });
    }

    lista.sort(function(a, b) {
      if (a.dataDesligamentoISO !== b.dataDesligamentoISO) {
        return a.dataDesligamentoISO < b.dataDesligamentoISO ? -1 : 1;
      }
      return normalizarTexto(a.nome).localeCompare(normalizarTexto(b.nome), "pt-BR");
    });

    return {
      sucesso: true,
      mes: mesNumero,
      ano: anoNumero,
      dados: lista,
      total: lista.length
    };

  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao buscar demitidos do mês: " + e.message };
  }
}

function buscarAniversariantesHoje() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = obterAbaColaboradores();
    var dados = aba.getDataRange().getValues();

    var hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    var aniversarios = [];
    var aniversariosEmpresa = [];

    if (dados.length <= 1) {
      return {
        sucesso: true,
        dataReferencia: formatarDataExibicao(hoje),
        aniversarios: [],
        aniversariosEmpresa: []
      };
    }

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];

      var estaBloqueado = String(linha[16] || "").trim().toLowerCase() === "sim";
      var temDesligamento = linha[14] !== "" && linha[14] !== null;

      if (estaBloqueado || temDesligamento) continue;

      var dtNascimento = converterDataFlexivel(linha[5]);
      var dtContratacao = converterDataFlexivel(linha[6]);

      var base = {
        linhaPlanilha: i + 1,
        nome: linha[0] || "",
        email: linha[1] || "",
        celular: linha[4] || "",
        cargo: linha[8] || "",
        loja: linha[11] || "Sem Loja",
        supervisor: linha[13] || "Sem Supervisor"
      };

      if (dtNascimento &&
          dtNascimento.getMonth() === hoje.getMonth() &&
          dtNascimento.getDate() === hoje.getDate()) {
        aniversarios.push({
          linhaPlanilha: base.linhaPlanilha,
          nome: base.nome,
          email: base.email,
          celular: base.celular,
          cargo: base.cargo,
          loja: base.loja,
          supervisor: base.supervisor,
          dataNascimento: formatarDataExibicao(dtNascimento),
          dataNascimentoISO: obterDataISO(dtNascimento)
        });
      }

      if (dtContratacao &&
          dtContratacao.getMonth() === hoje.getMonth() &&
          dtContratacao.getDate() === hoje.getDate()) {
        var anos = hoje.getFullYear() - dtContratacao.getFullYear();

        aniversariosEmpresa.push({
          linhaPlanilha: base.linhaPlanilha,
          nome: base.nome,
          email: base.email,
          celular: base.celular,
          cargo: base.cargo,
          loja: base.loja,
          supervisor: base.supervisor,
          dataContratacao: formatarDataExibicao(dtContratacao),
          dataContratacaoISO: obterDataISO(dtContratacao),
          anosEmpresa: anos,
          tempoEmpresa: anos === 1 ? "1 ano" : anos + " anos"
        });
      }
    }

    aniversarios.sort(function(a, b) {
      return normalizarTexto(a.nome).localeCompare(normalizarTexto(b.nome), "pt-BR");
    });

    aniversariosEmpresa.sort(function(a, b) {
      return normalizarTexto(a.nome).localeCompare(normalizarTexto(b.nome), "pt-BR");
    });

    return {
      sucesso: true,
      dataReferencia: formatarDataExibicao(hoje),
      aniversarios: aniversarios,
      aniversariosEmpresa: aniversariosEmpresa
    };

  } catch (e) {
    return {
      sucesso: false,
      mensagem: "Erro ao buscar aniversariantes: " + e.message
    };
  }
}

/**
 * Converte datas vindas da planilha em Date sem quebrar valores vazios.
 */
function converterDataFlexivel(valor) {
  if (!valor && valor !== 0) return null;

  if (valor instanceof Date && !isNaN(valor.getTime())) {
    var dt = new Date(valor);
    dt.setHours(0, 0, 0, 0);
    return dt;
  }

  if (typeof valor === "number" && isFinite(valor)) {
    var serial = new Date(Date.UTC(1899, 11, 30) + Math.round(valor * 86400000));
    if (!isNaN(serial.getTime())) {
      return new Date(
        serial.getUTCFullYear(),
        serial.getUTCMonth(),
        serial.getUTCDate()
      );
    }
  }

  var texto = String(valor).trim();
  if (!texto) return null;

  var iso = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    var d1 = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
    if (!isNaN(d1.getTime())) {
      d1.setHours(0, 0, 0, 0);
      return d1;
    }
  }

  var br = texto.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (br) {
    var d2 = new Date(parseInt(br[3], 10), parseInt(br[2], 10) - 1, parseInt(br[1], 10));
    if (!isNaN(d2.getTime())) {
      d2.setHours(0, 0, 0, 0);
      return d2;
    }
  }

  var generica = new Date(texto);
  if (!isNaN(generica.getTime())) {
    generica.setHours(0, 0, 0, 0);
    return generica;
  }

  return null;
}

// AUXILIARES DE TRATAMENTO DE DATAS
function formatarDataExibicao(d) {
  if (!d) return "";
  var dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt.getTime())) return String(d);
  var dia = String(dt.getDate()).padStart(2, '0');
  var mes = String(dt.getMonth() + 1).padStart(2, '0');
  var ano = dt.getFullYear();
  return dia + '/' + mes + '/' + ano;
}

function obterDataISO(d) {
  if (!d) return "";
  var dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt.getTime())) return "";
  var dia = String(dt.getDate()).padStart(2, '0');
  var mes = String(dt.getMonth() + 1).padStart(2, '0');
  var ano = dt.getFullYear();
  return ano + '-' + mes + '-' + dia;
}

function formatarDataISO(strData) {
  if (!strData) return "";

  if (strData instanceof Date && !isNaN(strData.getTime())) {
    return strData;
  }

  if (typeof strData === "number" && isFinite(strData)) {
    var serial = new Date(Date.UTC(1899, 11, 30) + Math.round(strData * 86400000));
    if (!isNaN(serial.getTime())) {
      return new Date(serial.getUTCFullYear(), serial.getUTCMonth(), serial.getUTCDate());
    }
  }

  var texto = String(strData).trim();
  if (!texto) return "";

  var br = texto.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (br) {
    var dtBR = new Date(parseInt(br[3], 10), parseInt(br[2], 10) - 1, parseInt(br[1], 10));
    if (!isNaN(dtBR.getTime())) return dtBR;
  }

  var iso = texto.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    var dtISO = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
    if (!isNaN(dtISO.getTime())) return dtISO;
  }

  var dt = new Date(texto);
  if (isNaN(dt.getTime())) return strData;
  return dt;
}