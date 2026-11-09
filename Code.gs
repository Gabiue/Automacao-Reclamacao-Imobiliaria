/**
 * SISTEMA DE RECLAMAÇÕES IMOBILIÁRIAS
 * @version 2.0 (abas mensais + trigger automático + checkbox "Resolvido")
 *
 * FLUXO:
 *   1. O Forms escreve na aba "Respostas do formulário 1" (caixa de entrada).
 *   2. O trigger aoReceberReclamacao() dispara, move a linha nova para a aba
 *      do mês correspondente (ex: "09-2026") e limpa a caixa de entrada.
 *   3. Em seguida, o trigger JÁ gera o PDF e envia o e-mail da linha nova
 *      (PROCESSAR_AUTOMATICAMENTE). Se falhar, a linha fica pendente e o
 *      dono do script recebe um aviso. O menu continua disponível pra
 *      processar manualmente quando precisar.
 *   4. Coluna K = STATUS (lógica do sistema, não mexer).
 *      Coluna L = RESOLVIDO (checkbox que a construtora marca ao concluir).
 */

// ========================================
// 1. CONFIGURAÇÕES E CONSTANTES
// ========================================

// Valores sensíveis (ID da pasta do Drive e e-mails) NÃO ficam no código.
// Eles são cadastrados em Configurações do projeto → Propriedades do script
// e lidos com lerConfig(). Veja a lista no README.
//
// Obrigatórias:
//   ID_PASTA_RECLAMACOES     → ID da pasta do Drive onde os PDFs são salvos
//   EMAIL_EQUIPE_MANUTENCAO  → e-mail que recebe o termo em PDF
// Opcional:
//   EMAIL_AVISO_FALHA        → quem recebe o aviso de falha (em branco = dono do script)

/**
 * Lê uma Propriedade do script obrigatória pelo nome.
 * @param {string} nome - ex: "ID_PASTA_RECLAMACOES"
 * @return {string} o valor cadastrado
 */
function lerConfig(nome) {
  const valor = (PropertiesService.getScriptProperties().getProperty(nome) || "").trim();
  if (!valor) {
    throw new Error(`Propriedade do script "${nome}" não configurada. ` +
      `Cadastre em Configurações do projeto → Propriedades do script.`);
  }
  return valor;
}

// Caixa de entrada do Google Forms (o Forms sempre escreve aqui)
const NOME_ABA_RESPOSTAS = "Respostas do formulário 1";

const COLUNAS = {
  TIMESTAMP: 0, NOME_COMPLETO: 1, CPF: 2, NUMERO_CONTRATO: 3,
  ENDERECO: 4, TELEFONE: 5, CONDOMINIO_APTO: 6, DESCRICAO_PROBLEMA: 7,
  FOTOS_PROBLEMA: 8, FOTOS_CONTRATO: 9
};

// Quantidade de colunas que vêm do formulário (A até J)
const NUM_COLUNAS_FORMULARIO = Object.keys(COLUNAS).length;

const COLUNA_STATUS = 11;     // Coluna K — base da lógica do sistema (NÃO MEXER)
const COLUNA_RESOLVIDO = 12;  // Coluna L — checkbox marcado pela construtora
const STATUS_PROCESSADO = "PROCESSADO";

// Colunas "numéricas" que o Google Sheets tenta converter pra número
// sozinho, o que DERRUBA zeros à esquerda (ex: CPF "041..." vira "41...").
// Forçamos essas colunas como Texto Simples ("@") pra impedir isso.
// Números de coluna em base 1 (igual ao getRange).
const COLUNAS_TEXTO_SIMPLES = [
  COLUNAS.CPF + 1,             // Coluna C
  COLUNAS.NUMERO_CONTRATO + 1, // Coluna D
  COLUNAS.TELEFONE + 1         // Coluna F
];

// Formato do nome das abas mensais. "/" não é permitido em nome de aba,
// por isso usamos "-": ex. "09-2026"
const FORMATO_NOME_ABA_MES = "MM-yyyy";

// Nome da função que o trigger "ao enviar formulário" chama
const NOME_FUNCAO_TRIGGER = "aoReceberReclamacao";

// Aviso curto de "chegou reclamação nova".
// DESLIGADO: com PROCESSAR_AUTOMATICAMENTE = true, o e-mail com o PDF já
// chega segundos depois, então o aviso curto viraria e-mail duplicado.
// Para religar (ex: se desligar o processamento automático), ponha true.
const NOTIFICAR_CHEGADA = false;

// Gera o PDF e envia o e-mail automaticamente assim que a reclamação chega,
// sem precisar abrir a planilha. Processa APENAS as linhas recém-chegadas —
// nunca o histórico antigo. Ponha false pra voltar ao modo 100% manual.
const PROCESSAR_AUTOMATICAMENTE = true;

// Se o processamento automático falhar, avisa o dono do script por e-mail.
const AVISAR_FALHA_AUTOMATICA = true;

const DESTINOS_ENVIO = {
  email: true,
  telegram: false,  // para o futuro
  whatsapp: false   // para o futuro
};

// ========================================
// 2. FUNÇÕES PRINCIPAIS (PROCESSAMENTO)
// ========================================

/**
 * Processa TODAS as reclamações pendentes da ABA DO MÊS ATUAL numa única execução.
 *
 * Faz 1 única leitura em lote da coluna de status pra achar todas as
 * linhas pendentes e depois processa uma por uma, na ordem da planilha.
 *
 * Se a execução for interrompida no meio (limite de tempo do Apps Script),
 * a linha em andamento continua pendente e será reprocessada na próxima
 * execução, já que o e-mail só é enviado ANTES de marcar como processado.
 */
function processarTodasPendentes() {
  const separador = "=".repeat(60);
  console.log("🚀 PROCESSANDO TODAS AS RECLAMAÇÕES PENDENTES...");
  console.log(separador);

  const aba = obterAbaMesAtual();
  console.log(`📅 Aba do mês atual: "${aba.getName()}"`);

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha <= 1) {
    console.log("✅ Nenhuma reclamação encontrada nesta aba!");
    console.log(separador);
    return { sucesso: true, processadas: 0, falhas: 0 };
  }

  // 1 ÚNICA chamada remota: lê a coluna de status inteira de uma vez
  const statusValores = aba.getRange(2, COLUNA_STATUS, ultimaLinha - 1, 1).getValues();

  // Filtra em memória -> lista exata dos números de linha pendentes
  const linhasPendentes = [];
  statusValores.forEach((linha, i) => {
    if (linha[0] !== STATUS_PROCESSADO) {
      linhasPendentes.push(i + 2); // +2 porque os dados começam na linha 2
    }
  });

  if (linhasPendentes.length === 0) {
    console.log("✅ Nenhuma reclamação pendente para processar!");
    console.log(separador);
    return { sucesso: true, processadas: 0, falhas: 0 };
  }

  console.log(`📋 ${linhasPendentes.length} pendente(s) encontrada(s): linhas ${linhasPendentes.join(', ')}`);
  console.log(separador);

  let processadas = 0;
  let falhas = 0;
  const resultados = [];

  for (let indice = 0; indice < linhasPendentes.length; indice++) {
    const numeroLinha = linhasPendentes[indice];
    console.log(`▶️ [${indice + 1}/${linhasPendentes.length}] Processando linha ${numeroLinha}...`);

    const resultado = _processarLinha(aba, numeroLinha);
    resultados.push({ linha: numeroLinha, resultado: resultado });

    if (resultado.sucesso) {
      processadas++;
      console.log(`✅ Linha ${numeroLinha} concluída.`);
    } else {
      falhas++;
      console.log(`❌ Linha ${numeroLinha} falhou: ${resultado.erro || resultado.motivo}`);
    }
    console.log(separador);
  }

  console.log("🎉 LOTE CONCLUÍDO!");
  console.log(`📊 Processadas com sucesso: ${processadas}`);
  console.log(`📊 Falhas: ${falhas}`);
  console.log(separador);

  return { sucesso: true, processadas: processadas, falhas: falhas, detalhes: resultados };
}

/**
 * Processa a próxima reclamação com status em branco (pendente) da aba do mês atual.
 */
function processarProximaPendente() {
  const separador = "=".repeat(60);
  console.log("🚀 PROCESSANDO PRÓXIMA RECLAMAÇÃO PENDENTE...");
  console.log(separador);

  const aba = obterAbaMesAtual();
  console.log(`📅 Aba do mês atual: "${aba.getName()}"`);

  const ultimaLinha = aba.getLastRow();
  let linhaPendente = null;

  if (ultimaLinha > 1) {
    const statusValores = aba.getRange(2, COLUNA_STATUS, ultimaLinha - 1, 1).getValues();
    for (let i = 0; i < statusValores.length; i++) {
      if (statusValores[i][0] !== STATUS_PROCESSADO) {
        linhaPendente = i + 2;
        break;
      }
    }
  }

  if (!linhaPendente) {
    console.log("✅ Nenhuma reclamação pendente para processar!");
    console.log(separador);
    return { sucesso: false, motivo: "Nenhuma pendente" };
  }

  console.log(`📋 Processando linha ${linhaPendente}...`);
  return _processarLinha(aba, linhaPendente);
}

/**
 * Função interna que executa o processamento completo de uma linha
 * de uma aba específica (gera PDF, salva no Drive, envia e-mail, marca status).
 */
function _processarLinha(aba, numeroLinha) {
  const separador = "=".repeat(60);
  console.log("🚀 SISTEMA DE RECLAMAÇÕES - Iniciando processamento...");
  console.log(separador);

  try {
    const dadosBrutos = obterDadosPlanilha(aba, numeroLinha);
    const dadosFormatados = formatarDados(dadosBrutos);
    const imagens = processarImagens(dadosFormatados);
    const documentoPDF = gerarPDF(dadosFormatados, imagens);

    const pasta = DriveApp.getFolderById(lerConfig("ID_PASTA_RECLAMACOES"));
    const arquivoSalvo = pasta.createFile(documentoPDF.blob.setName(documentoPDF.nome));

    // IMPORTANTE: o e-mail é enviado ANTES de marcar como processado.
    // Se a execução morrer entre essas duas linhas, a reclamação pode
    // ser reenviada por e-mail na próxima execução (duplicado, raro),
    // mas NUNCA fica marcada como processada sem o e-mail ter saído.
    enviarDocumento(dadosFormatados, arquivoSalvo);
    marcarComoProcessado(aba, dadosBrutos.numeroLinha);

    console.log(separador);
    console.log("🎉 PROCESSAMENTO CONCLUÍDO!");
    console.log(`📄 Arquivo: ${arquivoSalvo.getName()}`);
    console.log(`🔗 Link: ${arquivoSalvo.getUrl()}`);
    console.log(separador);

    return {
      sucesso: true,
      arquivo: {
        nome: arquivoSalvo.getName(),
        url: arquivoSalvo.getUrl(),
        id: arquivoSalvo.getId()
      }
    };

  } catch (erro) {
    console.log(separador);
    console.log("❌ ERRO NO PROCESSAMENTO");
    console.log(`💥 ${erro.message}`);
    console.log(separador);
    return { sucesso: false, erro: erro.message };
  }
}

// ========================================
// 3. MÓDULO: ORGANIZAÇÃO POR MÊS (TRIGGER)
// ========================================

/**
 * Chamada automaticamente pelo trigger "ao enviar formulário".
 * (Instale com instalarTriggerAutomatico() — item do menu.)
 *
 * Em vez de confiar no número da linha do evento (que pode mudar se
 * duas respostas chegarem quase ao mesmo tempo), ela simplesmente
 * esvazia a caixa de entrada inteira, movendo cada linha pra sua aba.
 */
function aoReceberReclamacao(e) {
  console.log("📨 Nova resposta recebida no formulário.");
  const resultado = organizarRespostasPorMes();

  if (!resultado || !resultado.sucesso) return;

  // Aviso curto de chegada (normalmente desligado — ver NOTIFICAR_CHEGADA).
  // Envolvido em try/catch pra que uma falha no e-mail NUNCA atrapalhe a
  // organização, que é a parte crítica. Fica AQUI e não dentro de
  // organizarRespostasPorMes pra que a organização MANUAL do histórico
  // não dispare dezenas de e-mails de uma vez.
  try {
    if (NOTIFICAR_CHEGADA && resultado.nomes && resultado.nomes.length > 0) {
      notificarChegada(resultado.nomes);
    }
  } catch (erro) {
    console.log(`⚠️ Aviso de chegada não pôde ser enviado: ${erro.message}`);
  }

  // Processamento automático: gera PDF e envia e-mail na hora.
  try {
    if (PROCESSAR_AUTOMATICAMENTE && resultado.linhasCriadas && resultado.linhasCriadas.length > 0) {
      processarLinhasRecemChegadas(resultado.linhasCriadas);
    }
  } catch (erro) {
    console.log(`⚠️ Erro no processamento automático: ${erro.message}`);
  }
}

/**
 * Processa APENAS as linhas que acabaram de ser movidas pelo trigger.
 *
 * CRÍTICO: não usa processarTodasPendentes() de propósito. Como o
 * histórico antigo está todo com STATUS em branco, aquela função
 * reprocessaria dezenas de reclamações velhas e mandaria um e-mail
 * pra cada uma. Aqui só tocamos nas linhas recém-criadas.
 *
 * @param {Array} linhasCriadas - [{ nomeAba: "09-2026", linha: 5 }, ...]
 */
function processarLinhasRecemChegadas(linhasCriadas) {
  const separador = "=".repeat(60);
  console.log(`⚙️ PROCESSAMENTO AUTOMÁTICO: ${linhasCriadas.length} linha(s) nova(s)`);
  console.log(separador);

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const falhas = [];
  let processadas = 0;

  linhasCriadas.forEach(item => {
    const aba = planilha.getSheetByName(item.nomeAba);
    if (!aba) {
      falhas.push({ aba: item.nomeAba, linha: item.linha, erro: "Aba não encontrada" });
      return;
    }

    const resultado = _processarLinha(aba, item.linha);
    if (resultado.sucesso) {
      processadas++;
    } else {
      falhas.push({ aba: item.nomeAba, linha: item.linha, erro: resultado.erro || "erro desconhecido" });
    }
  });

  console.log(separador);
  console.log(`✅ Processadas automaticamente: ${processadas} | ❌ Falhas: ${falhas.length}`);
  console.log(separador);

  // Avisa o dono se algo falhou (a linha continua PENDENTE e pode ser
  // reprocessada pelo menu depois de corrigir o problema).
  if (falhas.length > 0 && AVISAR_FALHA_AUTOMATICA) {
    try {
      notificarFalhaAutomatica(falhas);
    } catch (erro) {
      console.log(`⚠️ Não foi possível avisar sobre as falhas: ${erro.message}`);
    }
  }

  return { processadas: processadas, falhas: falhas.length };
}

/**
 * Avisa o dono do script que o processamento automático falhou em
 * uma ou mais linhas, dizendo exatamente onde e por quê.
 */
function notificarFalhaAutomatica(falhas) {
  // EMAIL_AVISO_FALHA é opcional: sem ela, usa o e-mail de quem instalou
  // o trigger (ou seja, o dono do script) — sem precisar configurar nada.
  const destinatario = PropertiesService.getScriptProperties().getProperty("EMAIL_AVISO_FALHA")
    || Session.getEffectiveUser().getEmail();
  if (!destinatario) {
    console.log("⚠️ Sem destinatário para o aviso de falha.");
    return;
  }

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const assunto = `[AÇÃO NECESSÁRIA] Falha ao processar ${falhas.length} reclamação(ões)`;

  const corpo =
    `O processamento automático falhou nas linhas abaixo.\n` +
    `Elas continuam PENDENTES e nenhum e-mail foi enviado à equipe.\n\n` +
    falhas.map(f => `• Aba "${f.aba}", linha ${f.linha}\n  Motivo: ${f.erro}`).join('\n\n') +
    `\n\nCorrija o problema na planilha e depois use o menu\n` +
    `🏠 Reclamações → ⏩ Processar TODAS Pendentes.\n\n` +
    `Planilha: ${planilha.getUrl()}\n\n--\nSistema Automático de Reclamações`;

  GmailApp.sendEmail(destinatario, assunto, corpo);
  console.log(`📧 Aviso de falha enviado para ${destinatario}`);
}

/**
 * Envia um e-mail curto avisando que chegou reclamação nova, dizendo de
 * quem é e pedindo pra conferir a planilha. Não anexa nada — o PDF
 * completo continua indo só quando você processa pelo menu.
 */
function notificarChegada(nomes) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const destinatario = lerConfig("EMAIL_EQUIPE_MANUTENCAO");
  const ehPlural = nomes.length > 1;

  const assunto = ehPlural
    ? `Nova reclamação recebida (${nomes.length})`
    : `Nova reclamação recebida - ${nomes[0]}`;

  const corpo = ehPlural
    ? `Chegaram ${nomes.length} reclamações novas pelo formulário:\n\n` +
      nomes.map(n => `- ${n}`).join('\n') +
      `\n\nAcesse a planilha para conferir:\n${planilha.getUrl()}\n\n--\nSistema Automático de Reclamações`
    : `Chegou uma reclamação nova de ${nomes[0]}.\n\n` +
      `Acesse a planilha para conferir:\n${planilha.getUrl()}\n\n--\nSistema Automático de Reclamações`;

  GmailApp.sendEmail(destinatario, assunto, corpo);
  console.log(`📧 Aviso de chegada enviado para ${destinatario} (${nomes.length} reclamação(ões))`);
}

/**
 * Move TODAS as linhas da caixa de entrada ("Respostas do formulário 1")
 * para as abas mensais correspondentes (ex: "09-2026"), criando a aba
 * se ainda não existir. Depois remove as linhas movidas da caixa de entrada.
 *
 * Pode ser rodada manualmente pelo menu (útil pra organizar o histórico
 * antigo na primeira vez).
 */
function organizarRespostasPorMes() {
  const separador = "=".repeat(60);
  console.log("🗂️ ORGANIZANDO RESPOSTAS POR MÊS...");
  console.log(separador);

  // Lock: evita que duas execuções (ex: duas respostas quase simultâneas)
  // mexam na caixa de entrada ao mesmo tempo.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (erro) {
    console.log("⏳ Outra execução já está organizando. Saindo.");
    return { sucesso: false, motivo: "Lock ocupado" };
  }

  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaEntrada = planilha.getSheetByName(NOME_ABA_RESPOSTAS);
    if (!abaEntrada) throw new Error(`Aba "${NOME_ABA_RESPOSTAS}" não encontrada`);

    const ultimaLinha = abaEntrada.getLastRow();
    if (ultimaLinha <= 1) {
      console.log("✅ Caixa de entrada vazia. Nada para organizar.");
      console.log(separador);
      return { sucesso: true, movidas: 0, nomes: [], linhasCriadas: [] };
    }

    const qtdLinhas = ultimaLinha - 1;
    const valores = abaEntrada.getRange(2, 1, qtdLinhas, NUM_COLUNAS_FORMULARIO).getValues();
    const fusoHorario = planilha.getSpreadsheetTimeZone();

    // Agrupa as linhas por aba de mês (em memória, sem I/O)
    const grupos = {}; // { "09-2026": [linha, linha, ...] }
    const nomes = [];  // nomes de quem enviou, na ordem, pro aviso de chegada
    valores.forEach((linha, i) => {
      // Ignora linhas totalmente vazias
      if (linha.every(v => v === "" || v === null)) return;

      let dataRef = linha[COLUNAS.TIMESTAMP];
      if (!(dataRef instanceof Date) || isNaN(dataRef.getTime())) {
        console.log(`   ⚠️ Linha ${i + 2}: timestamp inválido, usando data de hoje.`);
        dataRef = new Date();
      }

      const nome = linha[COLUNAS.NOME_COMPLETO];
      nomes.push(nome ? nome.toString().trim() : "(nome não informado)");

      const nomeAba = nomeAbaDoMes(dataRef, fusoHorario);
      if (!grupos[nomeAba]) grupos[nomeAba] = [];
      grupos[nomeAba].push(linha.map(neutralizarFormula));
    });

    let totalMovidas = 0;
    // Onde cada linha foi parar — usado pelo processamento automático pra
    // tocar SÓ nas linhas novas, nunca no histórico antigo.
    const linhasCriadas = []; // [{ nomeAba: "09-2026", linha: 5 }, ...]

    Object.keys(grupos).forEach(nomeAba => {
      const linhasDoMes = grupos[nomeAba];
      const abaMes = obterOuCriarAbaMes(planilha, nomeAba, abaEntrada);

      const linhaInicial = abaMes.getLastRow() + 1;
      // Escreve todas as linhas do mês de uma vez (1 chamada)
      abaMes.getRange(linhaInicial, 1, linhasDoMes.length, NUM_COLUNAS_FORMULARIO).setValues(linhasDoMes);
      // Cria o checkbox "Resolvido" (desmarcado) em cada linha nova
      abaMes.getRange(linhaInicial, COLUNA_RESOLVIDO, linhasDoMes.length, 1).insertCheckboxes();

      // Registra o número exato de cada linha escrita
      for (let k = 0; k < linhasDoMes.length; k++) {
        linhasCriadas.push({ nomeAba: nomeAba, linha: linhaInicial + k });
      }

      totalMovidas += linhasDoMes.length;
      console.log(`   ✅ ${linhasDoMes.length} linha(s) → aba "${nomeAba}" (a partir da linha ${linhaInicial})`);
    });

    // Remove da caixa de entrada EXATAMENTE as linhas que foram lidas.
    // Se uma resposta nova chegou enquanto isso, ela está abaixo dessas
    // e vai "subir" pra linha 2 — o trigger dela cuida depois.
    abaEntrada.deleteRows(2, qtdLinhas);

    console.log(separador);
    console.log(`🎉 ${totalMovidas} reclamação(ões) organizada(s) por mês.`);
    console.log(separador);
    return { sucesso: true, movidas: totalMovidas, nomes: nomes, linhasCriadas: linhasCriadas };

  } catch (erro) {
    console.log(`❌ Erro ao organizar: ${erro.message}`);
    console.log(separador);
    return { sucesso: false, erro: erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Retorna a aba do mês informado; cria se não existir (com cabeçalho,
 * colunas STATUS e RESOLVIDO, primeira linha congelada, e as colunas
 * sensíveis já protegidas como Texto Simples).
 */
function obterOuCriarAbaMes(planilha, nomeAba, abaOrigemCabecalho) {
  let aba = planilha.getSheetByName(nomeAba);
  if (aba) return aba;

  console.log(`   🆕 Criando aba "${nomeAba}"...`);
  aba = planilha.insertSheet(nomeAba);

  // Cabeçalho = cabeçalho do formulário (A-J) + STATUS + RESOLVIDO
  const cabecalhoForm = abaOrigemCabecalho.getRange(1, 1, 1, NUM_COLUNAS_FORMULARIO).getValues()[0];
  const cabecalho = cabecalhoForm.concat(["STATUS", "RESOLVIDO"]);
  aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight("bold");
  aba.setFrozenRows(1);

  // Protege CPF, Nº CONTRATO e TELEFONE contra a auto-conversão pra número
  // (que derruba zeros à esquerda) já na criação da aba.
  forcarColunasComoTexto(aba);

  return aba;
}

/**
 * Converte número de coluna (1 = A, 2 = B, ...) para letra.
 * Suporta até coluna ZZ, mais que suficiente pra esta planilha.
 */
function numeroParaLetraColuna(numero) {
  let letra = "";
  while (numero > 0) {
    const resto = (numero - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    numero = Math.floor((numero - resto) / 26);
  }
  return letra;
}

/**
 * Aplica formato "Texto Simples" (@) na COLUNA INTEIRA de cada coluna
 * listada em COLUNAS_TEXTO_SIMPLES. Formatar a coluna inteira (não só
 * as linhas já preenchidas) garante que RESPOSTAS FUTURAS do Forms
 * também cheguem como texto, preservando zeros à esquerda.
 */
function forcarColunasComoTexto(aba) {
  COLUNAS_TEXTO_SIMPLES.forEach(numeroColuna => {
    const letra = numeroParaLetraColuna(numeroColuna);
    aba.getRange(`${letra}:${letra}`).setNumberFormat("@");
  });
  console.log(`   🔒 Colunas CPF/CONTRATO/TELEFONE fixadas como Texto Simples em "${aba.getName()}"`);
}

/**
 * Corrige retroativamente TODAS as abas já existentes (a caixa de entrada
 * e todas as abas de mês já criadas), forçando CPF/CONTRATO/TELEFONE
 * como Texto Simples. Rode isso UMA VEZ pelo menu para proteger o que
 * já existe — não recupera zeros já perdidos, só evita perder mais.
 */
function corrigirFormatoColunasExistentes() {
  const separador = "=".repeat(60);
  console.log("🔒 APLICANDO TEXTO SIMPLES EM CPF/CONTRATO/TELEFONE...");
  console.log(separador);

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const todasAbas = planilha.getSheets();
  let corrigidas = 0;

  todasAbas.forEach(aba => {
    const nome = aba.getName();
    const ehCaixaEntrada = nome === NOME_ABA_RESPOSTAS;
    const ehAbaDeMes = /^\d{2}-\d{4}$/.test(nome);

    if (!ehCaixaEntrada && !ehAbaDeMes) {
      return; // ignora outras abas que não são deste sistema
    }

    forcarColunasComoTexto(aba);
    corrigidas++;
  });

  console.log(separador);
  console.log(`🎉 ${corrigidas} aba(s) protegida(s) contra perda de zero à esquerda.`);
  console.log("⚠️ Isso NÃO recupera zeros já perdidos em dados antigos — só evita novas perdas.");
  console.log(separador);
  _alertar(`${corrigidas} aba(s) corrigida(s)! ✅\nCPF, Nº Contrato e Telefone agora ficam como texto.\n\nAtenção: isso não recupera zeros já perdidos em respostas antigas.`);
  return { sucesso: true, abasCorrigidas: corrigidas };
}

/**
 * Nome da aba correspondente a uma data. Ex: 15/09/2026 → "09-2026"
 */
function nomeAbaDoMes(data, fusoHorario) {
  const tz = fusoHorario || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(data, tz, FORMATO_NOME_ABA_MES);
}

/**
 * Retorna a aba do MÊS ATUAL (calculado pela data de hoje).
 * Lança erro amigável se ela ainda não existir.
 */
function obterAbaMesAtual() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = nomeAbaDoMes(new Date(), planilha.getSpreadsheetTimeZone());
  const aba = planilha.getSheetByName(nomeAba);
  if (!aba) {
    throw new Error(`Aba do mês atual "${nomeAba}" não existe ainda. ` +
      `Nenhuma reclamação chegou neste mês (ou rode "Organizar respostas por mês").`);
  }
  return aba;
}

/**
 * Instala o trigger "ao enviar formulário" que chama aoReceberReclamacao().
 * Seguro rodar mais de uma vez: não duplica o trigger.
 */
function instalarTriggerAutomatico() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const jaExiste = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === NOME_FUNCAO_TRIGGER &&
    t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT
  );

  if (jaExiste) {
    console.log("✅ Trigger automático já está instalado.");
    _alertar("Trigger automático já está instalado. ✅");
    return;
  }

  ScriptApp.newTrigger(NOME_FUNCAO_TRIGGER)
    .forSpreadsheet(planilha)
    .onFormSubmit()
    .create();

  // Protege a caixa de entrada também, pra respostas novas já chegarem como texto.
  const abaEntrada = planilha.getSheetByName(NOME_ABA_RESPOSTAS);
  if (abaEntrada) forcarColunasComoTexto(abaEntrada);

  console.log("🎉 Trigger automático instalado! Novas respostas serão organizadas por mês sozinhas.");
  _alertar("Trigger instalado! 🎉\nNovas reclamações serão movidas automaticamente para a aba do mês.");
}

/** Mostra um alerta na planilha (ignora se rodando sem interface) */
function _alertar(mensagem) {
  try { SpreadsheetApp.getUi().alert(mensagem); } catch (e) {}
}

// ========================================
// 4. MÓDULO: OBTENÇÃO E FORMATAÇÃO DE DADOS
// ========================================

function obterUltimaLinha(aba) {
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha <= 1) throw new Error(`Nenhuma reclamação encontrada na aba "${aba.getName()}"`);
  return ultimaLinha;
}

function obterDadosPlanilha(aba, numeroLinha) {
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha <= 1) throw new Error("Nenhuma reclamação encontrada");
  if (numeroLinha > ultimaLinha || numeroLinha < 2) throw new Error(`Linha ${numeroLinha} não existe ou é inválida`);

  const valores = aba.getRange(numeroLinha, 1, 1, NUM_COLUNAS_FORMULARIO).getValues()[0];

  return {
    numeroLinha: numeroLinha,
    timestamp: valores[COLUNAS.TIMESTAMP],
    nomeCompleto: valores[COLUNAS.NOME_COMPLETO],
    cpf: valores[COLUNAS.CPF],
    numeroContrato: valores[COLUNAS.NUMERO_CONTRATO],
    endereco: valores[COLUNAS.ENDERECO],
    telefone: valores[COLUNAS.TELEFONE],
    condominioApartamento: valores[COLUNAS.CONDOMINIO_APTO],
    descricaoProblema: valores[COLUNAS.DESCRICAO_PROBLEMA],
    fotosProblema: valores[COLUNAS.FOTOS_PROBLEMA],
    fotosContrato: valores[COLUNAS.FOTOS_CONTRATO]
  };
}

function formatarDados(dados) {
  const camposObrigatorios = ['nomeCompleto', 'cpf', 'numeroContrato', 'endereco', 'descricaoProblema'];
  const ausentes = camposObrigatorios.filter(campo => !dados[campo] || dados[campo].toString().trim() === '');
  if (ausentes.length > 0) throw new Error(`Campos ausentes: ${ausentes.join(', ')}`);

  return {
    nomeCompleto: dados.nomeCompleto.toString().trim().toUpperCase(),
    cpf: formatarCPF(dados.cpf),
    numeroContrato: formatarContrato(dados.numeroContrato),
    endereco: dados.endereco.toString().trim().toUpperCase(),
    telefone: dados.telefone || "",
    condominioApartamento: dados.condominioApartamento || "",
    descricaoProblema: dados.descricaoProblema || "",
    timestamp: dados.timestamp,
    linksImagensProblema: extrairLinks(dados.fotosProblema),
    linksImagensContrato: extrairLinks(dados.fotosContrato)
  };
}

// ========================================
// 5. MÓDULO: PROCESSAMENTO DE IMAGENS (OTIMIZADO)
// ========================================

function processarImagens(dados) {
  const linksProblema = dados.linksImagensProblema;
  const linksContrato = dados.linksImagensContrato;
  const todosLinks = [...linksProblema, ...linksContrato];

  if (todosLinks.length === 0) {
    return { problemas: [], contratos: [] };
  }

  // Baixa TODAS as imagens em paralelo de uma vez só
  const todasImagens = baixarImagensParalelo(todosLinks);

  return {
    problemas: todasImagens.slice(0, linksProblema.length),
    contratos: todasImagens.slice(linksProblema.length)
  };
}

function baixarImagensParalelo(urls) {
  if (urls.length === 0) return [];

  const token = ScriptApp.getOAuthToken(); // 1 chamada só, fora do loop

  // Extrai todos os IDs primeiro (operação local, sem I/O)
  const ids = urls.map(url => extrairIdDrive(url));

  // Monta todas as requisições e separa as válidas
  const indicesValidos = [];
  const requisicoesValidas = [];

  ids.forEach((id, i) => {
    if (!id) {
      console.log(`   ⚠️ Imagem ${i + 1}: ID não encontrado na URL`);
      return;
    }
    indicesValidos.push(i);
    requisicoesValidas.push({
      url: `https://drive.google.com/thumbnail?id=${id}&sz=w800`,
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
  });

  if (requisicoesValidas.length === 0) return [];

  console.log(`   🚀 Baixando ${requisicoesValidas.length} imagem(ns) em paralelo...`);

  // ✅ Uma única chamada de rede para todas as imagens
  const respostas = UrlFetchApp.fetchAll(requisicoesValidas);

  const resultado = new Array(urls.length).fill(null);

  respostas.forEach((resposta, i) => {
    const indiceOriginal = indicesValidos[i];
    try {
      if (resposta.getResponseCode() !== 200) {
        console.log(`   ⚠️ Imagem ${indiceOriginal + 1}: HTTP ${resposta.getResponseCode()}`);
        return;
      }
      const blob = resposta.getBlob().setContentType('image/jpeg');
      const tamanhoKB = Math.round(blob.getBytes().length / 1024);
      console.log(`   ✅ Imagem ${indiceOriginal + 1}: ${tamanhoKB}KB`);
      resultado[indiceOriginal] = {
        blob: blob,
        nome: `imagem_${indiceOriginal + 1}.jpg`,
        tamanho: tamanhoKB
      };
    } catch (erro) {
      console.log(`   ⚠️ Erro imagem ${indiceOriginal + 1}: ${erro.message}`);
    }
  });

  return resultado.filter(img => img !== null);
}

// ========================================
// 6. MÓDULO: GERAÇÃO DE PDF
// ========================================

function gerarPDF(dados, imagens) {
  let imagemHTML = "";
  if (imagens.problemas.length > 0) {
    const blob = imagens.problemas[0].blob;
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    imagemHTML = `
    <div style="margin-top:10px;">
      <p><strong>FOTO DO PROBLEMA:</strong></p>
      <img src="data:${mimeType};base64,${base64}" width="150" style="display:block; margin:auto;" />
    </div>`;
  }

  const htmlContent = `
  <!DOCTYPE html>
  <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; font-size: 12pt; margin: 40px; }
        h1 { text-align: center; font-size: 14pt; }
        table { width: 100%; margin-bottom: 20px; }
        td { padding: 4px; vertical-align: top; }
        .assinaturas { margin-top: 60px; text-align: center; }
        .linha-assinatura { margin-top: 40px; border-top: 1px solid #000; width: 300px; margin-left: auto; margin-right: auto; }
      </style>
    </head>
    <body>
      <h1>TERMO DE MANUTENÇÃO DE IMÓVEL</h1>

      <table>
        <tr><td><strong>SOLICITANTE:</strong> ${escaparHTML(dados.nomeCompleto)}</td><td><strong>CPF:</strong> ${escaparHTML(dados.cpf)}</td></tr>
        <tr><td><strong>Nº CONTRATO:</strong> ${escaparHTML(dados.numeroContrato)}</td><td><strong>APTO:</strong> ${escaparHTML(dados.condominioApartamento || '---')}</td></tr>
        <tr><td><strong>ENDEREÇO:</strong> ${escaparHTML(dados.endereco)}</td><td><strong>CONDOMÍNIO:</strong> ${escaparHTML(dados.condominioApartamento || '---')}</td></tr>
        <tr><td><strong>CONTATO:</strong> ${escaparHTML(dados.telefone || '---')}</td><td><strong>DT PEDIDO:</strong> ${formatarDataBR(dados.timestamp)}</td></tr>
      </table>

      <div>
        <p><strong>MOTIVO DA RECLAMAÇÃO</strong></p>
        <ul>
          <li>${escaparHTML(dados.descricaoProblema)}</li>
        </ul>
      </div>

      <div>
        <p><strong>ANEXOS:</strong></p>
      </div>

      ${imagemHTML}

      <div style="margin-top: 10px;">
        <p><strong>DATA CONCLUSÃO DE MANUTENÇÃO:</strong> ______/______/_________</p>
      </div>

      <div class="assinaturas">
        <div class="linha-assinatura">PROFISSIONAL CONSTRUTORA</div>
        <br><br>
        <div class="linha-assinatura">${escaparHTML(dados.nomeCompleto)}</div>
      </div>
    </body>
  </html>`;

  const blob = HtmlService.createHtmlOutput(htmlContent)
    .getAs(MimeType.PDF)
    .setName(gerarNomeArquivo(dados));

  return {
    blob: blob,
    nome: blob.getName(),
    tamanho: blob.getBytes().length
  };
}

// ========================================
// 7. MÓDULO: ENVIO (EMAIL, TELEGRAM, WHATSAPP)
// ========================================

function enviarDocumento(dados, arquivo) {
  console.log("📤 Iniciando envio do documento...");
  let enviosRealizados = 0;

  try {
    if (DESTINOS_ENVIO.email) {
      console.log("📧 Enviando por email...");
      enviarPorEmail(dados, arquivo);
      enviosRealizados++;
      console.log("✅ Email enviado com sucesso!");
    }

    if (DESTINOS_ENVIO.telegram) {
      console.log("📱 Enviando por Telegram...");
      enviarPorTelegram(dados, arquivo);
      enviosRealizados++;
      console.log("✅ Telegram enviado com sucesso!");
    }

    if (DESTINOS_ENVIO.whatsapp) {
      console.log("💬 Enviando por WhatsApp...");
      enviarPorWhatsApp(dados, arquivo);
      enviosRealizados++;
      console.log("✅ WhatsApp enviado com sucesso!");
    }

    console.log(`📤 Total de envios realizados: ${enviosRealizados}`);

  } catch (erro) {
    console.log(`❌ Erro no envio: ${erro.message}`);
    throw erro;
  }
}

function enviarPorEmail(dados, arquivo) {
  const assunto = `Nova Reclamação - ${dados.nomeCompleto} - Contrato ${dados.numeroContrato}`;

  const corpoEmail = `NOVA RECLAMAÇÃO PARA MANUTENÇÃO

DADOS DO CLIENTE:
- Nome: ${dados.nomeCompleto}
- CPF: ${dados.cpf}
- Contrato: ${dados.numeroContrato}
- Endereço: ${dados.endereco}
- Apartamento: ${dados.condominioApartamento || 'Não informado'}
- Telefone: ${dados.telefone || 'Não informado'}

PROBLEMA RELATADO:
${dados.descricaoProblema}

Data da Reclamação: ${formatarDataBR(dados.timestamp)}

O documento PDF está em anexo para impressão e assinatura.

--
Sistema Automático de Reclamações`;

  GmailApp.sendEmail(
    lerConfig("EMAIL_EQUIPE_MANUTENCAO"),
    assunto,
    corpoEmail,
    { attachments: [arquivo.getBlob()] }
  );
}

function enviarPorTelegram(dados, arquivo) { /* futuro */ }
function enviarPorWhatsApp(dados, arquivo) { /* futuro */ }

// ========================================
// 8. MÓDULO: STATUS NA PLANILHA
// ========================================

/**
 * Lê o status da coluna K de uma linha em uma aba específica.
 * Em branco = PENDENTE | "PROCESSADO" = PROCESSADO
 */
function verificarStatus(aba, numeroLinha) {
  const status = aba.getRange(numeroLinha, COLUNA_STATUS).getValue();
  return status === STATUS_PROCESSADO ? 'PROCESSADO' : 'PENDENTE';
}

/**
 * Escreve "PROCESSADO" na coluna K da linha correspondente.
 */
function marcarComoProcessado(aba, numeroLinha) {
  aba.getRange(numeroLinha, COLUNA_STATUS).setValue(STATUS_PROCESSADO);
  console.log(`🎉 Linha ${numeroLinha} da aba "${aba.getName()}" marcada como PROCESSADO`);
}

/**
 * Lista todas as reclamações pendentes (coluna K em branco) da aba do mês atual.
 */
function listarReclamacoesPendentes() {
  const aba = obterAbaMesAtual();
  console.log(`📋 RECLAMAÇÕES PENDENTES — aba "${aba.getName()}":`);

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha <= 1) {
    console.log("✅ Nenhuma reclamação nesta aba!");
    return;
  }

  // Lê tudo de uma vez: colunas A e B (timestamp, nome) + coluna K (status)
  const dados = aba.getRange(2, 1, ultimaLinha - 1, COLUNA_STATUS).getValues();
  let contador = 0;

  dados.forEach((linha, i) => {
    if (linha[COLUNA_STATUS - 1] !== STATUS_PROCESSADO) {
      const timestamp = linha[COLUNAS.TIMESTAMP];
      const nome = linha[COLUNAS.NOME_COMPLETO];
      console.log(`Linha ${i + 2}: ${nome} (${formatarDataBR(timestamp)})`);
      contador++;
    }
  });

  if (contador === 0) {
    console.log("✅ Nenhuma reclamação pendente!");
  } else {
    console.log(`📊 Total: ${contador} reclamação(ões) pendente(s)`);
  }
}

// ========================================
// 9. MÓDULO: UTILITÁRIOS
// ========================================

function formatarCPF(cpf) {
  const numeros = cpf.toString().replace(/\D/g, '');

  if (numeros.length < 11 && typeof cpf === 'number') {
    // Quase certo: o Sheets salvou a resposta como número e comeu o(s)
    // zero(s) à esquerda (ex: "041..." virou "41..."). Não dá pra
    // adivinhar o valor certo com segurança (documento oficial),
    // então paramos aqui e pedimos correção manual.
    throw new Error(
      `CPF com ${numeros.length} dígito(s) (esperado 11) — provável zero à ` +
      `esquerda perdido pelo Google Sheets. Corrija manualmente na planilha ` +
      `digitando o CPF certo com um apóstrofo na frente (ex: '0${numeros}) ` +
      `pra forçar texto, e rode o processamento de novo.`
    );
  }

  if (numeros.length !== 11) throw new Error(`CPF deve ter 11 dígitos (recebido: ${numeros.length})`);
  if (!validarCPF(numeros)) throw new Error("CPF inválido");
  return `${numeros.substr(0,3)}.${numeros.substr(3,3)}.${numeros.substr(6,3)}-${numeros.substr(9,2)}`;
}

function validarCPF(cpf) {
  return true; // substituir por validação real se necessário
}

function formatarContrato(contrato) {
  const numeros = contrato.toString().replace(/\D/g, '');
  return `${numeros.substr(0,1)}.${numeros.substr(1,4)}.${numeros.substr(5,7)}-${numeros.substr(12,1)}`;
}

// Texto do formulário vai dentro do HTML do PDF: sem escapar, "<h1>..." viraria tag e adulteraria o termo.
function escaparHTML(valor) {
  return String(valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// setValues transforma texto iniciado por = + - @ em fórmula (ex: =IMPORTXML vazando CPFs).
// O apóstrofo força texto e não aparece na célula nem no getValues.
function neutralizarFormula(valor) {
  return (typeof valor === 'string' && /^[=+\-@]/.test(valor)) ? "'" + valor : valor;
}

function extrairLinks(texto) {
  if (!texto) return [];
  return texto.toString().split(',')
    .map(link => link.trim())
    .filter(link => link.length > 0 && link.includes('drive.google.com'));
}

function extrairIdDrive(url) {
  if (!url) return null;
  let match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  return null;
}

function formatarDataBR(data) {
  if (!data) return '';
  const d = new Date(data);
  const dia = String(d.getDate()).padStart(2, '0');
  const mes = String(d.getMonth() + 1).padStart(2, '0');
  const ano = d.getFullYear();
  return `${dia}/${mes}/${ano}`;
}

function gerarNomeArquivo(dados) {
  const nome = dados.nomeCompleto.replace(/[^A-Z0-9\s]/g, '').replace(/\s+/g, '_').substring(0, 25);
  const agora = new Date();
  const timestamp = `${agora.getFullYear()}${String(agora.getMonth() + 1).padStart(2, '0')}${String(agora.getDate()).padStart(2, '0')}_${String(agora.getHours()).padStart(2, '0')}${String(agora.getMinutes()).padStart(2, '0')}`;
  return `Reclamacao_${nome}_${timestamp}.pdf`;
}

// ========================================
// 10. MÓDULO: INTERFACE/MENU
// ========================================

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🏠 Reclamações')
      .addItem('⏩ Processar TODAS Pendentes (mês atual)', 'processarTodasPendentes')
      .addItem('⏭️ Processar Próxima Pendente (mês atual)', 'processarProximaPendente')
      .addItem('📋 Listar Pendentes (mês atual)', 'listarReclamacoesPendentes')
      .addSeparator()
      .addItem('🗂️ Organizar respostas por mês (manual)', 'organizarRespostasPorMes')
      .addItem('⚡ Instalar trigger automático', 'instalarTriggerAutomatico')
      .addItem('🔒 Corrigir formato CPF/Contrato/Telefone', 'corrigirFormatoColunasExistentes')
      .addSeparator()
      .addItem('🧪 Testar Sistema', 'testarSistema')
      .addItem('📖 Ajuda', 'ajuda')
      .addSeparator()
      .addItem('🚀 Processar Última Reclamação (mês atual)', 'processarUltimaReclamacao')
      .addToUi();
  } catch (erro) {}
}

// ========================================
// 11. MÓDULO: TESTES E DEBUG
// ========================================

function testarSistema() {
  const separador = "=".repeat(50);
  console.log("🧪 TESTE DO SISTEMA");
  console.log(separador);
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();

    // 0. Propriedades do script
    lerConfig("ID_PASTA_RECLAMACOES");
    lerConfig("EMAIL_EQUIPE_MANUTENCAO");
    console.log("🔑 Propriedades do script: OK");

    // 1. Caixa de entrada
    const abaEntrada = planilha.getSheetByName(NOME_ABA_RESPOSTAS);
    if (!abaEntrada) {
      console.log(`❌ Aba "${NOME_ABA_RESPOSTAS}" não encontrada`);
      return false;
    }
    const naEntrada = Math.max(abaEntrada.getLastRow() - 1, 0);
    console.log(`📥 Caixa de entrada: ${naEntrada} resposta(s) aguardando organização`);

    // 2. Trigger
    const triggerOk = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === NOME_FUNCAO_TRIGGER);
    console.log(triggerOk ? "⚡ Trigger automático: instalado" : "⚠️ Trigger automático: NÃO instalado (use o menu)");

    // 3. Aba do mês atual
    const nomeMes = nomeAbaDoMes(new Date(), planilha.getSpreadsheetTimeZone());
    const abaMes = planilha.getSheetByName(nomeMes);
    if (!abaMes) {
      console.log(`⚠️ Aba do mês atual "${nomeMes}" ainda não existe`);
    } else {
      const qtd = Math.max(abaMes.getLastRow() - 1, 0);
      console.log(`📅 Aba "${nomeMes}": ${qtd} reclamação(ões)`);
      if (qtd > 0) {
        const dados = obterDadosPlanilha(abaMes, abaMes.getLastRow());
        formatarDados(dados);
        console.log("✅ Última linha do mês validada com sucesso");
      }
    }

    // 4. Pasta do Drive
    DriveApp.getFolderById(lerConfig("ID_PASTA_RECLAMACOES"));
    console.log("📁 Pasta do Drive: OK");

    console.log("✅ Sistema configurado corretamente!");
    console.log(separador);
    return true;
  } catch (erro) {
    console.log(`❌ Erro: ${erro.message}`);
    console.log(separador);
    return false;
  }
}

function ajuda() {
  const separador = "=".repeat(60);
  console.log("📖 SISTEMA DE RECLAMAÇÕES - GUIA DE USO (v2.0)");
  console.log(separador);
  console.log("");
  console.log("🗂️ ORGANIZAÇÃO POR MÊS:");
  console.log("  O Forms escreve na aba 'Respostas do formulário 1' (caixa de entrada).");
  console.log("  O trigger move cada resposta pra aba do mês (ex: '09-2026'),");
  console.log("  criando a aba automaticamente se não existir.");
  console.log("");
  console.log("🎯 FUNÇÕES PRINCIPAIS (sempre na aba do MÊS ATUAL):");
  console.log("  processarTodasPendentes()    - Processa TODAS as pendentes de uma vez");
  console.log("  processarProximaPendente()   - Processa a próxima em branco");
  console.log("  processarUltimaReclamacao()  - Gera PDF da última linha");
  console.log("  listarReclamacoesPendentes() - Lista pendentes no console");
  console.log("  organizarRespostasPorMes()   - Move manualmente a caixa de entrada pras abas");
  console.log("  instalarTriggerAutomatico()  - Instala o trigger 'ao enviar formulário'");
  console.log("  testarSistema()              - Verifica configuração");
  console.log("");
  console.log("⚙️ PROCESSAMENTO AUTOMÁTICO (PROCESSAR_AUTOMATICAMENTE = true):");
  console.log("  Ao chegar uma reclamação, o trigger move pra aba do mês E JÁ gera o");
  console.log("  PDF + envia o e-mail, sem você abrir a planilha.");
  console.log("  Ele processa SOMENTE as linhas recém-chegadas — o histórico antigo");
  console.log("  com STATUS em branco NUNCA é tocado automaticamente.");
  console.log("  Se falhar (ex: CPF com zero perdido), a linha fica PENDENTE, nenhum");
  console.log("  e-mail vai pra equipe, e você recebe um aviso dizendo qual linha falhou.");
  console.log("");
  console.log("📧 E-MAILS ENVIADOS PELO SISTEMA:");
  console.log("  1. TERMO EM PDF → automático ao chegar (ou manual, pelo menu).");
  console.log("  2. AVISO DE FALHA → só pro dono, quando o automático não consegue.");
  console.log("  3. AVISO DE CHEGADA → desligado (NOTIFICAR_CHEGADA = false), pois o");
  console.log("     PDF já chega na hora. Religue se desligar o processamento automático.");
  console.log("");
  console.log("📋 COLUNAS DE CONTROLE (em cada aba mensal):");
  console.log("  K = STATUS    → em branco = pendente | PROCESSADO = PDF já enviado (NÃO MEXER)");
  console.log("  L = RESOLVIDO → checkbox que a construtora marca quando concluir a manutenção");
  console.log("");
  console.log("🔒 PROTEÇÃO CONTRA ZERO À ESQUERDA (CPF, CONTRATO, TELEFONE):");
  console.log("  O Sheets converte respostas só-números em 'Número', derrubando zeros à");
  console.log("  esquerda (ex: '041...' vira '41...'). As colunas C, D e F agora são");
  console.log("  fixadas como Texto Simples em toda aba nova. Rode 'Corrigir formato");
  console.log("  CPF/Contrato/Telefone' no menu pra proteger abas já existentes também");
  console.log("  (isso NÃO recupera zeros já perdidos, só evita perder mais).");
  console.log("");
  console.log("🚀 PRIMEIRA VEZ:");
  console.log("  0. Configurações do projeto → Propriedades do script: cadastre");
  console.log("     ID_PASTA_RECLAMACOES e EMAIL_EQUIPE_MANUTENCAO (veja o README)");
  console.log("  1. Menu → ⚡ Instalar trigger automático (autorize as permissões)");
  console.log("  2. Menu → 🗂️ Organizar respostas por mês (move o histórico que já existe)");
  console.log("  3. Menu → 🧪 Testar Sistema");
  console.log("  4. Menu → ⏩ Processar TODAS Pendentes");
  console.log("");
  console.log(separador);
}

// ========================================
// 12. FUNÇÃO LEGADA (mantida, mas no fim do arquivo
//     para aparecer por último no dropdown do editor)
// ========================================

/**
 * Processa a última reclamação da aba do mês atual.
 */
function processarUltimaReclamacao() {
  const aba = obterAbaMesAtual();
  const ultimaLinha = obterUltimaLinha(aba);
  return _processarLinha(aba, ultimaLinha);
}
