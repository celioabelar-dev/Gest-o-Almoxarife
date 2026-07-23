/** * SISTEMA INTEGRADO DE FERRAMENTARIA V1.6
 * REGRAS: ID AUTOMÁTICO WF-000, EXCLUSÃO PROTEGIDA, BAIXA PARCIAL
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Ferramentaria Pro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// INSTALAÇÃO AUTOMÁTICA - RODE ESTA FUNÇÃO UMA VEZ
function instalarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = [
    { nome: 'DB_FERRAMENTAS', cabecalho: ['ID', 'DESCRICAO', 'TIPO', 'SERIE_TAG', 'NCM', 'LOCALIZACAO', 'TOTAL', 'DISPONIVEL', 'EM_MALETAS', 'EM_SPOT'] },
    { nome: 'DB_TECNICOS', cabecalho: ['ID', 'NOME', 'MATRICULA', 'BANCADA', 'STATUS'] },
    { nome: 'DB_COMPOSICAO_MALETAS', cabecalho: ['ID_MALETA', 'ID_FERRAMENTA', 'QUANTIDADE'] },
    { nome: 'DB_MOVIMENTACAO', cabecalho: ['ID_MOV', 'ID_TRANSACAO_GRUPO', 'DATA_SAIDA', 'ID_TECNICO', 'ID_FERRAMENTA', 'QTD_SAIDA', 'ORIGEM', 'STATUS', 'DATA_RETORNO', 'CONDICAO'] },
    { nome: 'DB_ENTRADAS_NF', cabecalho: ['DATA_ENTRADA', 'NF', 'FORNECEDOR', 'ID_FERRAMENTA', 'QTD', 'VALOR_UNIT', 'IPI'] }
  ];
  abas.forEach(aba => {
    let s = ss.getSheetByName(aba.nome);
    if (!s) {
      s = ss.insertSheet(aba.nome);
      s.getRange(1, 1, 1, aba.cabecalho.length).setValues([aba.cabecalho]).setBackground('#333').setFontColor('#fff').setFontWeight('bold');
    }
  });
  return "Estrutura de abas verificada!";
}

// --- MÓDULO GESTÃO (SALVAR COM ID WF-000) ---
function salvarFerramenta(obj) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_FERRAMENTAS');
  const d = sh.getDataRange().getValues();
  
  // 1. Normalizar a descrição para comparação (Tudo maiúsculo e sem espaços extras)
  const novaDescricao = obj.descricao.trim().toUpperCase();

  if(obj.id && obj.id !== "") {
    // Lógica de EDIÇÃO (Permite salvar se for o próprio item)
    for(let i=1; i<d.length; i++) {
      if(d[i][0] == obj.id) {
        sh.getRange(i+1, 2, 1, 5).setValues([[novaDescricao, obj.tipo, obj.serie, obj.ncm, obj.localizacao.toUpperCase()]]);
        return { sucesso: true, mensagem: "Item " + obj.id + " atualizado!" };
      }
    }
  } else {
    // 2. TRAVA DE DUPLICIDADE para Novo Cadastro
    const duplicado = d.some(linha => linha[1].toString().trim().toUpperCase() === novaDescricao);
    
    if (duplicado) {
      return { sucesso: false, mensagem: "Erro: Já existe uma ferramenta cadastrada com esta descrição!" };
    }

    // 3. ENCONTRAR A PRÓXIMA LINHA DISPONÍVEL APENAS PELA COLUNA A
    const colA = sh.getRange("A:A").getValues();
    let ultimaLinhaReal = 1;
    let maiorNum = 0;

    // Varre a coluna A para achar o último ID e calcular o sequencial correto
    for(let i=1; i<colA.length; i++){
      let idStr = colA[i][0].toString().trim();
      if(idStr !== "") {
        ultimaLinhaReal = i + 1; // Guarda a posição física da última célula preenchida
      }
      if(idStr.includes("WF-")){
        let num = parseInt(idStr.split("-")[1]);
        if(num > maiorNum) maiorNum = num;
      }
    }

    const proximaLinha = ultimaLinhaReal + 1;
    const novoId = "WF-" + (maiorNum + 1).toString().padStart(3, '0');
    
    // 4. SALVA EXATAMENTE NA PRÓXIMA LINHA DA COLUNA A ATÉ F, E INICIA O TOTAL (COLUNA G) COMO 0
    sh.getRange(proximaLinha, 1, 1, 7).setValues([[
      novoId, 
      novaDescricao, 
      obj.tipo, 
      obj.serie, 
      obj.ncm, 
      obj.localizacao.toUpperCase(), 
      0 // Coluna G: Inicializa o estoque total em 0
    ]]);

    // Deixamos de enviar dados para as colunas H, I, J, K para que a fórmula MAP faça o trabalho dela livremente!

    return { sucesso: true, mensagem: "Cadastrado como " + novoId };
  }
}

// --- EXCLUSÃO COM TRAVAS ---
function excluirFerramentaBanco(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shFer = ss.getSheetByName("DB_FERRAMENTAS");
  const shMov = ss.getSheetByName("DB_MOVIMENTACAO");
  
  const dFer = shFer.getDataRange().getValues();
  let linha = -1;
  let temEstoque = false;

  for (let i = 1; i < dFer.length; i++) {
    if (dFer[i][0] == id) {
      linha = i + 1;
      // Soma Colunas H (Disp) + I (Maleta) + J (Spot)
      const saldo = Number(dFer[i][7]) + Number(dFer[i][8]) + Number(dFer[i][9]);
      if (saldo > 0) temEstoque = true;
      break;
    }
  }

  if (temEstoque) return { sucesso: false, mensagem: "Erro: Item possui saldo em estoque." };

  const temMov = shMov.getDataRange().getValues().some(r => r[4] == id);
  if (temMov) return { sucesso: false, mensagem: "Erro: Item possui histórico de movimentação." };

  if (linha !== -1) {
    shFer.deleteRow(linha);
    return { sucesso: true, mensagem: "Excluído com sucesso!" };
  }
  return { sucesso: false, mensagem: "ID não encontrado." };
}

// --- BUSCAS ---
function buscarFerramentas(termo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_FERRAMENTAS');
  const d = sh.getDataRange().getValues();
  const res = [];
  const busca = termo ? termo.toLowerCase() : "";

  for(let i=1; i<d.length; i++) {
    if(termo === 'todos' || d[i][1].toString().toLowerCase().includes(busca) || d[i][0].toString().toLowerCase().includes(busca)) {
      res.push({
        id: d[i][0], 
        descricao: d[i][1], 
        tipo: d[i][2], 
        serie: d[i][3], 
        ncm: d[i][4], 
        localizacao: d[i][5], 
        total: d[i][6] || 0, 
        disponivel: d[i][7] || 0, 
        maleta: d[i][8] || 0, 
        spot: d[i][9] || 0
      });
    }
  }
  return {lista: res};
}

// --- MOVIMENTAÇÃO E SALDOS ---
function atualizarSaldo(id, qtd, acao, origem) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_FERRAMENTAS');
  const d = sh.getDataRange().getValues();
  
  for(let i=1; i<d.length; i++) {
    if(d[i][0] == id) {
      let disp = Number(d[i][7]);
      let maleta = Number(d[i][8]);
      let spot = Number(d[i][9]);
      
      if(acao === 'SAIDA') {
        sh.getRange(i+1, 8).setValue(disp - qtd);
        if(origem.includes('SPOT')) {
          sh.getRange(i+1, 10).setValue(spot + Number(qtd));
        } else {
          sh.getRange(i+1, 9).setValue(maleta + Number(qtd));
        }
      } else { // ENTRADA / BAIXA
        sh.getRange(i+1, 8).setValue(disp + Number(qtd));
        if(origem.includes('SPOT')) {
          sh.getRange(i+1, 10).setValue(spot - qtd);
        } else {
          sh.getRange(i+1, 9).setValue(maleta - qtd);
        }
      }
      break;
    }
  }
}


// Manter funções de técnicos e entradas conforme versão original...
function buscarTecnicos(termo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_TECNICOS');
  const d = sh.getDataRange().getValues();
  const resultados = [];
  
  // Ignora o cabeçalho (i=1) e varre as linhas da planilha
  for(let i = 1; i < d.length; i++) {
    const id = d[i][0];
    const nome = d[i][1];
    const matricula = d[i][2];
    const bancada = d[i][3];
    const status = d[i][4] || "ATIVO"; // <--- IMPORTANTE: Lê a 5ª coluna (E) da planilha. Se estiver vazia, assume "ATIVO".
    
    // Filtro básico (ajuste conforme a sua lógica original de busca)
    if(termo === "todos" || (nome && nome.toLowerCase().includes(termo.toLowerCase()))) {
      resultados.push({
        id: id,
        nome: nome,
        matricula: matricula,
        bancada: bancada,
        status: status // <--- IMPORTANTE: Envia o status para o Front-End
      });
    }
  }
  
  return resultados;
}

function salvarTecnico(obj) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_TECNICOS');
  const d = sh.getDataRange().getValues();
  
  if(obj.id) {
    // Edição de Técnico Existente
    for(let i = 1; i < d.length; i++) { 
      if(d[i][0] == obj.id) { 
        // Alterado de (i+1, 2, 1, 3) para (i+1, 2, 1, 4) para incluir a coluna do STATUS
        sh.getRange(i+1, 2, 1, 4).setValues([[
          obj.nome, 
          obj.matricula, 
          obj.bancada, 
          obj.status || "ATIVO" // Salva o status enviado (ATIVO/INATIVO)
        ]]); 
        return "Atualizado!"; 
      } 
    }
  } else {
    // Cadastro de Novo Técnico
    sh.appendRow(["TEC-" + d.length, obj.nome, obj.matricula, obj.bancada, "ATIVO"]);
    return "Cadastrado!";
  }
}

// ============================================================
// MÓDULO DE MALETAS / KITS - FUNÇÕES DO SERVIDOR
// ============================================================

/**
 * Salva uma nova maleta ou atualiza uma existente.
 * Gera ID no formato: MF-001 NOME DA MALETA
 */
function salvarNovaMaleta(id, descManual, itens) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('DB_COMPOSICAO_MALETAS');
  
  if (!sh) return "Erro: Aba DB_COMPOSICAO_MALETAS não encontrada!";
  
  // Força a descrição manual a ficar em maiúsculo e limpa
  descManual = descManual ? descManual.trim().toUpperCase() : "";
  if (descManual.length < 3) return "Erro: Informe uma descrição válida (mín. 3 caracteres).";

  // Carrega os dados atuais para a memória
  let dados = sh.getDataRange().getValues();
  let idApenas = "";

  // =========================================================
  // 1. LÓGICA DE DEFINIÇÃO DE ID (NOVO OU EDIÇÃO)
  // =========================================================
  if (!id || id === "" || id === "MF-000") {
    // CADASTRO NOVO: Descobre o maior sequencial puro (ex: MF-001)
    let maior = 0;
    for (let i = 1; i < dados.length; i++) {
      let idLinha = dados[i][0] ? dados[i][0].toString().trim().toUpperCase() : "";
      if (idLinha.startsWith("MF-")) {
        // Pega apenas a parte numérica antes de qualquer espaço ou hífen secundário
        let numStr = idLinha.split(" ")[0].split("-")[1];
        let num = parseInt(numStr);
        if (!isNaN(num) && num > maior) maior = num;
      }
    }
    const novoNum = (maior + 1).toString().padStart(3, '0');
    idApenas = "MF-" + novoNum;
  } else {
    // EDIÇÃO: Garante que estamos pegando apenas o ID limpo (ex: "MF-001")
    idApenas = id.split(" ")[0].trim().toUpperCase();
    
    // Remove os registros antigos correspondentes de trás para frente de forma segura
    for (let i = dados.length - 1; i >= 1; i--) {
      let idNoBanco = dados[i][0] ? dados[i][0].toString().trim().toUpperCase() : "";
      if (idNoBanco === idApenas) {
        sh.deleteRow(i + 1);
      }
    }
  }

  // =========================================================
  // 2. GRAVAÇÃO EM MASSA (PERFORMANCE OTIMIZADA)
  // =========================================================
  if (itens && itens.length > 0) {
    let novasLinhas = [];
    
    itens.forEach(it => {
      let idFerramenta = it.id ? it.id.toString().trim().toUpperCase() : "";
      
      // Padronização opcional: Se no banco de composição você quer salvar como FW- ou WF-,
      // você pode tratar aqui. Exemplo mantendo o que veio do front:
      if (idFerramenta !== "") {
        novasLinhas.push([
          idApenas,                 // Coluna A: ID Puro (MF-001)
          idFerramenta,             // Coluna B: ID da Ferramenta
          parseInt(it.qtd) || 1     // Coluna C: Quantidade
        ]);
      }
    });

    // Envia todas as linhas de uma vez só para a planilha (Evita timeouts e lentidão)
    if (novasLinhas.length > 0) {
      const ultimaLinha = sh.getLastRow();
      sh.getRange(ultimaLinha + 1, 1, novasLinhas.length, 3).setValues(novasLinhas);
    }
  }
  
  // Retorna uma mensagem amigável combinando o ID gerado/editado com a descrição
  return "✅ Kit " + idApenas + " (" + descManual + ") processado com sucesso!";
}

/**
 * Retorna uma lista resumida de todas as maletas para a tabela principal
 */
function buscarResumoMaletas() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_COMPOSICAO_MALETAS');
  if (!sh) return [];
  
  const d = sh.getDataRange().getValues();
  const resumo = {};

  // Agrupa por ID da maleta para contar quantos itens diferentes existem em cada uma
  for(let i=1; i<d.length; i++) {
    const idMaleta = d[i][0];
    if(!idMaleta) continue;
    if(!resumo[idMaleta]) resumo[idMaleta] = 0;
    resumo[idMaleta]++;
  }

  // Converte o objeto em array para o JavaScript do HTML
  return Object.keys(resumo).map(key => ({ 
    id: key, 
    totalItens: resumo[key] 
  }));
}

/**
 * Busca os itens de uma maleta específica para carregar no formulário de edição
 */
// Busca todas as maletas únicas na aba de composição
function buscarTodasMaletas() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_COMPOSICAO_MALETAS');
  if (!sh) return [];
  const dados = sh.getDataRange().getValues();
  const mapa = new Set();
  const lista = [];

  for (let i = 1; i < dados.length; i++) {
    let id = dados[i][0];
    if (id && !mapa.has(id)) {
      lista.push({ id: id });
      mapa.add(id);
    }
  }
  return lista;
}

// Gera o HTML do Termo baseado no modelo enviado
function gerarTermoResponsabilidade(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tecnico = ss.getSheetByName('DB_TECNICOS').getDataRange().getValues().find(t => t[0] == dados.tecnicoId);
  const composicao = ss.getSheetByName('DB_COMPOSICAO_MALETAS').getDataRange().getValues().filter(r => r[0] === dados.idMaleta);
  const dFer = ss.getSheetByName('DB_FERRAMENTAS').getDataRange().getValues();
  
  let itensHtml = "";
  composicao.forEach(it => {
    const fer = dFer.find(f => f[0] == it[1]);
    itensHtml += `<tr>
      <td style="border:1px solid #000;padding:5px;">${it[1]}</td>
      <td style="border:1px solid #000;padding:5px;">${fer ? fer[1] : 'Item não cadastrado'}</td>
      <td style="border:1px solid #000;padding:5px;text-align:center;">${it[2]}</td>
    </tr>`;
  });

  const dataAtual = Utilities.formatDate(new Date(), "GMT-3", "dd 'de' MMMM 'de' yyyy");

  return `
    <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 30px; font-size: 11.5px; line-height: 1.4; color: #000; }
        .cabecalho { text-align: center; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th { background: #e9ecef; border: 1px solid #000; padding: 6px; font-weight: bold; }
        td { border: 1px solid #000; padding: 4px 8px; }
        .assinaturas { margin-top: 50px; display: flex; justify-content: space-between; }
        .campo { width: 45%; border-top: 1px solid #000; text-align: center; padding-top: 5px; }
        .no-print { background: #28a745; color: white; padding: 10px; text-align: center; margin-bottom: 20px; border-radius: 5px; cursor: pointer; border: none; width: 100%; font-weight: bold; }
        @media print { .no-print { display: none; } body { padding: 0; } }
      </style>
    </head>
    <body contenteditable="true">
      <button class="no-print" onclick="window.print()">CLIQUE AQUI PARA IMPRIMIR O TERMO</button>

      <div class="cabecalho">
        <h3 style="margin:0;">TERMO DE RESPONSABILIDADE E ENTREGA DE EQUIPAMENTOS/FERRAMENTAS</h3>
      </div>

      <p><strong>EMPREGADOR:</strong> CRAS FERRAMENTARIA, inscrita no CNPJ sob o nº 14.442.765/0001-96.<br>
      <strong>EMPREGADO:</strong> ${tecnico ? tecnico[1] : '[NOME DO FUNCIONÁRIO]'}, portador do CPF nº ${tecnico ? (tecnico[4] || '[CPF]') : '[CPF]'}.</p>

      <p><strong>1. DA ENTREGA DOS EQUIPAMENTOS</strong><br>
      O Empregado declara ter recebido da Empresa, a título de empréstimo para o desempenho de suas funções profissionais, os equipamentos/ferramentas da <strong>${dados.idMaleta}</strong> descritos na tabela abaixo, em perfeitas condições de uso e conservação:</p>

      <table>
        <thead>
          <tr>
            <th>Código</th>
            <th>Descrição</th>
            <th>Quantidade</th>
          </tr>
        </thead>
        <tbody>
          ${itensHtml}
        </tbody>
      </table>

      <p><strong>2. DO COMPROMISSO E USO</strong><br>
      O Empregado compromete-se a:<br>
      • Utilizar o equipamento exclusivamente para fins profissionais vinculados à Empresa.<br>
      • Zelar pela guarda, limpeza e conservação dos itens, evitando exposição a riscos desnecessários.<br>
      • Não ceder, emprestar ou alugar os equipamentos a terceiros sem autorização prévia por escrito.<br>
      • Comunicar imediatamente à gerência qualquer defeito, mau funcionamento, furto ou roubo.</p>

      <p><strong>3. DA DEVOLUÇÃO</strong><br>
      Ao término do contrato de trabalho, ou mediante solicitação da Empresa, o Empregado deverá devolver os equipamentos no estado em que os recebeu, salvo o desgaste natural pelo uso regular.</p>

      <p><strong>4. DA AUTORIZAÇÃO DE DESCONTO (Art. 462, §1º da CLT)</strong><br>
      Fica acordado que a Empresa poderá efetuar o desconto do valor correspondente ao equipamento ou ao seu conserto, diretamente na folha de pagamento, férias ou verbas rescisórias do Empregado, nas seguintes situações:<br>
      • <strong>Dano por Culpa ou Dolo:</strong> Em caso de negligência, imprudência, imperícia ou mau uso comprovado.<br>
      • <strong>Extravio ou Não Devolução:</strong> Em caso de perda do equipamento ou falta de entrega no encerramento do vínculo empregatício.<br>
      • <strong>Avaria por Falta de Cuidado:</strong> Danos físicos decorrentes de armazenamento inadequado ou queda por descuido.</p>
      
      <p><small>Observação: O valor do desconto será baseado no preço de mercado atualizado do equipamento ou no custo real do orçamento de manutenção.</small></p>

      <p style="margin-top: 20px;">Local e Data: Contagem, ${dataAtual}.</p>

      <div class="assinaturas">
        <div class="campo">CRAS (Empregador)</div>
        <div class="campo">${tecnico ? tecnico[1] : 'Empregado'}</div>
      </div>
    </body>
    </html>
  `;
}

function registrarNotaFiscalLote(nfDados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shEntradas = ss.getSheetByName('DB_ENTRADAS_NF');
  const shFerramentas = ss.getSheetByName('DB_FERRAMENTAS');
  const dataAtual = new Date();
  
  const dbFer = shFerramentas.getDataRange().getValues();

  try {
    nfDados.itens.forEach(item => {
      // 1. Grava na aba DB_ENTRADAS_NF
      // A:Data | B:NF | C:Fornecedor | D:Cod | E:QTD | F:Unit | G:IPI | H:NCM
      shEntradas.appendRow([
        dataAtual,          // A
        nfDados.nf,         // B
        nfDados.fornecedor, // C
        item.idFerramenta,  // D
        item.qtd,           // E
        item.valor,         // F
        item.ipi,           // G
        item.ncm            // H - COLUNA SOLICITADA
      ]);

      // 2. Atualiza Saldo em DB_FERRAMENTAS (G:Total e H:Disponível)
      for (let i = 1; i < dbFer.length; i++) {
        if (dbFer[i][0] == item.idFerramenta) {
          let totalAtual = Number(dbFer[i][6]) || 0;
          let dispoAtual = Number(dbFer[i][7]) || 0;
          
          shFerramentas.getRange(i + 1, 7).setValue(totalAtual + item.qtd);
          shFerramentas.getRange(i + 1, 8).setValue(dispoAtual + item.qtd);
          
          // Se o NCM no cadastro (Coluna E - índice 4) estiver vazio, preenche
          if (!dbFer[i][4]) {
             shFerramentas.getRange(i + 1, 5).setValue(item.ncm);
          }
          break;
        }
      }
    });

    return "✅ Nota Fiscal " + nfDados.nf + " processada! Itens adicionados e NCM gravado na coluna H.";
  } catch (e) {
    return "❌ Erro no lote: " + e.toString();
  }
}

function buscarMovimentacoesGestao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shMov = ss.getSheetByName('DB_MOVIMENTACAO'); // Certifique-se que o nome é este
  const shFer = ss.getSheetByName('DB_FERRAMENTAS');
  const shTec = ss.getSheetByName('DB_TECNICOS');
  
  if (!shMov) return [];

  const dadosMov = shMov.getDataRange().getValues();
  const dadosFer = shFer.getDataRange().getValues();
  const dadosTec = shTec.getDataRange().getValues();

  let resultado = [];

  // Começa do 1 para pular o cabeçalho
  for (let i = 1; i < dadosMov.length; i++) {
    // Só traz as que NÃO possuem data de devolução (Estoque em campo)
    // Supondo: Coluna A:DataSaida, B:ID_Tec, C:ID_Fer, D:Qtd, E:Tipo, F:DataDevolucao
    if (dadosMov[i][5] === "" || !dadosMov[i][5]) {
      
      // Busca nome do técnico
      const tecnico = dadosTec.find(t => t[0] == dadosMov[i][1]);
      // Busca descrição da ferramenta
      const ferramenta = dadosFer.find(f => f[0] == dadosMov[i][2]);

      resultado.push({
        idMov: i + 1, // Linha na planilha para facilitar a baixa
        dataSaida: Utilities.formatDate(new Date(dadosMov[i][0]), "GMT-3", "dd/MM/yyyy"),
        tecnico: tecnico ? tecnico[1] : "Não encontrado",
        item: ferramenta ? ferramenta[1] : dadosMov[i][2],
        qtd: dadosMov[i][3],
        modalidade: dadosMov[i][4]
      });
    }
  }
  return resultado.reverse(); // Mostrar as mais recentes primeiro
}

function registrarSaida(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shMov = ss.getSheetByName('DB_MOVIMENTACAO'); // Certifique-se de que o nome está correto (com 'N')
  const shComp = ss.getSheetByName('DB_COMPOSICAO_MALETAS');
  const dataHoje = new Date();
  const idTrans = "TR-" + dataHoje.getTime();

  if (!shMov) return "❌ Erro: Aba DB_MOVIMENTACAO não encontrada!";

  const tipoSaida = dados.tipoSaida ? dados.tipoSaida.toString().toUpperCase().trim() : "";

  // =========================================================
  // CENÁRIO 1: SAÍDA DE MALETA (Gera 1 linha para CADA item do Kit)
  // =========================================================
  if (tipoSaida === 'MALETA') {
    if (!shComp) return "❌ Erro: Aba DB_COMPOSICAO_MALETAS não encontrada!";

    let idMaletaPuro = dados.idMaleta ? dados.idMaleta.toString().split(" ")[0].trim().toUpperCase() : "";
    if (idMaletaPuro === "") return "❌ Erro: Selecione uma maleta válida!";

    // Busca todos os itens da composição da maleta
    const dadosComposicao = shComp.getDataRange().getValues();
    const comp = dadosComposicao.filter(r => r[0] ? r[0].toString().trim().toUpperCase() === idMaletaPuro : false);
    
    if (comp.length === 0) return "❌ Erro: Nenhuma ferramenta encontrada na composição de " + idMaletaPuro;
    
    let novasLinhasMov = [];
    
    // Varre os itens da maleta e prepara as linhas do histórico
    comp.forEach(it => {
      let idMov = "MOV-" + Math.floor(Math.random() * 900000 + 100000);
      novasLinhasMov.push([
        idMov,           // A: ID_MOV
        idTrans,         // B: ID_TRANSACAO
        dataHoje,        // C: DATA_SAIDA
        dados.tecnicoId, // D: ID_TECNICO
        it[1],           // E: ID_FERRAMENTA (Salva o código do item de dentro do kit, ex: FW-001)
        it[2],           // F: QTD_SAIDA
        idMaletaPuro,    // G: ORIGEM (Registra que veio da MF-001)
        "ABERTO",        // H: STATUS
        "",              // I: DATA_RETORNO
        "MALETA"         // J: CONDICAO/MODALIDADE
      ]);
    });
    
    // Gravação em lote super rápida
    if (novasLinhasMov.length > 0) {
      const ultimaLinha = shMov.getLastRow();
      shMov.getRange(ultimaLinha + 1, 1, novasLinhasMov.length, 10).setValues(novasLinhasMov);
      return "✅ Saída da Maleta " + idMaletaPuro + " (" + novasLinhasMov.length + " itens) registrada!";
    }
    
  } 
  // =========================================================
  // CENÁRIO 2: SAÍDA AVULSA (Gera apenas uma única linha comum)
  // =========================================================
  else {
    let ferramentaRaw = dados.ferramentaId || dados.ferramenta || "";
    let idFerramentaPuro = ferramentaRaw.toString().split(" ")[0].trim().toUpperCase();
    
    if (idFerramentaPuro === "") return "❌ Erro: Nenhuma ferramenta selecionada!";

    let idMov = "MOV-" + Math.floor(Math.random() * 900000 + 100000);
    let quantidade = parseInt(dados.qtd) || 1;
    let modalidade = dados.tipoEmprestimo || "SPOT";

    shMov.appendRow([
      idMov,             // A: ID_MOV
      idTrans,           // B: ID_TRANSACAO
      dataHoje,          // C: DATA_SAIDA
      dados.tecnicoId,   // D: ID_TECNICO
      idFerramentaPuro,  // E: ID_FERRAMENTA
      quantidade,        // F: QTD_SAIDA
      "AVULSO",          // G: ORIGEM
      "ABERTO",          // H: STATUS
      "",                // I: DATA_RETORNO
      modalidade         // J: CONDICAO/MODALIDADE (SPOT ou MOBILIZADO)
    ]);
    
    return "✅ Saída avulsa do item " + idFerramentaPuro + " registrada!";
  }
}


function buscarMovimentacoesAbertas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // CORREÇÃO CRÍTICA: Ajustado para o nome correto com a letra "N" -> DB_MOVIMENTACAO
  const shMov = ss.getSheetByName('DB_MOVIMENTACAO'); 
  const shTec = ss.getSheetByName('DB_TECNICOS');
  
  if (!shMov) {
    console.error("ERRO: Aba 'DB_MOVIMENTACAO' não foi encontrada. Verifique o nome da aba na sua planilha!");
    return [];
  }
  
  const dadosMov = shMov.getDataRange().getValues();
  const dadosTec = shTec.getDataRange().getValues();
  
  let lista = [];
  let maletasAgrupadas = {}; 

  // Mapeia Técnicos: ID (A) -> Nome (B)
  let tecMap = {};
  for(let j = 1; j < dadosTec.length; j++) {
    if(dadosTec[j][0]) {
      tecMap[dadosTec[j][0].toString().trim()] = dadosTec[j][1];
    }
  }

  // Percorre as linhas de movimentação (pula o cabeçalho)
  for (let i = 1; i < dadosMov.length; i++) {
    if (!dadosMov[i][0] && !dadosMov[i][1]) continue; // Pula linhas fantasma vazias

    // Pega o Status da coluna H (Índice 7)
    let statusEmprestimo = dadosMov[i][7] ? dadosMov[i][7].toString().trim().toUpperCase() : "";
    
    // FILTRO: Se estiver aberto, processa
    if (statusEmprestimo === "ABERTO") {
      
      let idTecnico     = dadosMov[i][3] ? dadosMov[i][3].toString().trim() : "";
      let nomeTecnico   = tecMap[idTecnico] || idTecnico || "Não Identificado";
      let codigoMaleta  = dadosMov[i][6] ? dadosMov[i][6].toString().trim() : "";
      let modalidade    = dadosMov[i][9] ? dadosMov[i][9].toString().trim().toUpperCase() : ""; // Padronizado para Maiúsculo
      let ferAvulsa     = dadosMov[i][4] ? dadosMov[i][4].toString().trim() : "";
      let quantidade    = dadosMov[i][5] !== "" ? Number(dadosMov[i][5]) : 1;

      // Formatação da Data de Saída (Coluna C -> Índice 2)
      let dataFormatada = "---";
      if (dadosMov[i][2]) {
        try {
          dataFormatada = Utilities.formatDate(new Date(dadosMov[i][2]), "GMT-3", "dd/MM/yyyy HH:mm");
        } catch(e) {
          dataFormatada = dadosMov[i][2].toString();
        }
      }

      // REGRA DE UNIFICAÇÃO DA MALETA
      if (modalidade === "MALETA") {
        if (codigoMaleta !== "") {
          let chaveMaleta = idTecnico + "_" + codigoMaleta;
          
          if (!maletasAgrupadas[chaveMaleta]) {
            maletasAgrupadas[chaveMaleta] = true; 
            
            lista.push({
              idMov: i + 1, 
              idTransacao: dadosMov[i][1] ? dadosMov[i][1].toString() : "",
              dataSaida: dataFormatada,
              tec: nomeTecnico,
              fer: codigoMaleta, // Exibe o código do kit (ex: MF-001)
              qtd: 1,            // Agrupa em 1 volume fixo
              origem: codigoMaleta,
              status: "MALETA"
            });
          }
        }
      } else {
        // REGRA PARA LOGÍSTICA AVULSA (SPOT / MOBILIZADO)
        lista.push({
          idMov: i + 1,
          idTransacao: dadosMov[i][1] ? dadosMov[i][1].toString() : "",
          dataSaida: dataFormatada,
          tec: nomeTecnico,
          fer: ferAvulsa, 
          qtd: quantidade,
          origem: codigoMaleta,
          status: modalidade !== "" ? modalidade : "SPOT"
        });
      }
    }
  }
  
  return lista.reverse(); 
}


function buscarItensDaMaleta(idMaleta) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaComposicao = planilha.getSheetByName("DB_COMPOSICAO_MALETAS");
  var abaFerramentas = planilha.getSheetByName("DB_FERRAMENTAS");
  
  if (!abaComposicao || !abaFerramentas) {
    throw new Error("Abas não encontradas.");
  }
  
  var dadosComposicao = abaComposicao.getDataRange().getValues();
  var dadosFerramentas = abaFerramentas.getDataRange().getValues();
  
  // Cria o mapa de ferramentas tratando espaços em branco e letras maiúsculas
  var mapaFerramentas = {};
  for (var f = 1; f < dadosFerramentas.length; f++) {
    var idFerramenta = dadosFerramentas[f][0];        // Coluna A
    var descricaoFerramenta = dadosFerramentas[f][1]; // Coluna B
    if (idFerramenta) {
      var chave = idFerramenta.toString().trim().toUpperCase();
      mapaFerramentas[chave] = descricaoFerramenta;
    }
  }
  
  var itensEncontrados = [];
  // Força o ID buscado a ignorar espaços antes/depois
  var idMaletaBusca = idMaleta.toString().trim();
  
  for (var i = 1; i < dadosComposicao.length; i++) {
    var row = dadosComposicao[i];
    var maletaIdNoBanco = row[0] ? row[0].toString().trim() : "";
    
    if (maletaIdNoBanco == idMaletaBusca) {
      var idFerr = row[1] ? row[1].toString().trim() : "";
      var qtdFerr = row[2];
      
      // Busca a descrição usando a chave tratada
      var chaveFerr = idFerr.toUpperCase();
      var descFerr = mapaFerramentas[chaveFerr] || "Descrição não encontrada"; 
      
      itensEncontrados.push({
        id: idFerr,
        desc: descFerr, // Aqui vai a descrição real da coluna B
        qtd: parseInt(qtdFerr) || 1
      });
    }
  }
  
  return itensEncontrados;
}

function confirmarBaixa(idLinha, motivo, observacao) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shMov = ss.getSheetByName('DB_MOVIMENTACAO');
    
    if (!shMov) throw new Error("Aba DB_MOVIMENTACAO não encontrada.");
    
    let linha = Number(idLinha);
    if (isNaN(linha) || linha <= 1) throw new Error("Identificador de linha inválido.");
    
    const dadosMov = shMov.getDataRange().getValues();
    
    // 1. Identifica os dados da linha clicada (Ponto de partida)
    let registroClicado = dadosMov[linha - 1]; 
    let modalidade = registroClicado[9] ? registroClicado[9].toString().toUpperCase().trim() : "";
    let idTecnicoClicado = registroClicado[3] ? registroClicado[3].toString().trim() : ""; 
    let codigoMaletaClicado = registroClicado[6] ? registroClicado[6].toString().trim() : ""; 
    
    let linhasParaBaixar = [];

    // 2. Lógica de Agrupamento: Se for Maleta, busca todos os itens do mesmo Técnico + Maleta
    if (modalidade === "MALETA") {
      for (let i = 1; i < dadosMov.length; i++) {
        let status = dadosMov[i][7] ? dadosMov[i][7].toString().trim().toUpperCase() : "";
        let idTecnicoLinha = dadosMov[i][3] ? dadosMov[i][3].toString().trim() : "";
        let codigoMaletaLinha = dadosMov[i][6] ? dadosMov[i][6].toString().trim() : "";

        // Verifica se está aberto E se pertence ao MESMO técnico E à MESMA maleta
        if (status === "ABERTO" && 
            idTecnicoLinha === idTecnicoClicado && 
            codigoMaletaLinha === codigoMaletaClicado) {
          linhasParaBaixar.push(i + 1); // +1 pois o getRange começa na linha 1
        }
      }
    } else {
      // Para itens avulsos (SPOT/MOBILIZADO), baixa apenas a linha específica
      linhasParaBaixar.push(linha);
    }

    // 3. Executa a baixa em todas as linhas identificadas
    motivo = motivo ? motivo.toUpperCase().trim() : "NORMAL";
    observacao = observacao ? observacao.trim() : "";
    let dataAtual = new Date();

    linhasParaBaixar.forEach(n => {
      shMov.getRange(n, 8).setValue("DEVOLVIDO"); // Coluna H
      shMov.getRange(n, 9).setValue(dataAtual);   // Coluna I
      if (motivo !== "NORMAL") {
        shMov.getRange(n, 11).setValue(`BAIXA ESPECIAL: [${motivo}] - ${observacao}`); // Coluna K
      }
    });

    // 4. Retorno para o front-end
    return {
      sucesso: true,
      msg: `✅ ${linhasParaBaixar.length} item(ns) da maleta baixado(s) com sucesso!`,
      dadosRomaneio: {
        idItem: modalidade === "MALETA" ? codigoMaletaClicado : registroClicado[4],
        quantidade: linhasParaBaixar.length,
        tecnicoId: idTecnicoClicado,
        motivoBaixa: motivo,
        obsBaixa: observacao
      }
    };
    
  } catch (erro) {
    console.error("Erro em confirmarBaixa:", erro.toString());
    return { sucesso: false, msg: "❌ Erro interno ao realizar baixa: " + erro.message };
  }
}

function gerarRomaneioEmprestimo(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Busca os dados do Técnico
  const dTec = ss.getSheetByName('DB_TECNICOS').getDataRange().getValues();
  const tecnico = dTec.find(t => t[0] == dados.tecnicoId);
  const nomeTecnico = tecnico ? tecnico[1] : 'Não Identificado';
  
  // 2. Busca os dados da Ferramenta e o Endereço (Coluna F / Índice 5)
  const dFer = ss.getSheetByName('DB_FERRAMENTAS').getDataRange().getValues();
  const ferramenta = dFer.find(f => f[0] == dados.idItem);
  const nomeFerramenta = ferramenta ? ferramenta[1] : (dados.idItem || 'Item Avulso');
  const enderecoFerramenta = (ferramenta && ferramenta[5]) ? ferramenta[5].toString() : 'Não Informado';

  const dataAtual = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
  const prazo = (dados.modalidade === "SPOT") ? "07 DIAS (Retorno)" : "MOBILIZADO";

  return `
    <html>
    <head>
      <style>
        @page { margin: 0; }
        body { 
          font-family: 'Courier New', Courier, monospace; 
          width: 280px; 
          margin: 0; 
          padding: 8px; 
          font-size: 11px; 
          line-height: 1.3; 
          color: #000; 
        }
        .text-center { text-align: center; }
        .text-right { text-align: right; }
        .bold { font-weight: bold; }
        .divisor { border-top: 1px dashed #000; margin: 6px 0; }
        .item-table { width: 100%; font-size: 11px; border-collapse: collapse; margin-top: 4px; }
        .item-table td { padding: 2px 0; vertical-align: top; }
        .no-print { 
          background: #007bff; color: white; padding: 8px; text-align: center; 
          margin-bottom: 15px; border-radius: 4px; cursor: pointer; border: none; 
          width: 100%; font-weight: bold; font-family: sans-serif; font-size: 12px;
        }
        @media print { .no-print { display: none; } }
      </style>
    </head>
    <body>
      <button class="no-print" onclick="window.print()">🖨️ IMPRIMIR ROMANEIO</button>

      <div class="text-center bold" style="font-size: 13px;">CRAS</div>
      <div class="text-center" style="font-size: 9px;">Ferramentaria</div>
      
      <div class="divisor"></div>
      <div class="text-center bold" style="font-size: 12px;">ROMANEIO DE EMPRÉSTIMO</div>
      <div class="divisor"></div>
      
      <div><span class="bold">DATA/HORA:</span> ${dataAtual}</div>
      <div><span class="bold">TÉCNICO:</span> ${nomeTecnico}</div>
      <div><span class="bold">MODALIDADE:</span> ${dados.modalidade}</div>
      <div><span class="bold">PRAZO RETORNO:</span> ${prazo}</div>
      
      <div class="divisor"></div>
      
      <table class="item-table">
        <thead>
          <tr class="bold">
            <td style="width: 75%;">ITEM / DESCRIÇÃO</td>
            <td class="text-right" style="width: 25%;">QTD</td>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>[${dados.idItem}] ${nomeFerramenta}</td>
            <td class="text-right bold">${dados.quantidade}</td>
          </tr>
          <tr>
            <td colspan="2" style="font-size: 10px; color: #333; padding-top: 4px;">
              📌 <span class="bold">ENDEREÇO FÍSICO:</span> ${enderecoFerramenta}
            </td>
          </tr>
        </tbody>
      </table>
      
      <div class="divisor"></div>
      <br><br><br>
      <div class="text-center">_________________________________</div>
      <div class="text-center bold" style="font-size: 9px; margin-top: 3px;">Assinatura do Colaborador (Recebi)</div>
      <br>
      <div class="text-center" style="font-size: 8px; opacity: 0.6;">Controle de Fluxo Interno</div>
    </body>
    </html>
  `;
}

function gerarRomaneioDevolucao(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Busca os dados do Técnico (Garante busca por ID ou por Nome Direto)
    const dTec = ss.getSheetByName('DB_TECNICOS').getDataRange().getValues();
    const termoTecnico = dados.tecnicoId ? dados.tecnicoId.toString().trim().toUpperCase() : "";
    
    const tecnico = dTec.find(t => {
      if (!t[0]) return false;
      // Compara tanto com o ID (coluna A) quanto com o Nome (coluna B) para evitar falhas
      return t[0].toString().trim().toUpperCase() === termoTecnico || 
             (t[1] && t[1].toString().trim().toUpperCase() === termoTecnico);
    });
    const nomeTecnico = tecnico ? tecnico[1] : (dados.tecnicoId || 'Não Identificado');
    
    // 2. Busca os dados da Ferramenta e o Endereço (Coluna F / Índice 5)
    const dFer = ss.getSheetByName('DB_FERRAMENTAS').getDataRange().getValues();
    const termoItem = dados.idItem ? dados.idItem.toString().trim().toUpperCase() : "";
    
    const ferramenta = dFer.find(f => f[0] && f[0].toString().trim().toUpperCase() === termoItem);
    const nomeFerramenta = ferramenta ? ferramenta[1] : (dados.idItem || 'Item Avulso');
    const enderecoFerramenta = (ferramenta && ferramenta[5]) ? ferramenta[5].toString() : 'Não Informado';

    const dataAtual = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
    
    // Define o status visual baseado na condição informada na baixa
    let statusFormatado = "🔄 RETORNADO AO ESTOQUE";
    if (dados.motivoBaixa && dados.motivoBaixa !== "NORMAL") {
      statusFormatado = `💥 BAIXA POR: ${dados.motivoBaixa.toUpperCase()}`;
    }

    return `
      <html>
      <head>
        <style>
          @page { margin: 0; }
          body { 
            font-family: 'Courier New', Courier, monospace; 
            width: 280px; 
            margin: 0; 
            padding: 8px; 
            font-size: 11px; 
            line-height: 1.3; 
            color: #000; 
          }
          .text-center { text-align: center; }
          .text-right { text-align: right; }
          .bold { font-weight: bold; }
          .divisor { border-top: 1px dashed #000; margin: 6px 0; }
          .item-table { width: 100%; font-size: 11px; border-collapse: collapse; margin-top: 4px; }
          .item-table td { padding: 2px 0; vertical-align: top; }
          .no-print { 
            background: #28a745; color: white; padding: 8px; text-align: center; 
            margin-bottom: 15px; border-radius: 4px; cursor: pointer; border: none; 
            width: 100%; font-weight: bold; font-family: sans-serif; font-size: 12px;
          }
          @media print { .no-print { display: none; } }
        </style>
      </head>
      <body>
        <button class="no-print" onclick="window.print()">🖨️ IMPRIMIR COMPROVANTE</button>

        <div class="text-center bold" style="font-size: 13px;">CRAS</div>
        <div class="text-center" style="font-size: 9px;">Ferramentaria</div>
        
        <div class="divisor"></div>
        <div class="text-center bold" style="font-size: 12px;">COMPROVANTE DE DEVOLUÇÃO</div>
        <div class="divisor"></div>
        
        <div><span class="bold">DATA/HORA:</span> ${dataAtual}</div>
        <div><span class="bold">TÉCNICO:</span> ${nomeTecnico}</div>
        <div><span class="bold">STATUS RETORNO:</span> ${statusFormatado}</div>
        ${dados.obsBaixa ? `<div><span class="bold">OBS:</span> ${dados.obsBaixa}</div>` : ""}
        
        <div class="divisor"></div>
        
        <table class="item-table">
          <thead>
            <tr class="bold">
              <td style="width: 75%;">ITEM / DESCRIÇÃO</td>
              <td class="text-right" style="width: 25%;">QTD</td>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>[${dados.idItem}] ${nomeFerramenta}</td>
              <td class="text-right bold">${dados.quantidade}</td>
            </tr>
            <tr>
              <td colspan="2" style="font-size: 10px; color: #333; padding-top: 4px;">
                📌 <span class="bold">GUARDAR EM:</span> ${enderecoFerramenta}
              </td>
            </tr>
          </tbody>
        </table>
        
        <div class="divisor"></div>
        <br><br><br>
        <div class="text-center">_________________________________</div>
        <div class="text-center bold" style="font-size: 9px; margin-top: 3px;">Visto do Operador (Ferramentaria)</div>
        <br>
        <div class="text-center" style="font-size: 8px; opacity: 0.6;">Controle de Fluxo Interno</div>
      </body>
      </html>
    `;
  } catch(e) {
    return `<html><body><h3>Erro no Servidor ao Gerar Layout: ${e.message}</h3></body></html>`;
  }
}

/**
 * CRUZA DADOS DE TODAS AS ABAS PARA GERAR INSIGHTS E MÉTRICAS EM TEMPO REAL
 */
function obterDadosRelatorios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shFer = ss.getSheetByName('DB_FERRAMENTAS');
  const shTec = ss.getSheetByName('DB_TECNICOS');
  const shMov = ss.getSheetByName('DB_MOVIMENTACAO');
  const shNf = ss.getSheetByName('DB_ENTRADAS_NF');

  const dFer = shFer ? shFer.getDataRange().getValues() : [];
  const dTec = shTec ? shTec.getDataRange().getValues() : [];
  const dMov = shMov ? shMov.getDataRange().getValues() : [];
  const dNf = shNf ? shNf.getDataRange().getValues() : [];

  // Mapeamento de Técnicos e Ferramentas para buscas ultra rápidas em memória
  const tecMap = {};
  for(let i = 1; i < dTec.length; i++) {
    if(dTec[i][0]) tecMap[dTec[i][0].toString().trim()] = dTec[i][1];
  }

  const ferMap = {};
  for(let i = 1; i < dFer.length; i++) {
    if(dFer[i][0]) {
      const idStr = dFer[i][0].toString().trim();
      ferMap[idStr] = {
        descricao: dFer[i][1],
        tipo: dFer[i][2],
        ncm: dFer[i][4],
        localizacao: dFer[i][5],
        total: Number(dFer[i][6]) || 0,
        disponivel: Number(dFer[i][7]) || 0,
        emMaletas: Number(dFer[i][8]) || 0,
        emSpot: Number(dFer[i][9]) || 0
      };
    }
  }

  // Preço Unitário Médio por Ferramenta (Cálculo de Valoração com base na DB_ENTRADAS_NF)
  const precoMap = {};
  for(let i = 1; i < dNf.length; i++) {
    const idFer = dNf[i][3] ? dNf[i][3].toString().trim() : "";
    const unit = Number(dNf[i][4]) || 0; // Quantidade comprada
    const valorUnit = Number(dNf[i][5]) || 0; // Valor Unitário
    const ipi = Number(dNf[i][6]) || 0; // Porcentagem do IPI
    
    const valorComIpi = valorUnit * (1 + (ipi / 100));
    
    if(idFer) {
      if(!precoMap[idFer]) {
        precoMap[idFer] = { totalGasto: 0, totalQtd: 0 };
      }
      precoMap[idFer].totalGasto += (valorComIpi * unit);
      precoMap[idFer].totalQtd += unit;
    }
  }

  // Calcula preço médio ponderado por ID
  const precoMedioMap = {};
  Object.keys(precoMap).forEach(key => {
    const p = precoMap[key];
    precoMedioMap[key] = p.totalQtd > 0 ? (p.totalGasto / p.totalQtd) : 0;
  });

  // --- 1. KPIs (Indicadores no topo da página) ---
  let totalItensEstoque = 0;
  let valorTotalEstoque = 0;
  let totalEmprestados = 0;
  let itensCriticos = 0; // Sem estoque disponível para novos empréstimos

  Object.keys(ferMap).forEach(id => {
    const f = ferMap[id];
    totalItensEstoque += f.total;
    const precoMedio = precoMedioMap[id] || 0;
    valorTotalEstoque += (f.total * precoMedio);
    totalEmprestados += (f.emMaletas + f.emSpot);
    if (f.total > 0 && f.disponivel === 0) {
      itensCriticos++;
    }
  });

  // --- 2. Relatório: Empréstimos Ativos (Em aberto) ---
  const emprestimosAtivos = [];
  const hoje = new Date();
  for(let i = 1; i < dMov.length; i++) {
    const status = dMov[i][7] ? dMov[i][7].toString().trim().toUpperCase() : "";
    if(status === "ABERTO") {
      const dataSaida = dMov[i][2] ? new Date(dMov[i][2]) : null;
      const idTec = dMov[i][3] ? dMov[i][3].toString().trim() : "";
      const idFer = dMov[i][4] ? dMov[i][4].toString().trim() : "";
      const qtd = Number(dMov[i][5]) || 0;
      const modalidade = dMov[i][9] ? dMov[i][9].toString().trim().toUpperCase() : "SPOT";
      const origem = dMov[i][6];

      // Alertas de atraso (SPOT maior que 7 dias)
      let atrasado = false;
      let diasEmprestimo = 0;
      if (dataSaida) {
        diasEmprestimo = Math.floor((hoje - dataSaida) / (1000 * 60 * 60 * 24));
        if (modalidade === "SPOT" && diasEmprestimo > 7) {
          atrasado = true;
        }
      }

      emprestimosAtivos.push({
        idMov: dMov[i][0],
        dataSaida: dataSaida ? Utilities.formatDate(dataSaida, "GMT-3", "dd/MM/yyyy HH:mm") : "---",
        tecnico: tecMap[idTec] || idTec || "Desconhecido",
        ferramenta: ferMap[idFer] ? ferMap[idFer].descricao : (idFer || "Não Encontrada"),
        qtd: qtd,
        modalidade: modalidade,
        origem: origem,
        dias: diasEmprestimo,
        atrasado: atrasado
      });
    }
  }

  // --- 3. Relatório: Histórico de Baixas Especiais (Danos, Perdas, Roubos) ---
  const baixasEspeciais = [];
  for(let i = 1; i < dMov.length; i++) {
    const obs = dMov[i][10] ? dMov[i][10].toString() : ""; // Coluna K (Índice 10)
    if (obs.includes("BAIXA ESPECIAL:")) {
      const idTec = dMov[i][3] ? dMov[i][3].toString().trim() : "";
      const idFer = dMov[i][4] ? dMov[i][4].toString().trim() : "";
      const dataDev = dMov[i][8] ? new Date(dMov[i][8]) : null;
      
      baixasEspeciais.push({
        data: dataDev ? Utilities.formatDate(dataDev, "GMT-3", "dd/MM/yyyy") : "---",
        tecnico: tecMap[idTec] || idTec || "---",
        ferramenta: ferMap[idFer] ? ferMap[idFer].descricao : idFer,
        qtd: dMov[i][5],
        detalhe: obs.replace("BAIXA ESPECIAL: ", "")
      });
    }
  }

  // --- 4. Relatório: Inventário e Valoração Completa ---
  const inventario = Object.keys(ferMap).map(id => {
    const f = ferMap[id];
    const vUnit = precoMedioMap[id] || 0;
    return {
      id: id,
      descricao: f.descricao,
      tipo: f.tipo,
      local: f.localizacao,
      disponivel: f.disponivel,
      emprestado: (f.emMaletas + f.emSpot),
      total: f.total,
      valorUnit: vUnit,
      valorTotal: f.total * vUnit
    };
  });

  // --- 5. Relatório: Histórico de Entradas Financeiras ---
  const compras = [];
  for(let i = 1; i < dNf.length; i++) {
    const data = dNf[i][0] ? new Date(dNf[i][0]) : null;
    const nf = dNf[i][1];
    const forn = dNf[i][2];
    const idFer = dNf[i][3] ? dNf[i][3].toString().trim() : "";
    const qtd = Number(dNf[i][4]) || 0;
    const unit = Number(dNf[i][5]) || 0;
    const ipi = Number(dNf[i][6]) || 0;
    const valorTotal = qtd * unit * (1 + (ipi / 100));

    compras.push({
      data: data ? Utilities.formatDate(data, "GMT-3", "dd/MM/yyyy") : "---",
      nf: nf,
      fornecedor: forn,
      ferramenta: ferMap[idFer] ? ferMap[idFer].descricao : idFer,
      qtd: qtd,
      total: valorTotal
    });
  }

  return {
    kpis: {
      totalItensEstoque,
      valorTotalEstoque: valorTotalEstoque,
      totalEmprestados,
      itensCriticos
    },
    emprestimosAtivos,
    baixasEspeciais,
    inventario,
    compras
  };
}