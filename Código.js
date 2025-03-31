/**
 * @OnlyCurrentDoc // Limits the script to only accessing the current spreadsheet.
 */

/**
 * Cria um menu personalizado na interface do usuário do Google Sheets
 * quando a planilha é aberta.
 *
 * @param {object} e O objeto de evento (não utilizado aqui, mas parte da assinatura padrão).
 */
function onOpen(e) {
  SpreadsheetApp.getUi() // Obtém o objeto da interface do usuário para esta planilha.
      .createMenu('✨ Agenda Fácil') // Cria um menu com o título especificado.
      .addItem('➕ Adicionar Eventos ao Calendário', 'addEventsFromSheet') // Adiciona um item que executa a função 'addEventsFromSheet'.
      .addToUi(); // Adiciona o menu à interface do usuário da planilha.
}


/**
 * Parses URL query parameters into an object.
 * Handles URL decoding for keys and values using decodeURIComponent.
 *
 * @param {string} url The URL string containing query parameters.
 * @return {object} An object with key-value pairs from the URL.
 */
function parseUrlParameters_(url) {
  var params = { text: '', dates: '', details: '', location: '', recur: '' };
  var queryString = url.split('?')[1];
  if (!queryString) {
    Logger.log('URL não contém string de consulta (query string): ' + url);
    return params; // Retorna objeto vazio se não houver parâmetros
  }

  var pairs = queryString.split('&');
  for (var i = 0; i < pairs.length; i++) {
    var pair = pairs[i].split('=');
    // Decodifica a chave (menos comum precisar disso)
    var key = decodeURIComponent(pair[0]);
    // Decodifica o valor, substituindo '+' por espaço PRIMEIRO, depois usando decodeURIComponent.
    // Lida com casos onde o valor pode estar ausente (ex: &details=).
    var value = pair[1] ? decodeURIComponent(pair[1].replace(/\+/g, ' ')) : '';

    // Usa hasOwnProperty para evitar problemas com propriedades herdadas (boa prática)
     if (params.hasOwnProperty(key)) {
       params[key] = value;
    } else {
       // Logger.log('Parâmetro desconhecido ignorado: ' + key); // Opcional: logar parâmetros não esperados
    }
  }
  // Logger.log('Parâmetros extraídos: ' + JSON.stringify(params)); // Descomente para depuração
  return params;
}

/**
 * Lê URLs de links de criação de eventos do Google Agenda da coluna A
 * da planilha ativa e cria os eventos correspondentes no calendário padrão do usuário.
 * Esta função é chamada pelo item de menu personalizado.
 */
function addEventsFromSheet() {
  // Obter a planilha e a aba ativa
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Obter todas as URLs da coluna A (ignora o cabeçalho se houver)
  // Assume que os links começam na linha 1. Se começar em outra linha, ajuste o A1.
  var firstRow = 1; // Primeira linha com dados
  var lastRow = sheet.getLastRow();
  if (lastRow < firstRow) {
      Logger.log('Nenhuma URL encontrada na coluna A a partir da linha ' + firstRow);
      SpreadsheetApp.getUi().alert('Nenhuma URL encontrada na coluna A.');
      return;
  }
  var range = sheet.getRange("A" + firstRow + ":A" + lastRow);
  var urls = range.getValues(); // Obtem [[url1], [url2], ...]

  // Obter o calendário padrão
  var calendar = CalendarApp.getDefaultCalendar();
  if (!calendar) {
    Logger.log('Calendário padrão não encontrado.');
    SpreadsheetApp.getUi().alert('Calendário padrão não encontrado.');
    return;
  }
  Logger.log('Usando o calendário: ' + calendar.getName());

  var eventsCreated = 0;
  var errorsEncountered = 0;

  // Iterar sobre cada linha que contém uma URL
  for (var i = 0; i < urls.length; i++) {
    var currentRow = firstRow + i; // Linha atual na planilha para logs
    var url = urls[i][0]; // Pega a URL da célula

    // Pula linhas vazias ou que não são URLs válidas do calendário
    if (!url || typeof url !== 'string' || !url.startsWith('https://www.google.com/calendar/')) {
      // Logger.log('Linha ' + currentRow + ' pulada (vazia ou formato inválido).');
      continue;
    }

    try {
      // Parsear os parâmetros da URL
      var params = parseUrlParameters_(url);

      // Verificar se os parâmetros essenciais foram extraídos
      if (!params.text || !params.dates) {
        Logger.log('Erro na linha ' + currentRow + ': Não foi possível extrair título ou datas da URL: ' + url);
        errorsEncountered++;
        continue;
      }

      var datesStr = params.dates;
      var startDate, endDate;
      var options = { description: params.details || '', location: params.location || '' };

      // Determinar se é evento de dia inteiro ou com horário específico
      if (datesStr.includes('T')) {
        // Evento com horário: format YYYYMMDDTHHMMSS/YYYYMMDDTHHMMSS
        var parts = datesStr.split('/');
        if (parts.length !== 2) throw new Error('Formato de data/hora inválido: ' + datesStr);
        var startStr = parts[0];
        var endStr = parts[1];

        // Parsear YYYYMMDDTHHMMSS manualmente para criar objetos Date
        startDate = new Date(
          parseInt(startStr.substring(0, 4), 10),  // Ano
          parseInt(startStr.substring(4, 6), 10) - 1, // Mês (0-indexado)
          parseInt(startStr.substring(6, 8), 10),  // Dia
          parseInt(startStr.substring(9, 11), 10), // Hora
          parseInt(startStr.substring(11, 13), 10),// Minuto
          parseInt(startStr.substring(13, 15), 10) // Segundo
        );
        endDate = new Date(
          parseInt(endStr.substring(0, 4), 10),
          parseInt(endStr.substring(4, 6), 10) - 1,
          parseInt(endStr.substring(6, 8), 10),
          parseInt(endStr.substring(9, 11), 10),
          parseInt(endStr.substring(11, 13), 10),
          parseInt(endStr.substring(13, 15), 10)
        );

         if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
           throw new Error('Data de início ou fim inválida após parse: ' + startStr + ' / ' + endStr);
        }

        // Criar o evento com horário
        calendar.createEvent(params.text, startDate, endDate, options);
        Logger.log('Evento com horário criado: "' + params.text + '" em ' + startDate.toLocaleString());
        eventsCreated++;

      } else {
        // Evento de dia inteiro: format YYYYMMDD/YYYYMM(DD+1)
        var parts = datesStr.split('/');
        if (parts.length !== 2) throw new Error('Formato de data inválido para evento de dia inteiro: ' + datesStr);
        var startStr = parts[0]; // createAllDayEvent só precisa da data de início

        // Parsear YYYYMMDD manualmente
        startDate = new Date(
          parseInt(startStr.substring(0, 4), 10),  // Ano
          parseInt(startStr.substring(4, 6), 10) - 1, // Mês (0-indexado)
          parseInt(startStr.substring(6, 8), 10)   // Dia
        );

        if (isNaN(startDate.getTime())) {
           throw new Error('Data de início inválida após parse: ' + startStr);
        }

        // Criar o evento de dia inteiro
        calendar.createAllDayEvent(params.text, startDate, options);
        Logger.log('Evento de dia inteiro criado: "' + params.text + '" em ' + startDate.toLocaleDateString());
        eventsCreated++;
      }

    } catch (e) {
      Logger.log('ERRO ao processar URL na linha ' + currentRow + ': ' + url + ' - Detalhes: ' + e.message);
      // Se quiser ver o stack trace completo no log, descomente a linha abaixo
      // Logger.log('Stack trace: ' + e.stack);
      errorsEncountered++;
    }
  } // Fim do loop

  // Mensagem final
  var message = 'Processamento concluído.\nEventos criados: ' + eventsCreated + '\nErros encontrados: ' + errorsEncountered;
  Logger.log(message);
  SpreadsheetApp.getUi().alert(message); // Mostra um pop-up para o usuário
}