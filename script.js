// Проект ID-121125: КЦ и чат с оператором

const $fio = document.getElementById('fio');
const $btn = document.getElementById('findBtn');
const $banksList = document.getElementById('banks');
const $details = document.getElementById('details');
const $container = document.querySelector('.container');

let excelData = null;
let instructionTemplates = null;
let instructionLinks = null; // Данные с листа "Ссылки на инструкции"

// Маппинг колонок банков
const BANK_COLUMNS = {
  'Сбер': { 
    participation: ['Участие Сбер', 'Сбер'], 
    call: ['Сценарий Сбер (звонок)', 'Сбер (звонок)'], 
    chat: ['Сценарий Сбер (чат)', 'Сбер (чат)'] 
  },
  'ВТБ': { 
    participation: ['Участие ВТБ', 'ВТБ'], 
    call: ['Сценарий ВТБ (звонок)', 'ВТБ (звонок)'], 
    chat: ['Сценарий ВТБ (чат)', 'ВТБ (чат)'] 
  },
  'Т-Банк': { 
    participation: ['Участие Т-Банк', 'Т-Банк'], 
    call: ['Сценарий Т-Банк (звонок)', 'Т-Банк (звонок)'], 
    chat: ['Сценарий Т-Банк (чат)', 'Т-Банк (чат)'] 
  },
  'Альфа банк': { 
    participation: ['Участие Альфа банк', 'Альфа банк'], 
    call: ['Сценарий Альфа банк (звонок)', 'Альфа банк (звонок)'], 
    chat: ['Сценарий Альфа банк (чат)', 'Альфа банк (чат)'] 
  }
};

// Функция для поиска значения в колонке по нескольким возможным названиям
function findColumnValue(row, possibleNames) {
  for (const name of possibleNames) {
    if (row[name] !== undefined && row[name] !== null && row[name] !== '') {
      return String(row[name]).trim();
    }
  }
  return '';
}

// Управление состояниями выполнения проверок
function getCompletionKey(fio, bank, type) {
  return `121125_${fio}_${bank}_${type}`.replace(/\s+/g, '_');
}

function getCompletionStatus(fio, bank, type) {
  const key = getCompletionKey(fio, bank, type);
  return localStorage.getItem(`completion_${key}`) === 'true';
}

function setCompletionStatus(fio, bank, type, completed) {
  const key = getCompletionKey(fio, bank, type);
  if (completed) {
    localStorage.setItem(`completion_${key}`, 'true');
  } else {
    localStorage.removeItem(`completion_${key}`);
  }
}

function toggleCompletion(fio, bank, type) {
  const currentStatus = getCompletionStatus(fio, bank, type);
  setCompletionStatus(fio, bank, type, !currentStatus);
  return !currentStatus;
}

const statusIndicator = document.createElement('div');
statusIndicator.className = 'status-indicator';
document.body.appendChild(statusIndicator);

function normalizeString(text) {
  if (!text) return '';
  return text.toString().toLowerCase().replace(/\s+/g, '').replace(/ё/g, 'е');
}

function htmlEscape(str) {
  const div = document.createElement('div');
  div.textContent = str ?? '';
  return div.innerHTML;
}

function showStatus(message, isError = false) {
  statusIndicator.textContent = message;
  statusIndicator.className = `status-indicator ${isError ? 'error' : 'success'} show`;
  setTimeout(() => statusIndicator.classList.remove('show'), 3000);
}

function showLoading(button, text = 'Загрузка...') {
  const originalText = button.textContent;
  button.disabled = true;
  button.innerHTML = `<span class="loading">${text}</span>`;
  return originalText;
}

function hideLoading(button, originalText) {
  button.disabled = false;
  button.textContent = originalText;
}

async function loadInstructionTemplates() {
  try {
    if (location.protocol === 'file:') {
      return false;
    }

    const response = await fetch('bank-instructions.json?v=' + Date.now(), { cache: 'no-store' });
    if (!response.ok) {
      throw new Error('JSON файл не найден');
    }

    instructionTemplates = await response.json();
    return true;
  } catch (e) {
    console.error('Ошибка загрузки шаблонов инструкций:', e);
    return false;
  }
}

// Функция для проверки, содержит ли значение слово "сценарий" и цифру
function hasScenario(value) {
  if (!value || typeof value !== 'string') return false;
  const normalized = value.toLowerCase().trim();
  // Проверяем на пустые значения, прочерки и т.д.
  if (normalized === '' || normalized === '-' || normalized === '—' || normalized === 'нет' || normalized === 'без обращения') {
    return false;
  }
  // Проверяем наличие слова "сценарий" И цифры
  const hasScenarioWord = normalized.includes('сценарий');
  const hasDigit = /\d/.test(normalized);
  return hasScenarioWord && hasDigit;
}

async function loadExcelFile() {
  try {
    if (location.protocol === 'file:') {
      showStatus('Откройте через http://localhost/ (не file://)', true);
      return false;
    }

    const fileName = 'ID-121125 - ВТБ КЦ - Для загрузки.xlsx';
    const withBust = `${fileName}?v=${Date.now()}`;
    const response = await fetch(encodeURI(withBust), { cache: 'no-store' });
    
    if (!response.ok) {
      throw new Error('Excel файл не найден');
    }

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    // Ищем лист с данными (обычно первый лист или "Отобранные участники")
    let sheetName = workbook.SheetNames.find(name => 
      name.toLowerCase().includes('отобранные') || 
      name.toLowerCase().includes('участники') ||
      name.toLowerCase().includes('для загрузки')
    ) || workbook.SheetNames[0];

    excelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

    // Загружаем данные с листа "Ссылки на инструкции"
    const linksSheetName = workbook.SheetNames.find(name => 
      name.toLowerCase().includes('ссылки') && name.toLowerCase().includes('инструкции')
    );
    
    if (linksSheetName) {
      // Читаем как массив массивов для сохранения структуры
      const linksSheet = workbook.Sheets[linksSheetName];
      const linksData = XLSX.utils.sheet_to_json(linksSheet, { header: 1, defval: '' });
      instructionLinks = linksData;
    } else {
      instructionLinks = [];
    }

    showStatus(`Excel загружен (${excelData.length} записей)`);
    return true;
  } catch (e) {
    console.error(e);
    showStatus('Ошибка загрузки Excel', true);
    return false;
  }
}

// Функция для поиска ссылки по банку и номеру сценария
function findInstructionLink(bankName, scenarioText) {
  console.log('=== ПОИСК ССЫЛКИ ===');
  console.log('Банк:', bankName);
  console.log('Текст сценария:', scenarioText);
  console.log('instructionLinks загружен:', !!instructionLinks, 'длина:', instructionLinks ? instructionLinks.length : 0);
  
  if (!instructionLinks || instructionLinks.length === 0) {
    console.log('ОШИБКА: instructionLinks пуст');
    return '';
  }
  if (!scenarioText || typeof scenarioText !== 'string') {
    console.log('ОШИБКА: scenarioText пуст или не строка');
    return '';
  }
  
  // Маппинг названий банков на индексы колонок (B=1, C=2, D=3, E=4)
  const bankColumnMap = {
    'Сбер': 1,
    'ВТБ': 2,
    'Т-Банк': 3,
    'Альфа банк': 4,
    'Альфа-банк': 4
  };
  
  // Определяем индекс колонки для банка
  const columnIndex = bankColumnMap[bankName];
  console.log('Индекс колонки для банка', bankName, ':', columnIndex);
  if (columnIndex === undefined) {
    console.log('ОШИБКА: Банк не найден в маппинге');
    return '';
  }
  
  // Извлекаем номер сценария из текста (например, "Сценарий 1" -> "Сценарий 1")
  const scenarioMatch = scenarioText.match(/сценарий\s*\d+/i);
  if (!scenarioMatch) {
    console.log('ОШИБКА: Не найден номер сценария в тексте:', scenarioText);
    return '';
  }
  
  const scenarioName = scenarioMatch[0].trim();
  const normalizedScenario = scenarioName.toLowerCase();
  console.log('Ищем сценарий:', scenarioName, 'нормализованный:', normalizedScenario);
  
  // Пропускаем первую строку (заголовки)
  console.log('Проверяем строки на листе "Ссылки на инструкции":');
  for (let i = 1; i < instructionLinks.length; i++) {
    const row = instructionLinks[i];
    if (!row || row.length === 0) continue;
    
    // Проверяем первую колонку (A, индекс 0) на совпадение с номером сценария
    const firstCell = String(row[0] || '').trim();
    if (!firstCell) continue;
    
    const normalizedFirstCell = firstCell.toLowerCase();
    console.log(`  Строка ${i}: "${firstCell}" (нормализовано: "${normalizedFirstCell}")`);
    
    // Ищем точное совпадение (например, "сценарий 1" === "сценарий 1")
    if (normalizedFirstCell === normalizedScenario) {
      console.log('  ✓ НАЙДЕНО СОВПАДЕНИЕ!');
      // Берем значение из соответствующей колонки банка
      const link = String(row[columnIndex] || '').trim();
      console.log(`  Значение в колонке ${columnIndex}: "${link}"`);
      if (link) {
        console.log('  ✓ ССЫЛКА НАЙДЕНА:', link);
        return link;
      } else {
        console.log('  ✗ Ссылка пуста в ячейке');
      }
    }
  }
  
  console.log('✗ ССЫЛКА НЕ НАЙДЕНА');
  return '';
}

async function findBanks(fio) {
  if (!excelData) {
    const ok = await loadExcelFile();
    if (!ok) return [];
  }

  const normalizedFio = normalizeString(fio);
  const results = [];

  excelData.forEach(row => {
    // Пробуем разные варианты названия колонки с ФИО
    const testerFio = normalizeString(
      row['Укажите Ваши ФИО'] || 
      row['ФИО'] || 
      row['ФИО тестировщика'] || 
      ''
    );
    
    if (testerFio.includes(normalizedFio)) {
      // Проверяем каждый банк
      Object.keys(BANK_COLUMNS).forEach(bankName => {
        const bankConfig = BANK_COLUMNS[bankName];
        const participation = findColumnValue(row, bankConfig.participation);
        
        if (participation.toLowerCase() === 'да') {
          const callScenarioRaw = findColumnValue(row, bankConfig.call);
          const chatScenarioRaw = findColumnValue(row, bankConfig.chat);
          
          // Проверяем наличие слова "сценарий" в ячейках
          const callScenario = hasScenario(callScenarioRaw) ? callScenarioRaw : '';
          const chatScenario = hasScenario(chatScenarioRaw) ? chatScenarioRaw : '';
          
          // Добавляем банк если есть хотя бы один сценарий
          if (callScenario || chatScenario) {
            results.push({
              bank: bankName,
              callScenario: callScenario,
              chatScenario: chatScenario,
              fio: row['Укажите Ваши ФИО'] || row['ФИО'] || row['ФИО тестировщика'] || ''
            });
          }
        }
      });
    }
  });

  // Убираем дубликаты банков (если один банк встречается несколько раз)
  const uniqueBanks = [];
  const seenBanks = new Set();
  
  results.forEach(item => {
    if (!seenBanks.has(item.bank)) {
      seenBanks.add(item.bank);
      uniqueBanks.push(item);
    } else {
      // Если банк уже есть, обновляем сценарии если они есть
      const existing = uniqueBanks.find(b => b.bank === item.bank);
      if (item.callScenario && !existing.callScenario) existing.callScenario = item.callScenario;
      if (item.chatScenario && !existing.chatScenario) existing.chatScenario = item.chatScenario;
    }
  });

  return uniqueBanks;
}

function renderBanks(banks, fio) {
  const $instructionText = document.getElementById('instructionText');
  
  if (!banks || banks.length === 0) {
    $banksList.innerHTML = '<div class="bank-item">Банки не найдены для данного ФИО</div>';
    $banksList.style.display = 'block';
    $instructionText.style.display = 'none';
    $container.classList.add('with-result');
    return;
  }
  
  // Показываем текст только когда есть найденные банки
  $instructionText.style.display = 'block';

  const html = banks.map(bank => {
    // Проверяем наличие сценариев через функцию hasScenario
    const hasCallScenario = hasScenario(bank.callScenario);
    const hasChatScenario = hasScenario(bank.chatScenario);
    
    // Получаем статусы завершения только для существующих сценариев
    const callCompleted = hasCallScenario ? getCompletionStatus(fio, bank.bank, 'call') : false;
    const chatCompleted = hasChatScenario ? getCompletionStatus(fio, bank.bank, 'chat') : false;
    
    // Банк завершен если завершены все доступные сценарии
    let allCompleted = false;
    if (hasCallScenario && hasChatScenario) {
      allCompleted = callCompleted && chatCompleted;
    } else if (hasCallScenario) {
      allCompleted = callCompleted;
    } else if (hasChatScenario) {
      allCompleted = chatCompleted;
    }
    
    const statusClass = allCompleted ? 'completed' : 'pending';
    
    // Формируем HTML для чекбоксов только для существующих сценариев
    let completionItemsHtml = '';
    if (hasCallScenario) {
      completionItemsHtml += `
        <div class="completion-item">
          <span class="completion-label">Звонок:</span>
          <div class="completion-toggle ${callCompleted ? 'checked' : ''}" data-type="call" title="${callCompleted ? 'Отменить выполнение' : 'Отметить как выполненное'}"></div>
        </div>
      `;
    }
    if (hasChatScenario) {
      completionItemsHtml += `
        <div class="completion-item">
          <span class="completion-label">Чат:</span>
          <div class="completion-toggle ${chatCompleted ? 'checked' : ''}" data-type="chat" title="${chatCompleted ? 'Отменить выполнение' : 'Отметить как выполненное'}"></div>
        </div>
      `;
    }
    
    return `
      <div class="bank-item ${statusClass}" data-bank="${htmlEscape(bank.bank)}" data-fio="${htmlEscape(fio)}">
        <div class="bank-header">
          <strong>${htmlEscape(bank.bank)}</strong>
          <span class="bank-subtitle">(проверка звонка в КЦ и обращения в чат с оператором)</span>
        </div>
        <div class="completion-status">
          ${completionItemsHtml}
        </div>
      </div>
    `;
  }).join('');

  $banksList.innerHTML = html;
  $banksList.style.display = 'block';
  $container.classList.add('with-result');

  // Обработчики событий
  document.querySelectorAll('.bank-item').forEach(item => {
    const bank = item.dataset.bank;
    const fioValue = item.dataset.fio;
    
    // Клик на банк открывает инструкцию
    item.addEventListener('click', (e) => {
      if (!e.target.classList.contains('completion-toggle')) {
        showInstruction(bank, fioValue, banks.find(b => b.bank === bank));
      }
    });

    // Обработчики для галочек
    item.querySelectorAll('.completion-toggle').forEach(toggle => {
      toggle.addEventListener('click', (e) => {
        e.stopPropagation();
        const type = toggle.dataset.type;
        const newStatus = toggleCompletion(fioValue, bank, type);
        
        if (newStatus) {
          toggle.classList.add('checked');
          toggle.title = 'Отменить выполнение';
          showStatus(`Проверка "${bank}" (${type === 'call' ? 'звонок' : 'чат'}) отмечена как выполненная`, false);
        } else {
          toggle.classList.remove('checked');
          toggle.title = 'Отметить как выполненное';
          showStatus(`Отметка выполнения снята с "${bank}" (${type === 'call' ? 'звонок' : 'чат'})`, false);
        }
        
        // Обновляем статус банка с учетом только существующих сценариев
        const bankData = banks.find(b => b.bank === bank);
        if (bankData) {
          const hasCallScenario = hasScenario(bankData.callScenario);
          const hasChatScenario = hasScenario(bankData.chatScenario);
          
          const callCompleted = hasCallScenario ? getCompletionStatus(fioValue, bank, 'call') : false;
          const chatCompleted = hasChatScenario ? getCompletionStatus(fioValue, bank, 'chat') : false;
          
          let allCompleted = false;
          if (hasCallScenario && hasChatScenario) {
            allCompleted = callCompleted && chatCompleted;
          } else if (hasCallScenario) {
            allCompleted = callCompleted;
          } else if (hasChatScenario) {
            allCompleted = chatCompleted;
          }
          
          if (allCompleted) {
            item.classList.remove('pending');
            item.classList.add('completed');
          } else {
            item.classList.remove('completed');
            item.classList.add('pending');
          }
        }
      });
    });
  });
}

function showInstruction(bank, fio, bankData) {
  if (!bankData) return;
  
  $details.style.display = 'block';
  $details.querySelector('.tester').innerHTML = `Тестировщик: <strong>${htmlEscape(fio)}</strong>`;
  
  // Проверяем наличие сценариев через функцию hasScenario
  const hasCallScenario = hasScenario(bankData.callScenario);
  const hasChatScenario = hasScenario(bankData.chatScenario);
  
  // Сохраняем данные банка в data-атрибуте для доступа из обработчиков
  const bankInfoDiv = $details.querySelector('.bank-info');
  
  // Формируем HTML только для доступных вкладок
  let instructionTypesHtml = '';
  if (hasCallScenario) {
    instructionTypesHtml += `
      <div class="instruction-type active" data-type="call">
        <strong>Инструкция для звонка в колл-центр</strong>
      </div>
    `;
  }
  if (hasChatScenario) {
    instructionTypesHtml += `
      <div class="instruction-type active" data-type="chat">
        <strong>Инструкция для обращения в чат поддержки</strong>
      </div>
    `;
  }
  
  bankInfoDiv.innerHTML = `
    <div><strong>${htmlEscape(bank)}</strong></div>
    <div class="instruction-types">
      ${instructionTypesHtml}
    </div>
  `;
  
  // Сохраняем данные банка для использования в обработчиках
  bankInfoDiv.dataset.bankData = JSON.stringify(bankData);
  
  // Показываем первую доступную инструкцию
  let initialType = hasCallScenario ? 'call' : (hasChatScenario ? 'chat' : null);
  if (initialType) {
    displayInstruction(bankData, initialType);
  } else {
    $details.querySelector('.instruction').innerHTML = '<p>Инструкции не найдены для данного банка.</p>';
  }
  
  // Обработчики переключения между инструкциями
  $details.querySelectorAll('.instruction-type').forEach(typeEl => {
    // Удаляем старые обработчики если они есть
    const newTypeEl = typeEl.cloneNode(true);
    typeEl.parentNode.replaceChild(newTypeEl, typeEl);
    
    newTypeEl.addEventListener('click', () => {
      const type = newTypeEl.dataset.type;
      if (newTypeEl.classList.contains('active')) {
        const savedBankData = JSON.parse(bankInfoDiv.dataset.bankData);
        displayInstruction(savedBankData, type);
        // Обновляем активный класс
        $details.querySelectorAll('.instruction-type').forEach(el => el.classList.remove('selected'));
        newTypeEl.classList.add('selected');
      }
    });
  });
  
  // Устанавливаем первую инструкцию как выбранную
  if (initialType) {
    const firstTypeEl = $details.querySelector(`.instruction-type[data-type="${initialType}"]`);
    if (firstTypeEl) firstTypeEl.classList.add('selected');
  }
  
  $details.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function displayInstruction(bankData, type) {
  const instructionDiv = $details.querySelector('.instruction');
  const scenario = type === 'call' ? bankData.callScenario : bankData.chatScenario;
  
  if (!hasScenario(scenario)) {
    instructionDiv.innerHTML = `<p>Инструкция для ${type === 'call' ? 'звонка' : 'чата'} не найдена.</p>`;
    return;
  }
  
  // Используем шаблон из JSON если доступен
  if (instructionTemplates) {
    const templateKey = type === 'call' ? 'call_center' : 'chat';
    const template = instructionTemplates[templateKey];
    
    if (template) {
      let content = template.content;
      
      // Заменяем плейсхолдеры (БЕЗ кавычек вокруг названия банка)
      content = content.replace(/{BANK_NAME}/g, htmlEscape(bankData.bank || ''));
      content = content.replace(/{SCENARIO}/g, htmlEscape(scenario));
      
      // Ищем ссылку на листе "Ссылки на инструкции" по банку и номеру сценария
      const link = findInstructionLink(bankData.bank, scenario);
      console.log('Поиск ссылки для банка:', bankData.bank, 'сценарий:', scenario, 'найдена ссылка:', link);
      const linkHtml = link ? `<a href="${link}" target="_blank" rel="noopener noreferrer">${link}</a>` : 'ссылка не найдена';
      content = content.replace(/{LINK}/g, linkHtml);
      
      // Преобразуем переносы строк
      content = content.replace(/\n/g, '<br>');
      
      instructionDiv.innerHTML = content;
      
      // Активируем сворачиваемые блоки
      initCollapsibles();
      
      return;
    }
  }
  
  // Fallback на старый формат если JSON не загружен
  let formattedText = scenario
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  
  formattedText = makeLinksClickable(formattedText);
  const hasHtml = /<\/?[a-z][\s\S]*?>/i.test(formattedText);
  formattedText = hasHtml ? formattedText : formattedText.replace(/\n/g, '<br>');
  
  instructionDiv.innerHTML = formattedText;
}

function makeLinksClickable(text) {
  const urlRegex = /(https?:\/\/[^\s<>"]+)/g;
  return text.replace(urlRegex, '<a href="$1" target="_blank" rel="noopener noreferrer">$1</a>');
}

function initCollapsibles() {
  // Удаляем старые обработчики и добавляем новые
  document.querySelectorAll('.collapsible-header').forEach(header => {
    const newHeader = header.cloneNode(true);
    header.parentNode.replaceChild(newHeader, header);
    
    newHeader.addEventListener('click', () => {
      const collapsible = newHeader.parentElement;
      collapsible.classList.toggle('active');
    });
  });
}

async function performSearch() {
  const fio = $fio.value.trim();
  if (!fio) { 
    showStatus('Введите ФИО', true); 
    $fio.focus(); 
    return; 
  }
  
  const orig = showLoading($btn, 'Поиск банков...');
  try {
    const banks = await findBanks(fio);
    renderBanks(banks, fio);
    $details.style.display = 'none';
  } catch (e) {
    console.error(e); 
    showStatus('Ошибка поиска', true);
  } finally { 
    hideLoading($btn, orig); 
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  $fio.focus();
  $btn.addEventListener('click', performSearch);
  $fio.addEventListener('keypress', e => { 
    if (e.key === 'Enter') performSearch(); 
  });
  await loadInstructionTemplates();
  await loadExcelFile();
});

