// СПРАВОЧНИК РИЭЛТЕРОВ — единственный источник данных: config/agents.json
// Этот файл загружает его синхронно, чтобы window.AGENTS_CONFIG был готов
// до того, как realtor-service.js и app.js начнут исполняться.
/* eslint-disable no-sync */
(function () {
  try {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'config/agents.json', false); // синхронный запрос
    xhr.send(null);
    if (xhr.status === 200 || xhr.status === 0) {
      window.AGENTS_CONFIG = JSON.parse(xhr.responseText);
    } else {
      throw new Error('HTTP ' + xhr.status);
    }
  } catch (e) {
    console.error('[agents-config] Не удалось загрузить config/agents.json:', e);
    window.AGENTS_CONFIG = { agents: [] };
  }
}());
