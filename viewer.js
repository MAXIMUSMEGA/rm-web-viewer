'use strict';

const source = document.getElementById('source');
const frame = document.getElementById('frame');
const editor = document.getElementById('editor');
const empty = document.getElementById('empty');
const status = document.getElementById('status');
const apply = document.getElementById('apply');
const clear = document.getElementById('clear');
const settingsKey = 'rm.webviewer.url.v1';

let settings = null;
let current = '';
let sourceChanged = false;

function parseSource(input) {
  let value = input.trim();
  if (value.startsWith('<')) {
    const template = document.createElement('template');
    template.innerHTML = value;
    const iframe = template.content.querySelector('iframe[src]');
    value = iframe?.getAttribute('src')?.trim();
    if (!value) throw new Error('Нужен iframe с непустым атрибутом src.');
  }

  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error('Вставьте полный HTTPS-адрес.');
  }
  if (url.protocol !== 'https:' || url.username || url.password) {
    throw new Error('Разрешены только HTTPS-ссылки без логина и пароля в URL.');
  }
  return url.href;
}

function show(url) {
  current = url;
  source.value = url;
  empty.hidden = Boolean(url);
  frame.hidden = !url;
  if (url) frame.src = url;
  else frame.removeAttribute('src');
}

async function updateSource(url) {
  sourceChanged = true;
  show(url);
  if (!settings) {
    status.textContent = 'Предпросмотр. Ссылка не сохранена в книге.';
    return;
  }

  apply.disabled = true;
  clear.disabled = true;
  status.textContent = 'Сохраняю настройки надстройки…';
  try {
    if (url) settings.set(settingsKey, url);
    else settings.remove(settingsKey);
    await new Promise((resolve, reject) => {
      settings.saveAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error.message));
      });
    });
    status.textContent = url ? 'Ссылка сохранена в книге.' : 'Ссылка удалена из настроек книги.';
    if (url) editor.hidden = true;
  } catch (error) {
    status.textContent = `Не удалось сохранить настройки. ${error.message}`;
  } finally {
    apply.disabled = false;
    clear.disabled = false;
  }
}

apply.onclick = () => {
  try {
    void updateSource(parseSource(source.value));
  } catch (error) {
    status.textContent = error.message;
  }
};
clear.onclick = () => { void updateSource(''); };
document.getElementById('toggle').onclick = () => { editor.hidden = !editor.hidden; };
document.getElementById('reload').onclick = () => { if (current) frame.src = current; };
document.getElementById('open').onclick = () => {
  if (current) window.open(current, '_blank', 'noopener,noreferrer');
};

if (window.Office) {
  Office.onReady(info => {
    if (info.host !== Office.HostType.Excel) return;
    try {
      settings = Office.context.document.settings;
      if (sourceChanged || source.value.trim()) return;
      const saved = settings.get(settingsKey);
      if (saved) {
        show(parseSource(saved));
        editor.hidden = true;
      } else {
        status.textContent = 'Готово к работе.';
      }
    } catch (error) {
      status.textContent = `Не удалось загрузить настройки книги. ${error.message}`;
    }
  });
}
