// ==UserScript==
// @name         OneDrive Natural Sort
// @namespace    https://github.com/ShaddockNH3/OneDrive-Natural-Sort
// @version      2.2.0
// @description  Sort OneDrive native RenderListDataAsStream responses with numeric/natural order.
// @author       ShaddockNH3
// @license      MIT
// @homepageURL  https://github.com/ShaddockNH3/OneDrive-Natural-Sort
// @supportURL   https://github.com/ShaddockNH3/OneDrive-Natural-Sort/issues
// @updateURL    https://raw.githubusercontent.com/ShaddockNH3/OneDrive-Natural-Sort/main/onedrive-natural-sort.user.js
// @downloadURL  https://raw.githubusercontent.com/ShaddockNH3/OneDrive-Natural-Sort/main/onedrive-natural-sort.user.js
// @match        https://onedrive.live.com/*
// @match        https://*.sharepoint.com/*
// @match        https://*.my.sharepoint.com/*
// @grant        none
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';

  const DIRECTION_KEY = 'onedrive-natural-sort-direction';
  const UI_ATTR = 'data-onedrive-natural-sort-ui';
  const PATCH_MARK = '__onedriveNaturalSortPatched';
  const FULL_ROW_LIMIT = 5000;

  const collator = new Intl.Collator(undefined, {
    numeric: true,
    sensitivity: 'base',
    ignorePunctuation: false,
  });

  let direction = localStorage.getItem(DIRECTION_KEY) === 'desc' ? 'desc' : 'asc';
  let sortedResponseCount = 0;
  const sortedXhrTextCache = new WeakMap();

  function isRenderListDataUrl(url) {
    return /RenderListDataAsStream/i.test(String(url || ''));
  }

  function getRowsContainer(payload) {
    if (Array.isArray(payload?.ListData?.Row)) return payload.ListData;
    if (Array.isArray(payload?.Row)) return payload;
    return null;
  }

  function getItemName(item) {
    const directName = item.FileLeafRef || item.File_x0020_Name || item.LinkFilename || item.LinkFilenameNoMenu || item.Name;
    if (directName) return String(directName);

    const fileRef = item.FileRef || item.ServerRelativeUrl || '';
    const lastPart = String(fileRef).split('/').pop() || '';
    try {
      return decodeURIComponent(lastPart);
    } catch {
      return lastPart;
    }
  }

  function compareRows(left, right) {
    const leftIsFolder = String(left.FSObjType || '0') === '1';
    const rightIsFolder = String(right.FSObjType || '0') === '1';

    if (leftIsFolder !== rightIsFolder) return leftIsFolder ? -1 : 1;

    const result = collator.compare(getItemName(left), getItemName(right));
    if (result !== 0) return direction === 'asc' ? result : -result;

    return Number(left.ID || 0) - Number(right.ID || 0);
  }

  function sortPayload(payload) {
    const container = getRowsContainer(payload);
    if (!container || !Array.isArray(container.Row) || container.Row.length < 2) return false;

    container.Row.sort(compareRows);
    container.FirstRow = 1;
    container.LastRow = container.Row.length;
    container.RowLimit = Math.max(Number(container.RowLimit || 0), container.Row.length);
    container.NextHref = '';
    sortedResponseCount += 1;
    return true;
  }

  function parsePossiblyNestedPayload(text) {
    const payload = JSON.parse(text);
    if (typeof payload?.d?.RenderListDataAsStream === 'string') return JSON.parse(payload.d.RenderListDataAsStream);
    return payload;
  }

  function stringifyLikeOriginal(originalText, payload) {
    const original = JSON.parse(originalText);
    if (typeof original?.d?.RenderListDataAsStream === 'string') {
      original.d.RenderListDataAsStream = JSON.stringify(payload);
      return JSON.stringify(original);
    }
    return JSON.stringify(payload);
  }

  function trySortText(text) {
    if (!text || !/"(?:ListData|Row)"\s*:/.test(text)) return text;

    try {
      const payload = parsePossiblyNestedPayload(text);
      return sortPayload(payload) ? stringifyLikeOriginal(text, payload) : text;
    } catch {
      return text;
    }
  }

  function buildFullRowsBody(originalBody) {
    let body = {};

    try {
      body = originalBody ? JSON.parse(String(originalBody)) : {};
    } catch {
      body = {};
    }

    body.parameters = body.parameters || {};
    body.parameters.RenderOptions = body.parameters.RenderOptions || 2;
    body.parameters.DatesInUtc = body.parameters.DatesInUtc !== false;
    body.parameters.AddRequiredFields = true;
    body.parameters.OverrideViewXml = `<View><RowLimit>${FULL_ROW_LIMIT}</RowLimit></View>`;

    return JSON.stringify(body);
  }

  async function fetchFullRenderListText(url, init) {
    const headers = new Headers(init?.headers || {});
    headers.set('Accept', 'application/json;odata=nometadata');
    headers.set('Content-Type', 'application/json;odata=nometadata');

    const response = await originalFetchRef(url, {
      method: 'POST',
      headers,
      credentials: init?.credentials || 'same-origin',
      body: buildFullRowsBody(init?.body),
    });

    if (!response.ok) return null;

    const text = await response.text();
    const sortedText = trySortText(text);
    return sortedText;
  }

  let originalFetchRef = null;

  function patchFetch() {
    if (!window.fetch || window.fetch[PATCH_MARK]) return;

    originalFetchRef = window.fetch;
    const patchedFetch = async function (...args) {
      const url = String(args[0]?.url || args[0] || '');
      const init = args[1] || {};

      if (isRenderListDataUrl(url)) {
        const fullText = await fetchFullRenderListText(url, init).catch(() => null);
        if (fullText) {
          return new Response(fullText, {
            status: 200,
            statusText: 'OK',
            headers: { 'Content-Type': 'application/json;odata=nometadata' },
          });
        }
      }

      const response = await originalFetchRef.apply(this, args);
      if (!isRenderListDataUrl(url || response.url)) return response;

      const originalText = await response.clone().text();
      const sortedText = trySortText(originalText);
      if (sortedText === originalText) return response;

      return new Response(sortedText, {
        status: response.status,
        statusText: response.statusText,
        headers: response.headers,
      });
    };

    patchedFetch[PATCH_MARK] = true;
    window.fetch = patchedFetch;
  }

  function patchXhr() {
    if (XMLHttpRequest.prototype.open[PATCH_MARK]) return;

    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;
    const responseTextDescriptor = Object.getOwnPropertyDescriptor(XMLHttpRequest.prototype, 'responseText');
    const responseDescriptor = Object.getOwnPropertyDescriptor(XMLHttpRequest.prototype, 'response');

    function getNativeResponseText(xhr) {
      return responseTextDescriptor?.get ? responseTextDescriptor.get.call(xhr) : '';
    }

    function getSortedXhrText(xhr) {
      if (!isRenderListDataUrl(xhr.__odnsUrl) || xhr.readyState !== 4) {
        return getNativeResponseText(xhr);
      }

      if (sortedXhrTextCache.has(xhr)) return sortedXhrTextCache.get(xhr);

      const originalText = String(getNativeResponseText(xhr) || '');
      const sortedText = trySortText(originalText);
      sortedXhrTextCache.set(xhr, sortedText);
      return sortedText;
    }

    XMLHttpRequest.prototype.open = function (method, url, ...rest) {
      this.__odnsMethod = method;
      this.__odnsUrl = String(url || '');
      sortedXhrTextCache.delete(this);
      return originalOpen.call(this, method, url, ...rest);
    };

    XMLHttpRequest.prototype.send = function (body) {
      if (isRenderListDataUrl(this.__odnsUrl)) {
        try {
          body = buildFullRowsBody(body);
        } catch {}
      }

      return originalSend.call(this, body);
    };

    if (responseTextDescriptor?.get && responseTextDescriptor.configurable !== false) {
      Object.defineProperty(XMLHttpRequest.prototype, 'responseText', {
        configurable: true,
        enumerable: responseTextDescriptor.enumerable,
        get() {
          return getSortedXhrText(this);
        },
      });
    }

    if (responseDescriptor?.get && responseDescriptor.configurable !== false) {
      Object.defineProperty(XMLHttpRequest.prototype, 'response', {
        configurable: true,
        enumerable: responseDescriptor.enumerable,
        get() {
          if (isRenderListDataUrl(this.__odnsUrl) && (!this.responseType || this.responseType === 'text')) {
            return getSortedXhrText(this);
          }
          return responseDescriptor.get.call(this);
        },
      });
    }

    XMLHttpRequest.prototype.open[PATCH_MARK] = true;
  }

  function installInterceptor() {
    patchFetch();
    patchXhr();
  }

  function reloadOneDrive() {
    location.reload();
  }

  function setDirection(nextDirection) {
    direction = nextDirection;
    localStorage.setItem(DIRECTION_KEY, direction);
    updateButtonLabels();
    updateStatus('reload to apply');
  }

  function updateStatus(text) {
    const status = document.querySelector('[data-onedrive-natural-sort-status]');
    if (status) status.textContent = text || `sorted ${sortedResponseCount}`;
  }

  function updateButtonLabels() {
    const ascButton = document.querySelector('[data-onedrive-natural-sort-action="asc"]');
    const descButton = document.querySelector('[data-onedrive-natural-sort-action="desc"]');

    if (ascButton) ascButton.setAttribute('aria-pressed', String(direction === 'asc'));
    if (descButton) descButton.setAttribute('aria-pressed', String(direction === 'desc'));
  }

  function addControlPanel() {
    if (document.querySelector(`[${UI_ATTR}]`)) return;

    const panel = document.createElement('div');
    panel.setAttribute(UI_ATTR, 'true');
    panel.innerHTML = `
      <button type="button" data-onedrive-natural-sort-action="asc" title="Native natural sort ascending">A-&gt;Z</button>
      <button type="button" data-onedrive-natural-sort-action="desc" title="Native natural sort descending">Z-&gt;A</button>
      <button type="button" data-onedrive-natural-sort-action="reload" title="Reload OneDrive to apply native sort">Reload</button>
      <span data-onedrive-natural-sort-status></span>
    `;

    const style = document.createElement('style');
    style.textContent = `
      [${UI_ATTR}] {
        position: fixed;
        right: 16px;
        bottom: 16px;
        z-index: 2147483647;
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 8px;
        border: 1px solid rgba(0, 0, 0, 0.18);
        border-radius: 8px;
        background: rgba(255, 255, 255, 0.96);
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.18);
        color: #1f1f1f;
        font: 12px/1.2 system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }

      [${UI_ATTR}] button {
        min-width: 54px;
        height: 28px;
        border: 1px solid rgba(0, 0, 0, 0.24);
        border-radius: 6px;
        background: #fff;
        color: #1f1f1f;
        cursor: pointer;
      }

      [${UI_ATTR}] button[aria-pressed="true"] {
        border-color: #0078d4;
        background: #0078d4;
        color: #fff;
      }

      [data-onedrive-natural-sort-status] {
        min-width: 90px;
        max-width: 180px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        color: #444;
      }
    `;

    panel.addEventListener('click', (event) => {
      const button = event.target.closest('button');
      if (!button) return;

      const action = button.dataset.onedriveNaturalSortAction;
      if (action === 'asc' || action === 'desc') setDirection(action);
      if (action === 'reload') reloadOneDrive();
    });

    document.documentElement.appendChild(style);
    document.body.appendChild(panel);
    updateButtonLabels();
    updateStatus('native sort on');
  }

  installInterceptor();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', addControlPanel, { once: true });
  } else {
    addControlPanel();
  }
})();