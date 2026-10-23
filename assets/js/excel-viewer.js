/* One instance per raw/output host. No workbook parsing or inactive-image preloading. */
(() => {
  let nextId = 0;
  class ExcelScreenshotViewer {
    constructor(host, group, title) {
      this.group = group;
      this.title = title;
      this.cache = new Map(); // Two recently requested images per instance, including pending loads.
      this.sequence = 0;
      this.scale = 1;
      this.fit = false;
      this.id = `excel-preview-${++nextId}`;
      this.root = document.createElement('div');
      this.root.className = 'excel-viewer';
      this.root.innerHTML = `<div class="ev-heading"><div><small>Excel workbook · Screenshot preview</small><strong class="ev-filename"></strong></div></div>
        <div class="ev-tabs" role="tablist"></div>
        <div class="ev-toolbar" role="group" aria-label="Preview controls">
          <button type="button" data-action="out" aria-label="Zoom out">−</button><output class="ev-scale" aria-label="Current zoom">100%</output><button type="button" data-action="in" aria-label="Zoom in">+</button>
          <button type="button" data-action="reset" aria-label="Reset zoom to 100 percent">100%</button><button type="button" data-action="fit">Fit Width</button>
          <button type="button" data-action="expand" class="ev-expand">Expand</button><button type="button" data-action="close" class="ev-expand" hidden>Close</button>
        </div><div class="ev-panel" role="tabpanel"><div class="ev-viewport" tabindex="0"></div></div>
        <div class="ev-status" role="status" aria-live="polite"></div><p class="ev-note"></p>`;
      host.replaceChildren(this.root);
      this.viewport = this.root.querySelector('.ev-viewport');
      this.panel = this.root.querySelector('.ev-panel');
      this.panel.id = `${this.id}-panel`;
      this.status = this.root.querySelector('.ev-status');
      this.note = this.root.querySelector('.ev-note');
      this.tablist = this.root.querySelector('.ev-tabs');
      this.tablist.setAttribute('aria-label', `${title} ${host.dataset.excelViewer || ''} worksheets`);
      this.root.querySelector('.ev-filename').textContent = group.workbook?.filename || 'Worksheet previews';
      if (group.workbook?.downloadAvailable) {
        const link = document.createElement('a');
        link.href = group.workbook.path;
        link.download = group.workbook.filename;
        link.textContent = 'Download workbook';
        this.root.querySelector('.ev-heading').append(link);
      }
      this.tabs = group.sheets.map((sheet, index) => {
        const tab = document.createElement('button');
        tab.type = 'button'; tab.id = `${this.id}-tab-${index}`;
        tab.setAttribute('role', 'tab'); tab.setAttribute('aria-controls', this.panel.id);
        tab.textContent = sheet.label;
        tab.addEventListener('click', () => this.select(index));
        this.tablist.append(tab);
        return tab;
      });
      this.tablist.addEventListener('keydown', event => {
        const index = this.tabs.indexOf(document.activeElement);
        if (index < 0) return;
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          this.select(index);
          return;
        }
        const target = { ArrowRight:(index + 1) % this.tabs.length, ArrowLeft:(index - 1 + this.tabs.length) % this.tabs.length, Home:0, End:this.tabs.length - 1 }[event.key];
        if (target === undefined) return;
        event.preventDefault();
        this.tabs.forEach((tab, i) => { tab.tabIndex = i === target ? 0 : -1; });
        this.tabs[target].focus();
        this.revealTab(this.tabs[target]); // Manual activation keeps focus separate from selection.
      });
      this.root.querySelector('.ev-toolbar').addEventListener('click', event => {
        const action = event.target.closest('button')?.dataset.action;
        if (action === 'expand') this.expand();
        else if (action === 'close') this.dialog.close();
        else if (this.image) {
          this.fit = action === 'fit';
          if (action === 'in') this.scale = Math.min(3, Math.round((this.scale + .25) * 100) / 100);
          if (action === 'out') this.scale = Math.max(.25, Math.round((this.scale - .25) * 100) / 100);
          if (action === 'reset') this.scale = 1;
          this.resize();
          if (action === 'reset' || action === 'fit') this.resetScroll();
        }
      });
      this.observer = new ResizeObserver(() => { if (this.fit) this.resize(); });
      this.observer.observe(this.viewport);
      this.index = Math.max(0, group.sheets.findIndex(s => s.worksheet === group.defaultWorksheet));
      this.updateTabs(); this.updateControls();
      this.status.textContent = this.tabs.length ? 'Preview loads when this viewer is in view.' : 'No approved screenshot previews are available.';
      if (this.tabs.length) {
        if ('IntersectionObserver' in window) {
          this.lazy = new IntersectionObserver(entries => {
            if (entries.some(entry => entry.isIntersecting)) { this.lazy.disconnect(); this.select(this.index); }
          }, { rootMargin:'100px' });
          this.lazy.observe(this.root);
        } else this.select(this.index);
      }
    }
    revealTab(tab) {
      const left = tab.offsetLeft - this.tablist.offsetLeft;
      if (left < this.tablist.scrollLeft) this.tablist.scrollLeft = left;
      else if (left + tab.offsetWidth > this.tablist.scrollLeft + this.tablist.clientWidth) this.tablist.scrollLeft = left + tab.offsetWidth - this.tablist.clientWidth;
    }
    updateTabs() {
      this.tabs.forEach((tab, i) => { tab.setAttribute('aria-selected', String(i === this.index)); tab.tabIndex = i === this.index ? 0 : -1; });
      const tab = this.tabs[this.index];
      if (tab) { this.panel.setAttribute('aria-labelledby', tab.id); this.viewport.setAttribute('aria-label', `${this.group.sheets[this.index].label} screenshot; scroll to explore`); this.revealTab(tab); }
    }
    load(sheet) {
      if (this.cache.has(sheet.image)) return this.cache.get(sheet.image);
      const promise = new Promise((resolve, reject) => {
        const image = new Image();
        image.decoding = 'async';
        image.alt = `${this.title}: ${sheet.label} worksheet screenshot. ${sheet.coverageNote || ''}`;
        image.onload = async () => { try { await image.decode(); resolve(image); } catch (error) { reject(error); } };
        image.onerror = () => reject(new Error('Image unavailable'));
        image.src = sheet.image;
      });
      this.cache.set(sheet.image, promise);
      while (this.cache.size > 2) this.cache.delete(this.cache.keys().next().value);
      promise.catch(() => { if (this.cache.get(sheet.image) === promise) this.cache.delete(sheet.image); });
      return promise;
    }
    async select(index) {
      this.lazy?.disconnect();
      if (index === this.index && this.image) { this.resetScroll(); return; }
      const request = ++this.sequence;
      this.index = index; this.image = null; this.scale = 1; this.fit = false;
      const sheet = this.group.sheets[index];
      this.updateTabs(); this.updateControls(); this.viewport.replaceChildren(); this.resetScroll();
      this.panel.setAttribute('aria-busy', 'true');
      this.status.textContent = `Loading ${sheet.label}…`;
      this.note.textContent = sheet.coverageNote || 'Read-only screenshot preview.';
      try {
        const image = await this.load(sheet);
        if (request !== this.sequence) return;
        this.image = image;
        this.viewport.replaceChildren(image);
        this.resize(); this.resetScroll();
        this.status.textContent = `${sheet.label} · Scroll to explore. Use Fit Width or zoom for a closer look.`;
      } catch {
        if (request !== this.sequence) return;
        const error = document.createElement('div'); error.className = 'ev-error';
        const text = document.createElement('p'); text.textContent = `Unable to load ${sheet.label}. Please retry.`;
        const retry = document.createElement('button'); retry.type = 'button'; retry.textContent = 'Retry image';
        retry.addEventListener('click', () => this.select(this.index));
        error.append(text, retry); this.viewport.replaceChildren(error);
        this.status.textContent = 'Screenshot unavailable. Other worksheet tabs are still available.';
      } finally { if (request === this.sequence) this.panel.setAttribute('aria-busy', 'false'); }
    }
    resize() {
      if (!this.image) return;
      if (this.fit) this.scale = Math.min(1, this.viewport.clientWidth / this.image.naturalWidth);
      this.image.style.width = `${this.image.naturalWidth * this.scale}px`;
      this.image.style.height = `${this.image.naturalHeight * this.scale}px`;
      this.root.querySelector('.ev-scale').textContent = `${Math.round(this.scale * 100)}%`;
      this.updateControls();
    }
    updateControls() {
      for (const action of ['out', 'in', 'reset', 'fit']) {
        this.root.querySelector(`[data-action="${action}"]`).disabled = !this.image || (action === 'out' && this.scale <= .25) || (action === 'in' && this.scale >= 3);
      }
    }
    resetScroll() { this.viewport.scrollTop = 0; this.viewport.scrollLeft = 0; }
    expand() {
      if (this.dialog?.open) return;
      const expand = this.root.querySelector('[data-action="expand"]');
      const close = this.root.querySelector('[data-action="close"]');
      const marker = document.createComment('Viewer position');
      this.root.before(marker);
      const dialog = document.createElement('dialog'); this.dialog = dialog;
      dialog.className = 'excel-viewer-dialog'; dialog.setAttribute('aria-label', `${this.title} expanded workbook preview`);
      document.body.append(dialog); dialog.append(this.root);
      this.oldOverflow = document.body.style.overflow;
      document.body.style.overflow = 'hidden';
      expand.hidden = true; close.hidden = false;
      dialog.addEventListener('close', () => {
        marker.replaceWith(this.root); dialog.remove();
        expand.hidden = false; close.hidden = true;
        document.body.style.overflow = this.oldOverflow;
        this.resize(); expand.focus({ preventScroll:true });
      }, { once:true });
      dialog.showModal(); close.focus(); this.resize();
    }
  }
  globalThis.ExcelScreenshotViewer = ExcelScreenshotViewer;
})();
