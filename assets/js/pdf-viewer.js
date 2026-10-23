/* Native PDF preview. Only the selected document is checked and embedded. */
(() => {
  class NativePdfViewer {
    constructor(host, group, title) {
      this.documents = group.documents || [];
      this.sequence = 0;
      this.root = document.createElement('div');
      this.root.className = 'pdf-preview';
      this.root.innerHTML = `<div class="pv-toolbar"><label class="pv-select-label">Sample document<select class="pv-select" aria-label="Choose a sample PDF document"></select></label><a class="pv-open" target="_blank" rel="noopener noreferrer" hidden>Open PDF in New Tab <span aria-hidden="true"> ↗</span></a></div><div class="pv-stage" aria-label="Selected PDF preview"></div><p class="pv-status" role="status" aria-live="polite"></p><p class="pv-fallback">Use your browser's PDF controls to scroll or zoom. If inline viewing is unavailable, open the selected PDF in a new tab.</p>`;
      host.replaceChildren(this.root);
      this.select = this.root.querySelector('.pv-select');
      this.stage = this.root.querySelector('.pv-stage');
      this.status = this.root.querySelector('.pv-status');
      this.open = this.root.querySelector('.pv-open');
      this.title = title;
      this.documents.forEach(doc => {
        const option = document.createElement('option');
        option.value = doc.filename; option.textContent = doc.label;
        this.select.append(option);
      });
      this.select.value = this.documents.some(doc => doc.filename === group.defaultDocument) ? group.defaultDocument : this.documents[0]?.filename || '';
      this.select.addEventListener('change', () => this.load());
      this.status.textContent = this.documents.length ? 'Select a document or scroll into view to load its preview.' : 'No approved PDF documents are available.';
      if (!this.documents.length) { this.select.disabled = true; return; }
      if ('IntersectionObserver' in window) {
        this.lazy = new IntersectionObserver(entries => {
          if (entries.some(entry => entry.isIntersecting)) { this.lazy.disconnect(); this.load(); }
        }, { rootMargin:'100px' });
        this.lazy.observe(this.root);
      } else this.load();
    }
    async load() {
      this.lazy?.disconnect();
      const doc = this.documents.find(item => item.filename === this.select.value);
      if (!doc) return;
      this.controller?.abort();
      this.controller = new AbortController();
      const request = ++this.sequence;
      clearTimeout(this.timer);
      this.stage.replaceChildren();
      this.stage.setAttribute('aria-busy', 'true');
      this.open.href = doc.path; this.open.hidden = false;
      this.open.setAttribute('aria-label', `Open ${doc.label} PDF in a new tab`);
      this.status.textContent = `Loading ${doc.label}…`;
      try {
        // iframe load events alone cannot reliably detect a missing HTTP resource.
        const response = await fetch(doc.path, { method:'HEAD', signal:this.controller.signal });
        if (!response.ok) throw new Error('Document unavailable');
        if (request !== this.sequence) return;
        const frame = document.createElement('iframe');
        frame.title = `${this.title}: ${doc.label}`;
        frame.addEventListener('load', () => {
          if (request !== this.sequence) return;
          clearTimeout(this.timer);
          this.stage.setAttribute('aria-busy', 'false');
          this.status.textContent = `${doc.label} selected. If the preview does not appear, use Open PDF in New Tab.`;
        });
        frame.addEventListener('error', () => { if (request === this.sequence) this.fail(doc); });
        this.timer = setTimeout(() => {
          if (request !== this.sequence) return;
          this.stage.setAttribute('aria-busy', 'false');
          this.status.textContent = 'The embedded preview is taking longer than expected. You can open the selected PDF in a new tab.';
        }, 12000);
        frame.src = doc.path;
        this.stage.append(frame);
      } catch (error) {
        if (request === this.sequence && error.name !== 'AbortError') this.fail(doc);
      }
    }
    fail(doc) {
      clearTimeout(this.timer);
      this.stage.setAttribute('aria-busy', 'false');
      this.stage.replaceChildren();
      const box = document.createElement('div'); box.className = 'pv-error';
      box.textContent = `The ${doc.label} preview is unavailable. Try again or use Open PDF in New Tab.`;
      const retry = document.createElement('button'); retry.type = 'button'; retry.textContent = 'Retry PDF';
      retry.addEventListener('click', () => this.load()); box.append(retry); this.stage.append(box);
      this.status.textContent = 'PDF preview unavailable.';
    }
  }
  globalThis.NativePdfViewer = NativePdfViewer;
})();
