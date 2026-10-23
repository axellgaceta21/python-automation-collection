/* Shared footer: include footer.css and this deferred script on any page.
 * An optional <footer data-shared-footer> supplies a no-JavaScript fallback.
 * Without a mount, a footer is appended automatically. No network fetch required.
 */
(() => {
  const links = [
    { label: 'GitHub', url: 'https://github.com/axellgaceta21', icon: '<path d="M9 19c-4 1-4-2-6-2m12 4v-3.5c0-1 .1-1.5-.5-2 3.3-.4 6.5-1.6 6.5-7A5.4 5.4 0 0 0 19.5 5 5 5 0 0 0 19.4 1S18.1.6 15 2.5a13.4 13.4 0 0 0-6 0C5.9.6 4.6 1 4.6 1A5 5 0 0 0 4.5 5 5.4 5.4 0 0 0 3 8.5c0 5.4 3.2 6.6 6.5 7-.6.5-.6 1.3-.5 2V21"/>' },
    { label: 'Portfolio', url: 'https://axell-builds.lovable.app/', primary: true, icon: '<rect x="3" y="6" width="18" height="15" rx="2"/><path d="M8 6V3h8v3M3 11h18M10 11v3h4v-3"/>' },
    { label: 'Resume', url: 'https://drive.google.com/file/d/1KCkvbanLu4UV6E7QBP6FY8yVaKPkto5I/view?usp=sharing', icon: '<path d="M14 2H5v20h14V7zM14 2v5h5M8 12h8M8 16h8"/>' },
    { label: 'LinkedIn', url: 'https://www.linkedin.com/in/axellgaceta21/', icon: '<rect x="2" y="2" width="20" height="20" rx="3"/><path d="M7 10v7M7 7v.1M11 17v-7M11 13c0-4 6-4 6 0v4"/>' }
  ];
  function mountFooter() {
    let mounts = document.querySelectorAll('[data-shared-footer]');
    if (!mounts.length) {
      const footer = document.createElement('footer');
      footer.dataset.sharedFooter = '';
      document.body.append(footer);
      mounts = [footer];
    }
    mounts.forEach(footer => {
      if (footer.dataset.footerReady === 'true') return;
      footer.classList.add('shared-footer');
      footer.setAttribute('aria-label', 'About Axell Gaceta and external links');
      footer.innerHTML = `<div class="sf-content">
        <a class="sf-logo-link" href="/" aria-label="Python Automation Portfolio home"><img class="sf-logo" src="/assets/images/logo.png" alt="Axell Gaceta logo" width="88" height="81" loading="lazy" decoding="async"></a>
        <p class="sf-label">Built &amp; documented by</p>
        <h2 class="sf-name">Axell Gaceta</h2>
        <p class="sf-subtitle">Web Developer · AI Automation Builder</p>
        <nav class="sf-links" aria-label="Axell Gaceta external profiles and resume">
          ${links.map(link => `<a class="sf-link${link.primary ? ' sf-link-primary' : ''}" href="${link.url}" target="_blank" rel="noopener noreferrer"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false">${link.icon}</svg><span>${link.label}</span><span class="sf-new-tab" aria-hidden="true">↗</span><span class="sf-sr-only"> (opens in a new tab)</span></a>`).join('')}
        </nav>
        <div class="sf-bottom"><p>© ${new Date().getFullYear()} Axell Gaceta. All rights reserved.</p></div>
      </div>`;
      footer.dataset.footerReady = 'true';
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mountFooter, { once: true });
  else mountFooter();
})();
