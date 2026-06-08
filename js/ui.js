/**
 * ui.js — Componentes reutilizáveis de UI.
 *
 * Sidebar, topbar, loader, modal com focus trap, LGPD,
 * busca global de alunos, modo apresentação.
 * NÃO conhece regras de negócio — apenas manipula DOM.
 */

// ---------------------------------------------------------------------------
// Sidebar Navigation
// ---------------------------------------------------------------------------

let _onNavigate = null;

export function initSidebar(onNavigate) {
  _onNavigate = onNavigate;

  const links = document.querySelectorAll('.sidebar-nav a[data-section]');

  for (const link of links) {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const name = link.getAttribute('data-section');
      if (name) navigateTo(name);
    });
  }

  // Mobile toggle
  const toggle = document.getElementById('menu-toggle');
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebar-overlay');

  if (toggle && sidebar) {
    toggle.addEventListener('click', () => {
      const isOpen = sidebar.classList.toggle('open');
      toggle.setAttribute('aria-expanded', String(isOpen));
      if (isOpen) {
        if (overlay) overlay.classList.add('visible');
      } else {
        if (overlay) overlay.classList.remove('visible');
      }
    });
  }

  if (overlay) {
    overlay.addEventListener('click', closeSidebar);
  }
}

function closeSidebar() {
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebar-overlay');
  const toggle = document.getElementById('menu-toggle');
  if (sidebar) sidebar.classList.remove('open');
  if (overlay) overlay.classList.remove('visible');
  if (toggle) toggle.setAttribute('aria-expanded', 'false');
}

// ---------------------------------------------------------------------------
// Section Navigation
// ---------------------------------------------------------------------------

export function navigateTo(name) {
  // Hide all sections
  const sections = document.querySelectorAll('.section[data-section]');
  for (const sec of sections) sec.classList.remove('active');

  // Show target
  const target = document.querySelector(`.section[data-section="${name}"]`);
  if (target) target.classList.add('active');

  // Update sidebar active
  const links = document.querySelectorAll('.sidebar-nav a[data-section]');
  for (const link of links) link.classList.remove('active');
  const activeLink = document.querySelector(`.sidebar-nav a[data-section="${name}"]`);
  if (activeLink) activeLink.classList.add('active');

  // Update topbar
  if (activeLink) {
    setTopbarTitle(activeLink.textContent.trim());
  }

  // Close mobile sidebar
  closeSidebar();

  // Callback
  if (_onNavigate) _onNavigate(name);
}

export function getCurrentSection() {
  const active = document.querySelector('.section.active');
  return active ? active.getAttribute('data-section') : null;
}

// ---------------------------------------------------------------------------
// Topbar
// ---------------------------------------------------------------------------

export function setTopbarTitle(title) {
  const el = document.getElementById('topbar-title');
  if (el) el.textContent = title;
}

export function setTopbarStatus(text) {
  const el = document.getElementById('topbar-status');
  if (el) el.textContent = text;
}

// ---------------------------------------------------------------------------
// Loader
// ---------------------------------------------------------------------------

export function showLoader(msg = 'Carregando...') {
  const overlay = document.getElementById('loader');
  const textEl = document.getElementById('loader-text');
  if (textEl) textEl.textContent = msg;
  if (overlay) overlay.classList.add('active');
}

export function hideLoader() {
  const overlay = document.getElementById('loader');
  if (overlay) overlay.classList.remove('active');
}

// ---------------------------------------------------------------------------
// Modal com Focus Trap
// ---------------------------------------------------------------------------

let _lastFocus = null;
let _modalCloseCb = null;

export function showModal(overlayId, bodyHtml) {
  const overlay = document.getElementById(overlayId);
  if (!overlay) return;

  // Save last focus
  _lastFocus = document.activeElement;

  // Set body content
  const body = overlay.querySelector('.modal-body');
  if (body && bodyHtml !== undefined) {
    body.innerHTML = bodyHtml;
  }

  // Show modal
  overlay.classList.add('open');
  overlay.setAttribute('aria-hidden', 'false');

  // Focus first focusable element
  const modal = overlay.querySelector('.modal');
  if (modal) {
    const focusable = modal.querySelectorAll(
      'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
    );
    if (focusable.length > 0) {
      setTimeout(() => focusable[0].focus(), 100);
    }
  }

  // Trap focus
  overlay.addEventListener('keydown', _trapFocus);
}

export function onModalClose(overlayId, callback) {
  const overlay = document.getElementById(overlayId);
  if (!overlay) return;

  _modalCloseCb = callback;
}

export function closeModal(overlayId) {
  const overlay = document.getElementById(overlayId);
  if (!overlay) return;

  overlay.classList.remove('open');
  overlay.setAttribute('aria-hidden', 'true');
  overlay.removeEventListener('keydown', _trapFocus);

  // Return focus
  if (_lastFocus) {
    setTimeout(() => _lastFocus.focus(), 100);
    _lastFocus = null;
  }
}

function _trapFocus(e) {
  if (e.key !== 'Tab') return;

  const overlay = e.currentTarget;
  const modal = overlay.querySelector('.modal');
  if (!modal) return;

  const focusable = modal.querySelectorAll(
    'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
  );
  if (focusable.length === 0) return;

  const first = focusable[0];
  const last = focusable[focusable.length - 1];

  if (e.shiftKey) {
    if (document.activeElement === first) {
      e.preventDefault();
      last.focus();
    }
  } else {
    if (document.activeElement === last) {
      e.preventDefault();
      first.focus();
    }
  }
}

export function initModalCloseButtons(overlayId) {
  const overlay = document.getElementById(overlayId);
  if (!overlay) return;

  // Close on overlay click
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) {
      closeModal(overlayId);
    }
  });

  // Close on ESC
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && overlay.classList.contains('open')) {
      closeModal(overlayId);
    }
  });

  // Close buttons
  const closeBtns = overlay.querySelectorAll('.modal-close, [id$="-close"]');
  for (const btn of closeBtns) {
    btn.addEventListener('click', () => closeModal(overlayId));
  }
}

// ---------------------------------------------------------------------------
// LGPD Banner
// ---------------------------------------------------------------------------

export function initLgpd() {
  const banner = document.getElementById('lgpd-banner');
  const acceptBtn = document.getElementById('lgpd-accept');
  if (!banner || !acceptBtn) return;

  const dismissed = localStorage.getItem('cc_lgpd_dismissed');
  if (dismissed) return;

  setTimeout(() => banner.classList.add('visible'), 800);

  acceptBtn.addEventListener('click', () => {
    banner.classList.remove('visible');
    localStorage.setItem('cc_lgpd_dismissed', '1');
  });
}

// ---------------------------------------------------------------------------
// Alert messages
// ---------------------------------------------------------------------------

export function showAlert(containerId, type, message) {
  const container = document.getElementById(containerId);
  if (!container) return;

  const icons = {
    info: 'bi-info-circle-fill',
    warning: 'bi-exclamation-triangle-fill',
    error: 'bi-x-circle-fill',
    success: 'bi-check-circle-fill',
  };

  container.innerHTML = `
    <div class="alert alert-${type}" role="alert">
      <i class="bi ${icons[type] || icons.info}"></i>
      <span>${message}</span>
    </div>
  `;
}

// ---------------------------------------------------------------------------
// Empty state
// ---------------------------------------------------------------------------

export function showEmpty(containerId, icon, title, message) {
  const container = document.getElementById(containerId);
  if (!container) return;

  container.innerHTML = `
    <div class="empty-state">
      <div class="empty-state-icon"><i class="bi bi-${icon}"></i></div>
      <h3>${title}</h3>
      <p>${message}</p>
    </div>
  `;
}

// ---------------------------------------------------------------------------
// Busca global de alunos (filtra tabelas na página ativa)
// ---------------------------------------------------------------------------

export function initSearch() {
  const input = document.getElementById('search-aluno');
  if (!input) return;

  input.addEventListener('input', () => {
    const query = input.value.trim().toLowerCase();
    filterTables(query);
  });
}

function filterTables(query) {
  const section = document.querySelector('.section.active');
  if (!section) return;

  const rows = section.querySelectorAll('table tbody tr');
  for (const row of rows) {
    // Buscar no texto da primeira célula (nome do aluno)
    const firstCell = row.querySelector('td:first-child');
    if (!firstCell) continue;
    const text = firstCell.textContent.toLowerCase();

    if (!query || text.includes(query)) {
      row.style.display = '';
    } else {
      row.style.display = 'none';
    }
  }
}

export function clearSearch() {
  const input = document.getElementById('search-aluno');
  if (input) {
    input.value = '';
    filterTables('');
  }
}

// ---------------------------------------------------------------------------
// Modo apresentação (tela cheia)
// ---------------------------------------------------------------------------

let _fullscreen = false;

export function initFullscreen() {
  const btn = document.getElementById('btn-fullscreen');
  if (!btn) return;

  btn.addEventListener('click', () => {
    _fullscreen = !_fullscreen;
    const sidebar = document.getElementById('sidebar');
    const topbar = document.querySelector('.topbar');
    const content = document.querySelector('.content');
    const main = document.querySelector('.main');

    if (_fullscreen) {
      if (sidebar) sidebar.style.display = 'none';
      if (topbar) topbar.style.display = 'none';
      if (content) content.style.top = '0';
      if (main) main.style.left = '0';
      btn.innerHTML = '<i class="bi bi-fullscreen-exit"></i>';
      btn.title = 'Sair do modo apresentação';
    } else {
      if (sidebar) sidebar.style.display = '';
      if (topbar) topbar.style.display = '';
      if (content) content.style.top = '';
      if (main) main.style.left = '';
      btn.innerHTML = '<i class="bi bi-arrows-fullscreen"></i>';
      btn.title = 'Modo apresentação';
    }
  });
}
