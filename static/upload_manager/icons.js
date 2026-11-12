// Share inventory art with the toolbox; keep compact control glyphs local.
(() => {
  const sprites = new Map(Object.entries({
    search: 'search', folder: 'folder', 'folder-closed': 'folder', 'folder-search': 'folder',
    download: 'download', upload: 'upload', 'layout-dashboard': 'grid', 'layers-2': 'layers',
    'chart-no-axes-column-increasing': 'chart', 'chart-pie': 'chart', files: 'layers',
    'circle-check': 'check', 'clipboard-copy': 'document', history: 'layers', sparkles: 'compass',
    'table-import': 'download', 'table-export': 'upload',
    'file-json-import': 'download', 'file-json-export': 'upload',
  }));
  const paths = {
    plus: '<path d="M12 5v14M5 12h14"/>',
    x: '<path d="m6 6 12 12M6 18 18 6"/>',
    check: '<path d="m5 12 4 4L19 6"/>',
    'contact-round': '<rect x="4" y="3" width="16" height="18" rx="2"/><circle cx="12" cy="9" r="3"/><path d="M7 18a5 5 0 0 1 10 0M2 7h4M2 12h4M2 17h4"/>',
    'chevron-right': '<path d="m9 5 7 7-7 7"/>',
    'chevron-down': '<path d="m5 9 7 7 7-7"/>',
    'chevron-left': '<path d="m15 5-7 7 7 7"/>',
    'chevron-up': '<path d="m5 15 7-7 7 7"/>',
    'arrow-up-down': '<path d="M8 3v17m-4-4 4 4 4-4M16 21V4m-4 4 4-4 4 4"/>',
    'calendar-days': '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4M17 3v4M3 11h18M7 15h2m4 0h2M7 18h2"/>',
    'trash-2': '<path d="M3 6h18M9 6V3h6v3M6 6l1 15h10l1-15M10 10v7M14 10v7"/>',
    'pencil': '<path d="m15 4 5 5M4 20l2-6L17 3l4 4L10 18Z"/>',
    'refresh-cw': '<path d="M20 10a8 8 0 1 0-1 8M20 3v7h-7"/>',
    'hard-drive': '<rect x="3" y="13" width="18" height="8" rx="2"/><path d="m3 13 4-9h10l4 9M16 17h2"/>'
  };
  window.lucide = { createIcons() {
    document.querySelectorAll('i[data-lucide]').forEach(icon => {
      const name = icon.dataset.lucide;
      const sprite = sprites.get(name);
      if (sprite) {
        icon.removeAttribute('data-lucide');
        icon.classList.add('manager-icon', 'harbor-icon', `icon-${sprite}`);
        icon.setAttribute('aria-hidden', 'true');
        return;
      }
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      for (const [key, value] of Object.entries({viewBox:'0 0 24 24', width:'24', height:'24', fill:'none', stroke:'currentColor', 'stroke-width':'2', 'stroke-linecap':'round', 'stroke-linejoin':'round', 'aria-hidden':'true'})) svg.setAttribute(key, value);
      svg.setAttribute('class', 'lucide lucide-' + name + ' ' + (icon.className || ''));
      svg.innerHTML = paths[name] || '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/>';
      icon.replaceWith(svg);
    });
  }};
})();
