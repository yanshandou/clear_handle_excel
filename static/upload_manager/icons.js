// Small local SVG set: no CDN or package install is needed to display the manager.
(() => {
  const paths = {
    plus: '<path d="M12 5v14M5 12h14"/>',
    x: '<path d="m6 6 12 12M6 18 18 6"/>',
    check: '<path d="m5 12 4 4L19 6"/>',
    search: '<circle cx="10" cy="10" r="6"/><path d="m15 15 5 5"/>',
    folder: '<path d="M3 7V5h6l2 2h10v13H3Z"/>',
    download: '<path d="M12 3v12m-5-5 5 5 5-5M4 17v4h16v-4"/>',
    upload: '<path d="M12 16V4m-5 5 5-5 5 5M4 17v4h16v-4"/>',
    'table-import': '<rect x="4" y="3" width="17" height="18" rx="2"/><path d="M4 9h17M12 3v18M12 15h9M1 15h7m-3-3 3 3-3 3"/>',
    'table-export': '<rect x="3" y="3" width="17" height="18" rx="2"/><path d="M3 9h17M11 3v18M3 15h8M16 15h7m-3-3 3 3-3 3"/>',
    'contact-round': '<rect x="4" y="3" width="16" height="18" rx="2"/><circle cx="12" cy="9" r="3"/><path d="M7 18a5 5 0 0 1 10 0M2 7h4M2 12h4M2 17h4"/>',
    'file-json-import': '<path d="M11 2H3v20h13V7l-5-5v5h5M7 11H6v3H5l1 1v3h1M12 11h1v3h1l-1 1v3h-1M20 10v10m-3-3 3 3 3-3"/>',
    'file-json-export': '<path d="M11 2H3v20h13V7l-5-5v5h5M7 11H6v3H5l1 1v3h1M12 11h1v3h1l-1 1v3h-1M20 20V10m-3 3 3-3 3 3"/>',
    'chevron-right': '<path d="m9 5 7 7-7 7"/>',
    'chevron-down': '<path d="m5 9 7 7 7-7"/>',
    'chevron-left': '<path d="m15 5-7 7 7 7"/>',
    'chevron-up': '<path d="m5 15 7-7 7 7"/>',
    'arrow-up-down': '<path d="M8 3v17m-4-4 4 4 4-4M16 21V4m-4 4 4-4 4 4"/>',
    'calendar-days': '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4M17 3v4M3 11h18M7 15h2m4 0h2M7 18h2"/>',
    'clipboard-copy': '<rect x="5" y="5" width="14" height="17" rx="2"/><rect x="9" y="2" width="6" height="5" rx="1"/>',
    'trash-2': '<path d="M3 6h18M9 6V3h6v3M6 6l1 15h10l1-15M10 10v7M14 10v7"/>',
    'pencil': '<path d="m15 4 5 5M4 20l2-6L17 3l4 4L10 18Z"/>',
    'refresh-cw': '<path d="M20 10a8 8 0 1 0-1 8M20 3v7h-7"/>',
    'hard-drive': '<rect x="3" y="13" width="18" height="8" rx="2"/><path d="m3 13 4-9h10l4 9M16 17h2"/>',
    'layout-dashboard': '<path d="M3 3h8v8H3ZM15 3h6v6h-6ZM3 15h6v6H3ZM13 13h8v8h-8Z"/>',
    'layers-2': '<path d="m12 3 10 6-10 6L2 9ZM2 15l10 6 10-6"/>'
  };
  window.lucide = { createIcons() {
    document.querySelectorAll('i[data-lucide]').forEach(icon => {
      const name = icon.dataset.lucide;
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      for (const [key, value] of Object.entries({viewBox:'0 0 24 24', width:'24', height:'24', fill:'none', stroke:'currentColor', 'stroke-width':'2', 'stroke-linecap':'round', 'stroke-linejoin':'round', 'aria-hidden':'true'})) svg.setAttribute(key, value);
      svg.setAttribute('class', 'lucide lucide-' + name + ' ' + (icon.className || ''));
      svg.innerHTML = paths[name] || (name.includes('folder') ? paths.folder : '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/>');
      icon.replaceWith(svg);
    });
  }};
})();
