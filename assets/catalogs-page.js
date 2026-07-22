const catalogLibraryRoot = document.getElementById('catalogLibraryGroups');
const catalogSearchInput = document.getElementById('catalogSearch');
const catalogResultsCount = document.getElementById('catalogResultsCount');
const catalogYear = document.getElementById('catalogYear');
const navWrap = document.querySelector('.nav');
const navToggle = document.querySelector('.menu-toggle');

if (catalogYear) catalogYear.textContent = new Date().getFullYear();

function getCatalogLibraryGroups() {
  return Array.isArray(window.CATALOG_LIBRARY) ? window.CATALOG_LIBRARY : [];
}

function includesQuery(value, query) {
  return String(value || '').toLowerCase().includes(query);
}

function filterCatalogGroups(groups, query) {
  if (!query) return groups;

  return groups
    .map((group) => {
      const groupMatches =
        includesQuery(group.title, query) ||
        includesQuery(group.description, query) ||
        includesQuery(group.mark, query);

      const documents = (Array.isArray(group.documents) ? group.documents : []).filter((doc) => {
        return (
          includesQuery(doc.label, query) ||
          includesQuery(doc.description, query) ||
          includesQuery(doc.type, query)
        );
      });

      return {
        ...group,
        documents: groupMatches && !documents.length ? group.documents : documents
      };
    })
    .filter((group) => Array.isArray(group.documents) && group.documents.length);
}

function renderCatalogGroups() {
  if (!catalogLibraryRoot) return;

  const groups = getCatalogLibraryGroups();
  const query = String(catalogSearchInput?.value || '').trim().toLowerCase();
  const filteredGroups = filterCatalogGroups(groups, query);
  const totalDocuments = filteredGroups.reduce((sum, group) => {
    return sum + (Array.isArray(group.documents) ? group.documents.length : 0);
  }, 0);

  if (catalogResultsCount) {
    catalogResultsCount.textContent = query
      ? `${totalDocuments} matching files across ${filteredGroups.length} groups`
      : `${totalDocuments} files across ${filteredGroups.length} groups`;
  }

  if (!filteredGroups.length) {
    catalogLibraryRoot.innerHTML = '<p class="hint">No catalogs match this search yet. Try another keyword.</p>';
    return;
  }

  catalogLibraryRoot.innerHTML = '';

  filteredGroups.forEach((group) => {
    const section = document.createElement('section');
    section.className = 'catalog-group';
    section.id = group.id;

    const head = document.createElement('div');
    head.className = 'catalog-group-head';

    const copy = document.createElement('div');
    const title = document.createElement('h2');
    title.textContent = group.title;
    const description = document.createElement('p');
    description.textContent = group.description || '';
    copy.append(title, description);

    const total = document.createElement('span');
    total.className = 'catalog-group-total';
    total.textContent = `${Array.isArray(group.documents) ? group.documents.length : 0} files`;

    head.append(copy, total);

    const grid = document.createElement('div');
    grid.className = 'catalog-doc-grid';

    (Array.isArray(group.documents) ? group.documents : []).forEach((doc) => {
      const card = document.createElement('article');
      card.className = 'catalog-doc-card card';

      const top = document.createElement('div');
      top.className = 'catalog-doc-top';

      const titleWrap = document.createElement('div');
      const docTitle = document.createElement('h3');
      docTitle.textContent = doc.label;
      titleWrap.appendChild(docTitle);

      const docType = document.createElement('span');
      docType.className = 'catalog-doc-type';
      docType.textContent = doc.type || 'File';

      top.append(titleWrap, docType);

      const docDescription = document.createElement('p');
      docDescription.textContent = doc.description || 'Catalog document ready to open in a new tab.';

      const actions = document.createElement('div');
      actions.className = 'catalog-doc-actions';

      const open = document.createElement('a');
      open.className = 'btn';
      open.href = doc.href;
      open.target = '_blank';
      open.rel = 'noopener';
      open.textContent = 'Open PDF';

      const back = document.createElement('a');
      back.className = 'btn btn-secondary';
      back.href = './index.html#catalogs';
      back.textContent = 'Back to Homepage';

      actions.append(open, back);
      card.append(top, docDescription, actions);
      grid.appendChild(card);
    });

    section.append(head, grid);
    catalogLibraryRoot.appendChild(section);
  });
}

catalogSearchInput?.addEventListener('input', renderCatalogGroups);
renderCatalogGroups();

navToggle?.addEventListener('click', () => {
  const expanded = navToggle.getAttribute('aria-expanded') === 'true';
  navToggle.setAttribute('aria-expanded', String(!expanded));
  navWrap?.classList.toggle('nav-open', !expanded);
});

if (window.location.hash) {
  window.setTimeout(() => {
    const target = document.querySelector(window.location.hash);
    target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 0);
}
