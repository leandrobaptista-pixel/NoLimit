(() => {
  const heroImage = document.getElementById('heroProjectImage');
  const heroCaption = document.getElementById('heroProjectCaption');
  const galleries = window.GALLERY_MANIFEST || {};
  const variants = window.OPTIMIZED_IMAGE_VARIANTS || {};

  if (heroImage && galleries && typeof galleries === 'object') {
    const projects = Object.entries(galleries).flatMap(([category, paths]) =>
      (Array.isArray(paths) ? paths : []).map((path) => ({ category, path }))
    );
    try {
      const lastPath = localStorage.getItem('noLimitHeroLastProject');
      const choices = projects.filter((project) => project.path !== lastPath);
      const selected = (choices.length ? choices : projects)[Math.floor(Math.random() * (choices.length || 1))];

      if (selected) {
        const optimized = variants[selected.path] || {};
        heroImage.src = optimized.feature || optimized.thumb || selected.path;
        heroImage.alt = `No Limit Carpentry ${selected.category} project`;
        heroCaption.textContent = selected.category;
        localStorage.setItem('noLimitHeroLastProject', selected.path);
      }
    } catch {
      // Keep the static fallback image when browser storage is unavailable.
    }
  }

  const workCards = [...document.querySelectorAll('[data-selected-work-card]')];
  const galleryCategories = Object.entries(galleries).filter(([, paths]) => Array.isArray(paths) && paths.length);

  if (workCards.length && galleryCategories.length) {
    try {
      const lastCategories = JSON.parse(localStorage.getItem('noLimitSelectedWorkLastCategories') || '[]');
      const eligible = galleryCategories.filter(([category]) => !lastCategories.includes(category));
      const categoryPool = eligible.length >= workCards.length ? eligible : galleryCategories;
      const selectedCategories = [...categoryPool]
        .sort(() => Math.random() - 0.5)
        .slice(0, workCards.length);

      workCards.forEach((card, index) => {
        const [category, projects] = selectedCategories[index];
        const selected = projects[Math.floor(Math.random() * projects.length)];
        const image = card.querySelector('img');
        const label = card.querySelector('span');
        const optimized = variants[selected] || {};

        card.href = `?cat=${encodeURIComponent(category)}#gallery`;
        image.src = optimized.thumb || optimized.feature || selected;
        image.alt = `No Limit Carpentry ${category} project`;
        if (label?.firstChild) label.firstChild.textContent = `${category} `;
      });

      localStorage.setItem('noLimitSelectedWorkLastCategories', JSON.stringify(selectedCategories.map(([category]) => category)));
    } catch {
      // Keep the authored fallback cards when browser storage is unavailable.
    }
  }

  const splash = document.getElementById('entrySplash');
  if (!splash || window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;

  const key = 'noLimitEntrySplashSeen';
  const seen = sessionStorage.getItem(key);
  if (seen) {
    splash.classList.add('is-hidden');
    return;
  }

  window.setTimeout(() => splash.classList.add('is-hidden'), 2400);
  sessionStorage.setItem(key, 'true');
})();
