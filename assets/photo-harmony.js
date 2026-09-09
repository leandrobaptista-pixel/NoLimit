(() => {
  const SELECTOR = [
    '.hero-symbol-panel img',
    '.category-feature img',
    '.gallery-media-frame img',
    '.gallery-lightbox-stage img',
    '.process-card-media img',
    '.service-card-media img',
    '.viewer-main > img'
  ].join(',');

  const clamp = (value, min, max) => Math.min(max, Math.max(min, value));

  function analyzeAndBalance(img) {
    if (!img || img.dataset.photoHarmonyReady === 'true') return;
    if (!img.complete || !img.naturalWidth || !img.naturalHeight) {
      img.addEventListener('load', () => analyzeAndBalance(img), { once: true });
      return;
    }

    try {
      const canvas = document.createElement('canvas');
      canvas.width = 32;
      canvas.height = 32;
      const context = canvas.getContext('2d', { willReadFrequently: true });
      context.drawImage(img, 0, 0, canvas.width, canvas.height);
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;

      let luminanceTotal = 0;
      let saturationTotal = 0;
      let samples = 0;

      for (let offset = 0; offset < pixels.length; offset += 4) {
        if (pixels[offset + 3] < 180) continue;
        const red = pixels[offset] / 255;
        const green = pixels[offset + 1] / 255;
        const blue = pixels[offset + 2] / 255;
        const maximum = Math.max(red, green, blue);
        const minimum = Math.min(red, green, blue);
        luminanceTotal += (0.2126 * red) + (0.7152 * green) + (0.0722 * blue);
        saturationTotal += maximum ? (maximum - minimum) / maximum : 0;
        samples += 1;
      }

      if (!samples) return;
      const averageLuminance = luminanceTotal / samples;
      const averageSaturation = saturationTotal / samples;

      // Lift darker photos and gently restrain very bright ones.
      const brightness = clamp(0.68 / Math.max(averageLuminance, 0.08), 0.90, 1.28);
      // Colorful or tinted photos receive more desaturation than neutral ones.
      const saturation = clamp(0.105 / Math.max(averageSaturation, 0.08), 0.48, 0.86);

      img.style.setProperty('--photo-brightness', brightness.toFixed(3));
      img.style.setProperty('--photo-saturation', saturation.toFixed(3));
      img.dataset.photoHarmonyReady = 'true';
    } catch (_error) {
      // If an image cannot be sampled, the conservative CSS fallback remains.
      img.dataset.photoHarmonyReady = 'true';
    }
  }

  function balanceAll(root = document) {
    if (root.matches?.(SELECTOR)) analyzeAndBalance(root);
    root.querySelectorAll?.(SELECTOR).forEach(analyzeAndBalance);
  }

  balanceAll();
  new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) balanceAll(node);
      });
    });
  }).observe(document.body, { childList: true, subtree: true });
})();
