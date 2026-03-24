function formatDate(value) {
  if (!value) return 'Not available';
  return new Intl.DateTimeFormat('en-US', {
    dateStyle: 'medium',
    timeStyle: 'short'
  }).format(new Date(value));
}

export default function PreviewPanel({ item }) {
  if (!item) {
    return (
      <section className="panel preview-panel">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">Preview</p>
            <h2>No item selected</h2>
          </div>
        </div>
        <p className="muted">Upload a media file or select one from the table to see details.</p>
      </section>
    );
  }

  return (
    <section className="panel preview-panel">
      <div className="panel-heading">
        <div>
          <p className="eyebrow">Selected Media</p>
          <h2>{item.title}</h2>
        </div>
        <span className={`chip ${item.published ? 'chip-success' : 'chip-muted'}`}>
          {item.published ? 'Published' : 'Draft'}
        </span>
      </div>

      <div className="preview-grid">
        <article className="preview-card">
          <p className="preview-label">Original</p>
          <img src={item.originalUrl} alt={item.title} />
        </article>

        <article className="preview-card">
          <p className="preview-label">Promotional artwork</p>
          {item.generatedArtUrl ? (
            <img src={item.generatedArtUrl} alt={`${item.title} promotional artwork`} />
          ) : (
            <div className="preview-placeholder">Generate artwork to preview the branded square asset.</div>
          )}
        </article>
      </div>

      <dl className="meta-grid">
        <div>
          <dt>Category</dt>
          <dd>{item.categoryName}</dd>
        </div>
        <div>
          <dt>Created</dt>
          <dd>{formatDate(item.createdAt)}</dd>
        </div>
        <div>
          <dt>Social status</dt>
          <dd>{item.usedInSocial ? 'Used in social' : 'Not used yet'}</dd>
        </div>
      </dl>

      <label className="field">
        <span>Generated caption</span>
        <textarea readOnly rows="5" value={item.captionText || 'Caption will appear here.'} />
      </label>
    </section>
  );
}
