function formatDate(value) {
  if (!value) return '-';
  return new Intl.DateTimeFormat('en-US', { dateStyle: 'medium' }).format(new Date(value));
}

function StatusPill({ active, children }) {
  return <span className={`status-pill ${active ? 'is-active' : ''}`}>{children}</span>;
}

export default function MediaTable({
  items,
  selectedId,
  onSelect,
  onGenerate,
  onTogglePublished,
  onToggleSocial,
  busyId
}) {
  return (
    <section className="panel">
      <div className="panel-heading compact">
        <div>
          <p className="eyebrow">Library</p>
          <h2>Uploaded media</h2>
        </div>
        <span className="chip chip-muted">{items.length} items</span>
      </div>

      <div className="table-shell">
        <table className="media-table">
          <thead>
            <tr>
              <th>Preview</th>
              <th>Title</th>
              <th>Category</th>
              <th>Created</th>
              <th>Published</th>
              <th>Artwork</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {items.length ? (
              items.map((item) => {
                const active = item.id === selectedId;
                const processing = busyId === item.id;
                return (
                  <tr
                    key={item.id}
                    className={active ? 'is-selected' : ''}
                    onClick={() => onSelect(item.id)}
                  >
                    <td>
                      <img className="thumb" src={item.originalUrl} alt={item.title} />
                    </td>
                    <td>
                      <strong>{item.title}</strong>
                      <span className="cell-subtitle">{item.usedInSocial ? 'Used in social' : 'Ready for social'}</span>
                    </td>
                    <td>{item.categoryName}</td>
                    <td>{formatDate(item.createdAt)}</td>
                    <td>
                      <StatusPill active={item.published}>
                        {item.published ? 'Published' : 'Draft'}
                      </StatusPill>
                    </td>
                    <td>
                      <StatusPill active={Boolean(item.generatedArtUrl)}>
                        {item.generatedArtUrl ? 'Generated' : 'Pending'}
                      </StatusPill>
                    </td>
                    <td>
                      <div className="row-actions">
                        <button
                          className="button ghost small"
                          type="button"
                          onClick={(event) => {
                            event.stopPropagation();
                            onGenerate(item.id);
                          }}
                          disabled={processing}
                        >
                          {processing ? 'Working...' : 'Generate art'}
                        </button>
                        <button
                          className="button ghost small"
                          type="button"
                          onClick={(event) => {
                            event.stopPropagation();
                            onTogglePublished(item);
                          }}
                        >
                          {item.published ? 'Unpublish' : 'Publish'}
                        </button>
                        <button
                          className="button ghost small"
                          type="button"
                          onClick={(event) => {
                            event.stopPropagation();
                            onToggleSocial(item);
                          }}
                        >
                          {item.usedInSocial ? 'Mark unused' : 'Mark social'}
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            ) : (
              <tr>
                <td colSpan="7" className="empty-cell">
                  No media items match the current filters.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}
