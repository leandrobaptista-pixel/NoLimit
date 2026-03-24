export default function FilterBar({ filters, categories, onChange, onRefresh, loading }) {
  return (
    <section className="panel">
      <div className="panel-heading compact">
        <div>
          <p className="eyebrow">Find Content</p>
          <h2>Filters</h2>
        </div>
        <button className="button ghost" type="button" onClick={onRefresh} disabled={loading}>
          {loading ? 'Refreshing...' : 'Refresh'}
        </button>
      </div>

      <div className="filter-grid">
        <label className="field">
          <span>Search title</span>
          <input
            type="search"
            value={filters.search}
            onChange={(event) => onChange({ search: event.target.value })}
            placeholder="Search uploaded work"
          />
        </label>

        <label className="field">
          <span>Category</span>
          <select
            value={filters.category}
            onChange={(event) => onChange({ category: event.target.value })}
          >
            <option value="">All categories</option>
            {categories.map((category) => (
              <option key={category.id} value={category.slug}>
                {category.name}
              </option>
            ))}
          </select>
        </label>

        <label className="field">
          <span>Publish status</span>
          <select
            value={filters.published}
            onChange={(event) => onChange({ published: event.target.value })}
          >
            <option value="all">All media</option>
            <option value="true">Published only</option>
            <option value="false">Draft only</option>
          </select>
        </label>
      </div>
    </section>
  );
}
