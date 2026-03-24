import { useEffect, useMemo, useState } from 'react';
import FilterBar from './components/FilterBar.jsx';
import MediaTable from './components/MediaTable.jsx';
import PreviewPanel from './components/PreviewPanel.jsx';
import UploadForm from './components/UploadForm.jsx';
import { useAsyncData } from './hooks/useAsyncData.js';
import { generateArt, listCategories, listMedia, updateMediaStatus, uploadMedia } from './lib/api.js';

function StatCard({ label, value, note }) {
  return (
    <article className="stat-card">
      <span>{label}</span>
      <strong>{value}</strong>
      <p>{note}</p>
    </article>
  );
}

export default function App() {
  const categoriesState = useAsyncData([]);
  const mediaState = useAsyncData([]);
  const [filters, setFilters] = useState({
    search: '',
    category: '',
    published: 'all'
  });
  const [actionError, setActionError] = useState('');
  const [busyUpload, setBusyUpload] = useState(false);
  const [busyId, setBusyId] = useState('');
  const [selectedId, setSelectedId] = useState('');

  async function loadCategories() {
    const response = await categoriesState.run(() => listCategories());
    return response.categories || [];
  }

  async function loadMedia() {
    const response = await mediaState.run(() => listMedia(filters));
    return response.media || [];
  }

  useEffect(() => {
    loadCategories().catch(() => {});
  }, []);

  useEffect(() => {
    loadMedia().catch(() => {});
  }, [filters.category, filters.published, filters.search]);

  useEffect(() => {
    if (!mediaState.data.length) {
      setSelectedId('');
      return;
    }

    if (!mediaState.data.some((item) => item.id === selectedId)) {
      setSelectedId(mediaState.data[0].id);
    }
  }, [mediaState.data, selectedId]);

  const stats = useMemo(() => {
    const total = mediaState.data.length;
    const published = mediaState.data.filter((item) => item.published).length;
    const generated = mediaState.data.filter((item) => item.generatedArtUrl).length;
    return { total, published, generated };
  }, [mediaState.data]);

  const selectedItem = mediaState.data.find((item) => item.id === selectedId) || null;

  async function handleUpload(formData) {
    setBusyUpload(true);
    setActionError('');
    try {
      await uploadMedia(formData);
      await loadMedia();
    } catch (error) {
      setActionError(error.message || 'Upload failed.');
    } finally {
      setBusyUpload(false);
    }
  }

  async function handleGenerate(mediaId) {
    setBusyId(mediaId);
    setActionError('');
    try {
      await generateArt(mediaId);
      await loadMedia();
    } catch (error) {
      setActionError(error.message || 'Artwork generation failed.');
    } finally {
      setBusyId('');
    }
  }

  async function handleTogglePublished(item) {
    setBusyId(item.id);
    setActionError('');
    try {
      await updateMediaStatus(item.id, { published: !item.published });
      await loadMedia();
    } catch (error) {
      setActionError(error.message || 'Status update failed.');
    } finally {
      setBusyId('');
    }
  }

  async function handleToggleSocial(item) {
    setBusyId(item.id);
    setActionError('');
    try {
      await updateMediaStatus(item.id, { usedInSocial: !item.usedInSocial });
      await loadMedia();
    } catch (error) {
      setActionError(error.message || 'Social flag update failed.');
    } finally {
      setBusyId('');
    }
  }

  return (
    <div className="app-shell">
      <header className="hero">
        <div>
          <p className="eyebrow">No Limit</p>
          <h1>Media Management Admin</h1>
          <p className="hero-copy">
            Upload work photos, publish the website gallery, and generate square branded artwork
            for social use from one place.
          </p>
        </div>
        <div className="hero-actions">
          <span className="chip chip-success">Website-ready media</span>
          <span className="chip chip-warm">1080x1080 promo art</span>
        </div>
      </header>

      <section className="stats-grid">
        <StatCard label="Media items" value={stats.total} note="Currently returned by the active filters." />
        <StatCard label="Published" value={stats.published} note="Available for the public gallery API." />
        <StatCard label="Promo generated" value={stats.generated} note="Branded square assets ready for social use." />
      </section>

      {(categoriesState.error || mediaState.error || actionError) ? (
        <div className="alert error">
          {actionError || categoriesState.error || mediaState.error}
        </div>
      ) : null}

      <main className="layout-grid">
        <div className="left-column">
          <UploadForm
            categories={categoriesState.data}
            onSubmit={handleUpload}
            busy={busyUpload}
          />
          <PreviewPanel item={selectedItem} />
        </div>

        <div className="right-column">
          <FilterBar
            filters={filters}
            categories={categoriesState.data}
            onChange={(patch) => setFilters((current) => ({ ...current, ...patch }))}
            onRefresh={() => loadMedia().catch(() => {})}
            loading={mediaState.loading}
          />
          <MediaTable
            items={mediaState.data}
            selectedId={selectedId}
            onSelect={setSelectedId}
            onGenerate={handleGenerate}
            onTogglePublished={handleTogglePublished}
            onToggleSocial={handleToggleSocial}
            busyId={busyId}
          />
        </div>
      </main>
    </div>
  );
}
