import { useState } from 'react';

export default function UploadForm({ categories, onSubmit, busy }) {
  const [title, setTitle] = useState('');
  const [categoryId, setCategoryId] = useState('');
  const [published, setPublished] = useState(false);
  const [file, setFile] = useState(null);

  async function handleSubmit(event) {
    event.preventDefault();
    if (!file || !title || !categoryId) return;

    const formData = new FormData();
    formData.append('title', title);
    formData.append('categoryId', categoryId);
    formData.append('published', String(published));
    formData.append('image', file);

    await onSubmit(formData);

    setTitle('');
    setCategoryId('');
    setPublished(false);
    setFile(null);
    event.currentTarget.reset();
  }

  return (
    <section className="panel upload-panel">
      <div className="panel-heading">
        <div>
          <p className="eyebrow">Media Intake</p>
          <h2>Upload work photo</h2>
        </div>
        <span className="chip chip-warm">Original image</span>
      </div>

      <form className="stack" onSubmit={handleSubmit}>
        <label className="field">
          <span>Title</span>
          <input
            type="text"
            value={title}
            onChange={(event) => setTitle(event.target.value)}
            placeholder="Example: Custom white oak deck"
            required
          />
        </label>

        <label className="field">
          <span>Category</span>
          <select
            value={categoryId}
            onChange={(event) => setCategoryId(event.target.value)}
            required
          >
            <option value="">Select category</option>
            {categories.map((category) => (
              <option key={category.id} value={category.id}>
                {category.name}
              </option>
            ))}
          </select>
        </label>

        <label className="field">
          <span>Image file</span>
          <input
            type="file"
            accept="image/*"
            onChange={(event) => setFile(event.target.files?.[0] || null)}
            required
          />
        </label>

        <label className="checkbox-line">
          <input
            type="checkbox"
            checked={published}
            onChange={(event) => setPublished(event.target.checked)}
          />
          <span>Mark as published after upload</span>
        </label>

        <button className="button primary" type="submit" disabled={busy}>
          {busy ? 'Uploading...' : 'Upload image'}
        </button>
      </form>
    </section>
  );
}
