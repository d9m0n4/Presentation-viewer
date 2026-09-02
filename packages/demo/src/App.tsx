import { useState } from 'react';
import { parsePPTX, type Presentation } from '@presentation-viewer/core';
import { SlideCanvas } from './SlideCanvas';
import './App.css';

function App() {
  const [presentation, setPresentation] = useState<Presentation | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [currentSlide, setCurrentSlide] = useState(0);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    setCurrentSlide(0);

    try {
      const pres = await parsePPTX(file);
      setPresentation(pres);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to parse PPTX');
      console.error('Parse error:', err);
    } finally {
      setLoading(false);
    }
  };

  const goToPrevSlide = () => {
    setCurrentSlide((prev) => Math.max(0, prev - 1));
  };

  const goToNextSlide = () => {
    if (!presentation) return;
    setCurrentSlide((prev) => Math.min(presentation.slides.length - 1, prev + 1));
  };

  return (
    <div className="app">
      <header className="header">
        <h1>📊 Presentation Viewer Demo</h1>
        <p>Upload a PPTX file to see it rendered on Canvas</p>
      </header>

      <main className="main">
        <div className="upload-section">
          <label htmlFor="file-input" className="upload-button">
            {loading ? 'Loading...' : 'Choose PPTX File'}
          </label>
          <input
            id="file-input"
            type="file"
            accept=".pptx"
            onChange={handleFileUpload}
            disabled={loading}
            style={{ display: 'none' }}
          />
        </div>

        {error && (
          <div className="error">
            <strong>Error:</strong> {error}
          </div>
        )}

        {presentation && (
          <>
            <div className="viewer-section">
              <div className="slide-controls">
                <button onClick={goToPrevSlide} disabled={currentSlide === 0}>
                  ← Previous
                </button>
                <span className="slide-counter">
                  Slide {currentSlide + 1} / {presentation.slides.length}
                </span>
                <button
                  onClick={goToNextSlide}
                  disabled={currentSlide === presentation.slides.length - 1}
                >
                  Next →
                </button>
              </div>

              <div className="slide-canvas-container">
                <SlideCanvas slide={presentation.slides[currentSlide]} />
              </div>
            </div>

            <div className="result">
              <h2>Parsed Data</h2>

              {presentation.metadata && (
                <div className="metadata">
                  <h3>Metadata</h3>
                  {presentation.metadata.title && (
                    <p>
                      <strong>Title:</strong> {presentation.metadata.title}
                    </p>
                  )}
                  {presentation.metadata.author && (
                    <p>
                      <strong>Author:</strong> {presentation.metadata.author}
                    </p>
                  )}
                </div>
              )}

              <div className="slides">
                <h3>Slides ({presentation.slides.length})</h3>
                {presentation.slides.map((slide, index) => (
                  <div
                    key={slide.id}
                    className={`slide-card ${index === currentSlide ? 'active' : ''}`}
                    onClick={() => setCurrentSlide(index)}
                  >
                    <div className="slide-header">
                      <span className="slide-number">Slide {index + 1}</span>
                      <span className="slide-id">{slide.id}</span>
                    </div>
                    <div className="slide-info">
                      <p>Shapes: {slide.shapes.length}</p>
                      {slide.background && <p>Background: {slide.background.type}</p>}
                    </div>
                  </div>
                ))}
              </div>

              <details className="raw-data">
                <summary>Raw JSON Data</summary>
                <pre>{JSON.stringify(presentation, null, 2)}</pre>
              </details>
            </div>
          </>
        )}
      </main>
    </div>
  );
}

export default App;
