import { useEffect, useRef } from 'react';
import { SlideRenderer } from '@presentation-viewer/renderer';
import type { Slide } from '@presentation-viewer/core';

interface SlideCanvasProps {
  slide: Slide;
  width?: number;
  height?: number;
  debug?: boolean;
}

export function SlideCanvas({ slide, width = 960, height = 720, debug = true }: SlideCanvasProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    if (!canvasRef.current) return;

    const renderer = new SlideRenderer(canvasRef.current);
    renderer.render(slide, {
      width,
      height,
      backgroundColor: '#ffffff',
      debug,
    });

    return () => {
      renderer.clear();
    };
  }, [slide, width, height, debug]);

  return (
    <canvas
      ref={canvasRef}
      style={{
        width: '100%',
        height: 'auto',
        border: '1px solid #e0e0e0',
        borderRadius: '8px',
      }}
    />
  );
}
