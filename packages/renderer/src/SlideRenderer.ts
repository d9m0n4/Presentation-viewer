import type { Slide, TextShape, ImageShape, RectangleShape } from '@presentation-viewer/core';
import type { RenderOptions } from './types.js';

/**
 * SlideRenderer renders presentation slides to HTML Canvas
 */
export class SlideRenderer {
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;

  // PPTX uses EMU (English Metric Units): 914400 EMU = 1 inch = 96 pixels
  private readonly EMU_PER_PIXEL = 9525;

  // Standard slide dimensions in PPTX (10" x 7.5")
  private readonly SLIDE_WIDTH_EMU = 9144000;
  private readonly SLIDE_HEIGHT_EMU = 6858000;

  constructor(canvas: HTMLCanvasElement) {
    this.canvas = canvas;
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      throw new Error('Failed to get 2D context from canvas');
    }
    this.ctx = ctx;
  }

  /**
   * Render a slide to the canvas
   */
  render(slide: Slide, options: RenderOptions = {}): void {
    const { width = 960, height = 720, backgroundColor = '#ffffff', scale = 1, debug = false } = options;

    // Set canvas size
    this.canvas.width = width * scale;
    this.canvas.height = height * scale;

    // Scale context for high DPI
    this.ctx.scale(scale, scale);

    // Clear canvas
    this.ctx.fillStyle = backgroundColor;
    this.ctx.fillRect(0, 0, width, height);

    // Calculate scale factors
    const scaleX = width / (this.SLIDE_WIDTH_EMU / this.EMU_PER_PIXEL);
    const scaleY = height / (this.SLIDE_HEIGHT_EMU / this.EMU_PER_PIXEL);

    // Debug: draw slide boundaries
    if (debug) {
      this.ctx.strokeStyle = '#ff0000';
      this.ctx.lineWidth = 2;
      this.ctx.strokeRect(0, 0, width, height);
    }

    // Render background
    if (slide.background) {
      this.renderBackground(slide.background, width, height);
    }

    // Render shapes
    for (const shape of slide.shapes) {
      switch (shape.type) {
        case 'text':
          this.renderTextShape(shape as TextShape, scaleX, scaleY, debug);
          break;
        case 'image':
          this.renderImageShape(shape as ImageShape, scaleX, scaleY, debug);
          break;
        case 'rectangle':
          this.renderRectangleShape(shape as RectangleShape, scaleX, scaleY, debug);
          break;
        case 'line':
          this.renderLineShape(shape as any, scaleX, scaleY, debug);
          break;
      }
    }
  }

  private renderBackground(background: any, width: number, height: number): void {
    if (background.type === 'solid' && background.color) {
      this.ctx.fillStyle = background.color;
      this.ctx.fillRect(0, 0, width, height);
    }
  }

  private renderTextShape(shape: TextShape, scaleX: number, scaleY: number, debug = false): void {
    const x = (shape.position.x / this.EMU_PER_PIXEL) * scaleX;
    const y = (shape.position.y / this.EMU_PER_PIXEL) * scaleY;
    const width = (shape.size.width / this.EMU_PER_PIXEL) * scaleX;
    const height = (shape.size.height / this.EMU_PER_PIXEL) * scaleY;

    // Set text style
    const fontSize = shape.style?.fontSize || 18;
    const fontFamily = shape.style?.fontFamily || 'Arial, sans-serif';
    const color = shape.style?.color || '#000000';
    const align = shape.style?.align || 'left';
    const bold = shape.style?.bold ? 'bold ' : '';
    const italic = shape.style?.italic ? 'italic ' : '';

    this.ctx.font = `${italic}${bold}${fontSize}px ${fontFamily}`;
    this.ctx.fillStyle = color;
    this.ctx.textAlign = align;
    this.ctx.textBaseline = 'top';

    // Calculate text position based on alignment
    let textX = x;
    if (align === 'center') {
      textX = x + width / 2;
    } else if (align === 'right') {
      textX = x + width;
    }

    // Word wrap text
    const lines = this.wrapText(shape.text, width);
    const lineHeight = fontSize * 1.2;

    // Draw text lines
    lines.forEach((line, index) => {
      this.ctx.fillText(line, textX, y + index * lineHeight);
    });

    // Debug: draw bounding box
    if (debug) {
      this.ctx.strokeStyle = '#0000ff';
      this.ctx.lineWidth = 1;
      this.ctx.strokeRect(x, y, width, height);
    }
  }

  private renderImageShape(shape: ImageShape, scaleX: number, scaleY: number, debug = false): void {
    const x = (shape.position.x / this.EMU_PER_PIXEL) * scaleX;
    const y = (shape.position.y / this.EMU_PER_PIXEL) * scaleY;
    const width = (shape.size.width / this.EMU_PER_PIXEL) * scaleX;
    const height = (shape.size.height / this.EMU_PER_PIXEL) * scaleY;

    const img = new Image();
    img.onload = () => {
      this.ctx.drawImage(img, x, y, width, height);
    };
    img.onerror = () => {
      // Draw placeholder rectangle on error
      this.ctx.fillStyle = '#e0e0e0';
      this.ctx.fillRect(x, y, width, height);
      this.ctx.strokeStyle = '#999999';
      this.ctx.strokeRect(x, y, width, height);

      // Draw X to indicate missing image
      this.ctx.strokeStyle = '#999999';
      this.ctx.beginPath();
      this.ctx.moveTo(x, y);
      this.ctx.lineTo(x + width, y + height);
      this.ctx.moveTo(x + width, y);
      this.ctx.lineTo(x, y + height);
      this.ctx.stroke();
    };
    img.src = shape.src;

    if (debug) {
      this.ctx.strokeStyle = '#00ff00';
      this.ctx.lineWidth = 1;
      this.ctx.strokeRect(x, y, width, height);
    }
  }

  private renderRectangleShape(shape: RectangleShape, scaleX: number, scaleY: number, debug = false): void {
    const x = (shape.position.x / this.EMU_PER_PIXEL) * scaleX;
    const y = (shape.position.y / this.EMU_PER_PIXEL) * scaleY;
    const width = (shape.size.width / this.EMU_PER_PIXEL) * scaleX;
    const height = (shape.size.height / this.EMU_PER_PIXEL) * scaleY;

    // Draw fill
    if (shape.fill) {
      this.ctx.fillStyle = shape.fill.color;
      this.ctx.globalAlpha = shape.fill.opacity ?? 1;
      this.ctx.fillRect(x, y, width, height);
      this.ctx.globalAlpha = 1;
    }

    // Draw stroke
    if (shape.stroke) {
      this.ctx.strokeStyle = shape.stroke.color;
      this.ctx.lineWidth = shape.stroke.width ?? 1;
      this.ctx.strokeRect(x, y, width, height);
    }

    if (debug) {
      this.ctx.strokeStyle = '#ff00ff';
      this.ctx.lineWidth = 2;
      this.ctx.strokeRect(x, y, width, height);
    }
  }

  private renderLineShape(shape: any, scaleX: number, scaleY: number, debug = false): void {
    const x = (shape.position.x / this.EMU_PER_PIXEL) * scaleX;
    const y = (shape.position.y / this.EMU_PER_PIXEL) * scaleY;
    const width = (shape.size.width / this.EMU_PER_PIXEL) * scaleX;
    const height = (shape.size.height / this.EMU_PER_PIXEL) * scaleY;

    // Draw line from top-left to bottom-right
    this.ctx.beginPath();
    this.ctx.moveTo(x, y);
    this.ctx.lineTo(x + width, y + height);

    if (shape.stroke) {
      this.ctx.strokeStyle = shape.stroke.color;
      this.ctx.lineWidth = shape.stroke.width ?? 1;
    } else {
      this.ctx.strokeStyle = '#000000';
      this.ctx.lineWidth = 1;
    }

    this.ctx.stroke();

    if (debug) {
      this.ctx.strokeStyle = '#ffff00';
      this.ctx.lineWidth = 1;
      this.ctx.strokeRect(x, y, width, height);
    }
  }

  private wrapText(text: string, maxWidth: number): string[] {
    const words = text.split(' ');
    const lines: string[] = [];
    let currentLine = '';

    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word;
      const metrics = this.ctx.measureText(testLine);

      if (metrics.width > maxWidth && currentLine) {
        lines.push(currentLine);
        currentLine = word;
      } else {
        currentLine = testLine;
      }
    }

    if (currentLine) {
      lines.push(currentLine);
    }

    return lines;
  }

  /**
   * Clear the canvas
   */
  clear(): void {
    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
  }

  /**
   * Export canvas as data URL
   */
  toDataURL(type: string = 'image/png', quality?: number): string {
    return this.canvas.toDataURL(type, quality);
  }

  /**
   * Export canvas as Blob
   */
  async toBlob(type: string = 'image/png', quality?: number): Promise<Blob> {
    return new Promise((resolve, reject) => {
      this.canvas.toBlob(
        (blob) => {
          if (blob) {
            resolve(blob);
          } else {
            reject(new Error('Failed to create blob'));
          }
        },
        type,
        quality,
      );
    });
  }
}
