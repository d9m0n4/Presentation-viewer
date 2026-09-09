import { RectangleShape } from "@presentation-viewer/core";
import type {
  Slide,
  Shape,
  TextShape,
  ImageShape,
  GeometryShape,
  LineShape,
  GroupShape,
  Fill,
} from "./types.js";

import type { RenderOptions } from "./types.js";

export class SlideRenderer {
  private readonly canvas: HTMLCanvasElement;
  private readonly ctx: CanvasRenderingContext2D;

  constructor(canvas: HTMLCanvasElement) {
    this.canvas = canvas;

    const ctx = canvas.getContext("2d");

    if (!ctx) {
      throw new Error("Failed to get 2D context from canvas");
    }

    this.ctx = ctx;
  }

  async render(slide: Slide, options: RenderOptions = {}): Promise<void> {
    const {
      width = slide.width || 960,
      height = slide.height || 720,
      backgroundColor = "#ffffff",
      scale = 1,
      debug = false,
    } = options;

    this.canvas.width = Math.max(1, Math.round(width * scale));

    this.canvas.height = Math.max(1, Math.round(height * scale));

    this.ctx.setTransform(scale, 0, 0, scale, 0, 0);

    this.ctx.clearRect(0, 0, width, height);

    /*
     * Parser already converted EMU -> px.
     */
    const slideWidth = slide.width || width;
    const slideHeight = slide.height || height;

    /*
     * Scale the original slide into requested output size.
     */
    const scaleX = width / slideWidth;
    const scaleY = height / slideHeight;

    this.ctx.save();

    this.ctx.scale(scaleX, scaleY);

    /*
     * Background
     */
    this.ctx.fillStyle = backgroundColor;

    this.ctx.fillRect(0, 0, slideWidth, slideHeight);

    if (slide.background) {
      this.renderBackground(slide.background, slideWidth, slideHeight);
    }

    /*
     * Shapes are already ordered according to spTree z-order.
     */
    for (const shape of slide.shapes) {
      await this.renderShape(shape, debug);
    }

    if (debug) {
      this.renderSlideDebug(slideWidth, slideHeight);
    }

    this.ctx.restore();
  }

  private async renderShape(shape: Shape, debug: boolean): Promise<void> {
    switch (shape.type) {
      case "text":
        this.renderTextShape(shape, debug);
        break;

      case "image":
        await this.renderImageShape(shape, debug);
        break;

      case "rectangle":
        this.renderRectangleShape(shape, debug);
        break;

      case "shape":
      case "roundedRectangle":
      case "ellipse":
      case "triangle":
      case "rightTriangle":
      case "diamond":
      case "parallelogram":
      case "trapezoid":
      case "pentagon":
      case "hexagon":
      case "heptagon":
      case "octagon":
      case "homePlate":
      case "chevron":
      case "pie":
      case "chord":
      case "arc":
      case "leftArrow":
      case "rightArrow":
      case "upArrow":
      case "downArrow":
      case "leftRightArrow":
      case "upDownArrow":
      case "plus":
      case "minus":
      case "heart":
      case "cloud":
      case "lightningBolt":
      case "star5":
      case "star6":
      case "star7":
      case "star8":
      case "star10":
      case "star12":
      case "star16":
      case "star24":
      case "star32":
      case "flowChartProcess":
      case "flowChartDecision":
      case "flowChartTerminator":
        this.renderGeometryShape(shape, debug);
        break;

      case "line":
        this.renderLineShape(shape, debug);
        break;

      case "group":
        await this.renderGroupShape(shape, debug);
        break;

      case "graphicFrame":
        this.renderGraphicFrame(shape, debug);
        break;

      default:
        if (debug) {
          console.warn(
            `[SlideRenderer] Unsupported shape type: ${shape}`,
            shape,
          );
        }

        this.renderUnknownShape(shape);
    }
  }
  private renderBackground(
    background: NonNullable<Slide["background"]>,
    width: number,
    height: number,
  ): void {
    switch (background.type) {
      case "solid": {
        if (!background.color) {
          return;
        }

        this.ctx.save();

        this.ctx.globalAlpha = background.opacity ?? 1;

        this.ctx.fillStyle = background.color;

        this.ctx.fillRect(0, 0, width, height);

        this.ctx.restore();

        break;
      }

      case "gradient": {
        /*
         * Gradient parser currently only returns
         * the type, so there is nothing reliable
         * to render yet.
         */
        break;
      }

      case "image": {
        /*
         * Image background can be implemented once
         * parser resolves imageRelId to imageUrl.
         */
        break;
      }
    }
  }

  private renderGeometryShape(shape: GeometryShape, debug: boolean): void {
    const { x, y, width, height } = this.getShapeBounds(shape);

    this.ctx.save();

    this.applyTransform(
      x,
      y,
      width,
      height,
      shape.rotation ?? 0,
      shape.flipH ?? false,
      shape.flipV ?? false,
    );

    const geometry = shape.geometry?.preset ?? shape.type;

    const path = this.createGeometryPath(geometry, width, height);

    this.renderFill(shape.fill, path);

    this.renderStroke(shape.stroke, path);

    this.ctx.restore();

    /*
     * Text is rendered after the geometry.
     *
     * This is important for:
     *
     * rectangle + text
     * roundedRectangle + text
     * arrows + text
     * etc.
     */
    if (shape.text) {
      this.renderShapeText(shape, debug);
    }

    if (debug) {
      this.renderDebugBox(x, y, width, height, "#0000ff");
    }
  }

  private renderRectangleShape(shape: RectangleShape, debug: boolean): void {
    const { x, y, width, height } = this.getShapeBounds(shape);

    this.ctx.save();

    this.applyTransform(
      x,
      y,
      width,
      height,
      shape.rotation ?? 0,
      shape.flipH ?? false,
      shape.flipV ?? false,
    );

    const path = new Path2D();

    path.rect(0, 0, width, height);

    this.renderFill(shape.fill, path);

    this.renderStroke(shape.stroke, path);

    this.ctx.restore();

    if (debug) {
      this.renderDebugBox(x, y, width, height, "#ff0000");
    }
  }

  private renderShapeText(shape: GeometryShape, debug: boolean): void {
    if (!shape.text) {
      return;
    }

    this.renderTextContent(
      shape.text,
      shape.position.x,
      shape.position.y,
      shape.size.width,
      shape.size.height,
      shape.rotation ?? 0,
      shape.flipH ?? false,
      shape.flipV ?? false,
    );

    if (debug) {
      this.renderDebugBox(
        shape.position.x,
        shape.position.y,
        shape.size.width,
        shape.size.height,
        "#00ffff",
      );
    }
  }

  private renderTextShape(shape: TextShape, debug: boolean): void {
    this.renderTextContent(
      shape.text,
      shape.position.x,
      shape.position.y,
      shape.size.width,
      shape.size.height,
      shape.rotation ?? 0,
      shape.flipH ?? false,
      shape.flipV ?? false,
    );

    if (debug) {
      this.renderDebugBox(
        shape.position.x,
        shape.position.y,
        shape.size.width,
        shape.size.height,
        "#0000ff",
      );
    }
  }

  private renderTextContent(
    text: TextShape["text"],
    x: number,
    y: number,
    width: number,
    height: number,
    rotation: number,
    flipH: boolean,
    flipV: boolean,
  ): void {
    this.ctx.save();

    this.applyTransform(x, y, width, height, rotation, flipH, flipV);

    const marginLeft = text.marginLeft ?? 0;

    const marginRight = text.marginRight ?? 0;

    const marginTop = text.marginTop ?? 0;

    const marginBottom = text.marginBottom ?? 0;

    const contentWidth = Math.max(0, width - marginLeft - marginRight);

    const contentHeight = Math.max(0, height - marginTop - marginBottom);

    /*
     * Render paragraphs separately so that
     * paragraph alignment is preserved.
     */
    const lineHeight = this.getLineHeight(text);

    const totalLines = text.paragraphs.reduce((count, paragraph) => {
      const fontSize = this.getParagraphFontSize(paragraph);

      this.ctx.font = this.createFont(fontSize, paragraph.runs[0]);

      return count + this.wrapParagraph(paragraph.text, contentWidth).length;
    }, 0);

    const totalHeight = totalLines * lineHeight;

    let textY = marginTop;

    switch (text.anchor) {
      case "ctr":
        textY = marginTop + Math.max(0, (contentHeight - totalHeight) / 2);
        break;

      case "b":
        textY = height - marginBottom - totalHeight;
        break;
    }

    for (const paragraph of text.paragraphs) {
      const paragraphLines = this.wrapParagraph(paragraph.text, contentWidth);

      const alignment = this.mapAlignment(paragraph.alignment);

      for (const line of paragraphLines) {
        this.renderTextLine(
          paragraph,
          line,
          alignment,
          marginLeft,
          contentWidth,
          textY,
        );

        textY += lineHeight;
      }
    }

    this.ctx.restore();
  }

  private renderTextLine(
    paragraph: TextShape["text"]["paragraphs"][number],
    line: string,
    alignment: CanvasTextAlign,
    x: number,
    width: number,
    y: number,
  ): void {
    if (!paragraph.runs.length) {
      return;
    }

    console.log(line);

    /*
     * Canvas text drawing by runs.
     *
     * For now we preserve run formatting.
     */
    const totalWidth = paragraph.runs.reduce((result, run) => {
      const fontSize = run.fontSize ?? 18;

      this.ctx.font = this.createFont(fontSize, run);

      return result + this.ctx.measureText(run.text).width;
    }, 0);

    let startX = x;

    if (alignment === "center") {
      startX = x + (width - totalWidth) / 2;
    }

    if (alignment === "right") {
      startX = x + width - totalWidth;
    }

    let currentX = startX;

    for (const run of paragraph.runs) {
      const fontSize = run.fontSize ?? 18;

      this.ctx.font = this.createFont(fontSize, run);

      this.ctx.fillStyle = run.color ?? "#000000";

      this.ctx.globalAlpha = run.opacity ?? 1;

      this.ctx.textBaseline = "top";

      this.ctx.fillText(run.text, currentX, y);

      if (run.underline) {
        const metrics = this.ctx.measureText(run.text);

        this.ctx.beginPath();

        this.ctx.moveTo(currentX, y + fontSize + 2);

        this.ctx.lineTo(currentX + metrics.width, y + fontSize + 2);

        this.ctx.strokeStyle = run.color ?? "#000000";

        this.ctx.lineWidth = 1;

        this.ctx.stroke();
      }

      currentX += this.ctx.measureText(run.text).width;
    }

    this.ctx.globalAlpha = 1;
  }

  private createFont(
    fontSize: number,
    run?: {
      fontFamily?: string;
      bold?: boolean;
      italic?: boolean;
    },
  ): string {
    const fontFamily = run?.fontFamily ?? "Arial";

    const fontWeight = run?.bold ? "700" : "400";

    const fontStyle = run?.italic ? "italic" : "normal";

    return [fontStyle, fontWeight, `${fontSize}px`, fontFamily].join(" ");
  }

  private getParagraphFontSize(
    paragraph: TextShape["text"]["paragraphs"][number],
  ): number {
    return paragraph.runs[0]?.fontSize ?? 18;
  }

  private getLineHeight(text: TextShape["text"]): number {
    const firstRun = text.paragraphs[0]?.runs[0];

    const fontSize = firstRun?.fontSize ?? 18;

    return fontSize * 1.2;
  }

  private wrapParagraph(text: string, maxWidth: number): string[] {
    if (!text) {
      return [""];
    }

    const words = text.split(/\s+/);

    const lines: string[] = [];

    let current = "";

    for (const word of words) {
      const candidate = current ? `${current} ${word}` : word;

      if (current && this.ctx.measureText(candidate).width > maxWidth) {
        lines.push(current);
        current = word;
      } else {
        current = candidate;
      }
    }

    if (current) {
      lines.push(current);
    }

    return lines;
  }

  private mapAlignment(alignment?: string): CanvasTextAlign {
    switch (alignment) {
      case "ctr":
        return "center";

      case "r":
        return "right";

      default:
        return "left";
    }
  }

  private async renderImageShape(
    shape: ImageShape,
    debug: boolean,
  ): Promise<void> {
    const { x, y, width, height } = this.getShapeBounds(shape);

    if (!shape.src) {
      this.renderImagePlaceholder(x, y, width, height);

      return;
    }

    try {
      const image = await this.loadImage(shape.src);

      this.ctx.save();

      this.applyTransform(
        x,
        y,
        width,
        height,
        shape.rotation ?? 0,
        shape.flipH ?? false,
        shape.flipV ?? false,
      );

      this.ctx.drawImage(image, 0, 0, width, height);

      this.ctx.restore();

      if (debug) {
        this.renderDebugBox(x, y, width, height, "#00ff00");
      }
    } catch (error) {
      console.warn(`[SlideRenderer] Failed to render image ${shape.id}`, error);

      this.renderImagePlaceholder(x, y, width, height);
    }
  }

  private renderLineShape(shape: LineShape, debug: boolean): void {
    const { x, y, width, height } = this.getShapeBounds(shape);

    this.ctx.save();

    this.applyTransform(
      x,
      y,
      width,
      height,
      shape.rotation ?? 0,
      shape.flipH ?? false,
      shape.flipV ?? false,
    );

    this.ctx.beginPath();

    this.ctx.moveTo(0, 0);

    this.ctx.lineTo(width, height);

    const stroke = shape.stroke;

    this.ctx.strokeStyle = stroke?.color ?? "#000000";

    this.ctx.globalAlpha = stroke?.opacity ?? 1;

    this.ctx.lineWidth = stroke?.width ?? 1;

    if (stroke?.dash) {
      this.ctx.setLineDash(this.mapDashPattern(stroke.dash));
    }

    this.ctx.stroke();

    this.ctx.restore();

    if (debug) {
      this.renderDebugBox(x, y, width, height, "#ffff00");
    }
  }

  private async renderGroupShape(
    shape: GroupShape,
    debug: boolean,
  ): Promise<void> {
    this.ctx.save();

    for (const child of shape.children) {
      await this.renderShape(child, debug);
    }

    this.ctx.restore();

    if (debug) {
      this.renderDebugBox(
        shape.position.x,
        shape.position.y,
        shape.size.width,
        shape.size.height,
        "#ff00ff",
      );
    }
  }

  private renderGraphicFrame(shape: Shape, debug: boolean): void {
    if (!debug) {
      return;
    }

    this.renderDebugBox(
      shape.position.x,
      shape.position.y,
      shape.size.width,
      shape.size.height,
      "#ff8800",
    );
  }

  private renderFill(fill: Fill | undefined, path: Path2D): void {
    if (!fill) {
      return;
    }

    switch (fill.type) {
      case "solid":
      case "pattern": {
        if (!fill.color) {
          return;
        }

        this.ctx.save();

        this.ctx.globalAlpha = fill.opacity ?? 1;

        this.ctx.fillStyle = fill.color;

        this.ctx.fill(path);

        this.ctx.restore();

        break;
      }

      case "gradient": {
        if (!fill.colors?.length) {
          return;
        }

        const gradient = this.ctx.createLinearGradient(0, 0, 1, 1);

        for (const stop of fill.colors) {
          gradient.addColorStop(
            Math.max(0, Math.min(1, stop.position)),
            stop.color,
          );
        }

        this.ctx.save();

        this.ctx.fillStyle = gradient;

        this.ctx.fill(path);

        this.ctx.restore();

        break;
      }

      case "image":
        /*
         * Image fills need the same image
         * resolution mechanism as image shapes.
         */
        break;
    }
  }

  private renderStroke(stroke: Shape["stroke"], path: Path2D): void {
    if (!stroke?.color) {
      return;
    }

    this.ctx.save();

    this.ctx.globalAlpha = stroke.opacity ?? 1;

    this.ctx.strokeStyle = stroke.color;

    this.ctx.lineWidth = stroke.width ?? 1;

    if (stroke.dash) {
      this.ctx.setLineDash(this.mapDashPattern(stroke.dash));
    }

    if (stroke.cap) {
      this.ctx.lineCap = this.mapLineCap(stroke.cap);
    }

    this.ctx.stroke(path);

    this.ctx.restore();
  }

  private createGeometryPath(
    geometry: string,
    width: number,
    height: number,
  ): Path2D {
    const path = new Path2D();

    switch (geometry) {
      case "rect":
      case "rectangle":
      case "shape":
        path.rect(0, 0, width, height);
        break;

      case "roundRect":
      case "roundedRectangle":
        this.roundRectPath(
          path,
          0,
          0,
          width,
          height,
          Math.min(width, height) * 0.12,
        );
        break;

      case "ellipse":
        path.ellipse(
          width / 2,
          height / 2,
          width / 2,
          height / 2,
          0,
          0,
          Math.PI * 2,
        );
        break;

      case "triangle":
        path.moveTo(width / 2, 0);

        path.lineTo(width, height);

        path.lineTo(0, height);

        path.closePath();
        break;

      case "rightTriangle":
        path.moveTo(0, 0);

        path.lineTo(width, height);

        path.lineTo(0, height);

        path.closePath();
        break;

      case "diamond":
        path.moveTo(width / 2, 0);

        path.lineTo(width, height / 2);

        path.lineTo(width / 2, height);

        path.lineTo(0, height / 2);

        path.closePath();
        break;

      case "parallelogram": {
        const offset = width * 0.2;

        path.moveTo(offset, 0);

        path.lineTo(width, 0);

        path.lineTo(width - offset, height);

        path.lineTo(0, height);

        path.closePath();

        break;
      }

      case "trapezoid": {
        const offset = width * 0.2;

        path.moveTo(offset, 0);

        path.lineTo(width - offset, 0);

        path.lineTo(width, height);

        path.lineTo(0, height);

        path.closePath();

        break;
      }

      case "hexagon": {
        const offset = width * 0.25;

        path.moveTo(offset, 0);

        path.lineTo(width - offset, 0);

        path.lineTo(width, height / 2);

        path.lineTo(width - offset, height);

        path.lineTo(offset, height);

        path.lineTo(0, height / 2);

        path.closePath();

        break;
      }

      case "pentagon": {
        this.createRegularPolygon(path, width, height, 5);

        break;
      }

      case "heptagon": {
        this.createRegularPolygon(path, width, height, 7);

        break;
      }

      case "octagon": {
        this.createRegularPolygon(path, width, height, 8);

        break;
      }

      case "star5":
      case "star6":
      case "star7":
      case "star8":
      case "star10":
      case "star12":
      case "star16":
      case "star24":
      case "star32": {
        const points = Number(geometry.slice(4)) || 5;

        this.createStar(path, width, height, points);

        break;
      }

      case "plus": {
        const x = width * 0.3;

        const y = height * 0.3;

        path.moveTo(x, 0);

        path.lineTo(width - x, 0);

        path.lineTo(width - x, y);

        path.lineTo(width, y);

        path.lineTo(width, height - y);

        path.lineTo(width - x, height - y);

        path.lineTo(width - x, height);

        path.lineTo(x, height);

        path.lineTo(x, height - y);

        path.lineTo(0, height - y);

        path.lineTo(0, y);

        path.lineTo(x, y);

        path.closePath();

        break;
      }

      case "minus": {
        const y = height * 0.35;

        path.rect(0, y, width, height - y * 2);

        break;
      }

      case "leftArrow":
        this.createArrow(path, width, height, "left");
        break;

      case "rightArrow":
        this.createArrow(path, width, height, "right");
        break;

      case "upArrow":
        this.createArrow(path, width, height, "up");
        break;

      case "downArrow":
        this.createArrow(path, width, height, "down");
        break;

      default:
        path.rect(0, 0, width, height);
        break;
    }

    return path;
  }

  private createRegularPolygon(
    path: Path2D,
    width: number,
    height: number,
    sides: number,
  ): void {
    const cx = width / 2;

    const cy = height / 2;

    const radius = Math.min(width, height) / 2;

    for (let i = 0; i < sides; i++) {
      const angle = -Math.PI / 2 + (i * Math.PI * 2) / sides;

      const x = cx + Math.cos(angle) * radius;

      const y = cy + Math.sin(angle) * radius;

      if (i === 0) {
        path.moveTo(x, y);
      } else {
        path.lineTo(x, y);
      }
    }

    path.closePath();
  }

  private createStar(
    path: Path2D,
    width: number,
    height: number,
    points: number,
  ): void {
    const cx = width / 2;

    const cy = height / 2;

    const outerRadius = Math.min(width, height) / 2;

    const innerRadius = outerRadius * 0.45;

    const total = points * 2;

    for (let i = 0; i < total; i++) {
      const radius = i % 2 === 0 ? outerRadius : innerRadius;

      const angle = -Math.PI / 2 + (i * Math.PI * 2) / total;

      const x = cx + Math.cos(angle) * radius;

      const y = cy + Math.sin(angle) * radius;

      if (i === 0) {
        path.moveTo(x, y);
      } else {
        path.lineTo(x, y);
      }
    }

    path.closePath();
  }

  private createArrow(
    path: Path2D,
    width: number,
    height: number,
    direction: "left" | "right" | "up" | "down",
  ): void {
    const shaft = 0.4;

    if (direction === "right") {
      const shaftY = height * shaft;

      path.moveTo(width * 0.55, 0);

      path.lineTo(width, height / 2);

      path.lineTo(width * 0.55, height);

      path.lineTo(width * 0.55, shaftY + height * 0.25);

      path.lineTo(0, shaftY + height * 0.25);

      path.lineTo(0, height - shaftY - height * 0.25);

      path.lineTo(width * 0.55, height - shaftY - height * 0.25);

      path.closePath();

      return;
    }

    if (direction === "left") {
      path.moveTo(width * 0.45, 0);

      path.lineTo(0, height / 2);

      path.lineTo(width * 0.45, height);

      path.lineTo(width * 0.45, height * 0.65);

      path.lineTo(width, height * 0.65);

      path.lineTo(width, height * 0.35);

      path.lineTo(width * 0.45, height * 0.35);

      path.closePath();

      return;
    }

    if (direction === "up") {
      path.moveTo(width / 2, 0);

      path.lineTo(width, height * 0.45);

      path.lineTo(width * 0.65, height * 0.45);

      path.lineTo(width * 0.65, height);

      path.lineTo(width * 0.35, height);

      path.lineTo(width * 0.35, height * 0.45);

      path.lineTo(0, height * 0.45);

      path.closePath();

      return;
    }

    path.moveTo(width / 2, height);

    path.lineTo(width, height * 0.55);

    path.lineTo(width * 0.65, height * 0.55);

    path.lineTo(width * 0.65, 0);

    path.lineTo(width * 0.35, 0);

    path.lineTo(width * 0.35, height * 0.55);

    path.lineTo(0, height * 0.55);

    path.closePath();
  }

  private roundRectPath(
    path: Path2D,
    x: number,
    y: number,
    width: number,
    height: number,
    radius: number,
  ): void {
    const r = Math.min(radius, width / 2, height / 2);

    path.moveTo(x + r, y);

    path.lineTo(x + width - r, y);

    path.arcTo(x + width, y, x + width, y + r, r);

    path.lineTo(x + width, y + height - r);

    path.arcTo(x + width, y + height, x + width - r, y + height, r);

    path.lineTo(x + r, y + height);

    path.arcTo(x, y + height, x, y + height - r, r);

    path.lineTo(x, y + r);

    path.arcTo(x, y, x + r, y, r);

    path.closePath();
  }

  private applyTransform(
    x: number,
    y: number,
    width: number,
    height: number,
    rotation: number,
    flipH: boolean,
    flipV: boolean,
  ): void {
    const angle = (rotation * Math.PI) / 180;

    this.ctx.translate(x + width / 2, y + height / 2);

    this.ctx.rotate(angle);

    this.ctx.scale(flipH ? -1 : 1, flipV ? -1 : 1);

    this.ctx.translate(-width / 2, -height / 2);
  }

  private getShapeBounds(shape: Shape): {
    x: number;
    y: number;
    width: number;
    height: number;
  } {
    return {
      x: shape.position.x,
      y: shape.position.y,
      width: shape.size.width,
      height: shape.size.height,
    };
  }

  private renderImagePlaceholder(
    x: number,
    y: number,
    width: number,
    height: number,
  ): void {
    this.ctx.save();

    this.ctx.fillStyle = "#e0e0e0";

    this.ctx.fillRect(x, y, width, height);

    this.ctx.strokeStyle = "#999999";

    this.ctx.strokeRect(x, y, width, height);

    this.ctx.beginPath();

    this.ctx.moveTo(x, y);

    this.ctx.lineTo(x + width, y + height);

    this.ctx.moveTo(x + width, y);

    this.ctx.lineTo(x, y + height);

    this.ctx.stroke();

    this.ctx.restore();
  }

  private renderUnknownShape(shape: Shape): void {
    this.ctx.save();

    this.ctx.strokeStyle = "#ff8800";

    this.ctx.lineWidth = 2;

    this.ctx.setLineDash([5, 5]);

    this.ctx.strokeRect(
      shape.position.x,
      shape.position.y,
      shape.size.width,
      shape.size.height,
    );

    this.ctx.restore();
  }

  private renderDebugBox(
    x: number,
    y: number,
    width: number,
    height: number,
    color: string,
  ): void {
    this.ctx.save();

    this.ctx.strokeStyle = color;

    this.ctx.lineWidth = 1;

    this.ctx.strokeRect(x, y, width, height);

    this.ctx.restore();
  }

  private renderSlideDebug(width: number, height: number): void {
    this.ctx.save();

    this.ctx.strokeStyle = "#ff0000";

    this.ctx.lineWidth = 2;

    this.ctx.strokeRect(0, 0, width, height);

    this.ctx.restore();
  }

  private mapDashPattern(dash: string): number[] {
    switch (dash) {
      case "dash":
        return [8, 6];

      case "dot":
        return [2, 4];

      case "dashDot":
        return [8, 4, 2, 4];

      case "lgDash":
        return [14, 6];

      case "lgDashDot":
        return [14, 6, 2, 6];

      case "lgDashDotDot":
        return [14, 6, 2, 6, 2, 6];

      default:
        return [];
    }
  }

  private mapLineCap(cap: string): CanvasLineCap {
    switch (cap) {
      case "rnd":
        return "round";

      case "sq":
        return "square";

      default:
        return "butt";
    }
  }

  private loadImage(src: string): Promise<HTMLImageElement> {
    return new Promise((resolve, reject) => {
      const image = new Image();

      image.onload = () => resolve(image);

      image.onerror = () => reject(new Error(`Failed to load image: ${src}`));

      image.src = src;
    });
  }

  clear(): void {
    this.ctx.setTransform(1, 0, 0, 1, 0, 0);

    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
  }

  toDataURL(type = "image/png", quality?: number): string {
    return this.canvas.toDataURL(type, quality);
  }

  async toBlob(type = "image/png", quality?: number): Promise<Blob> {
    return new Promise((resolve, reject) => {
      this.canvas.toBlob(
        (blob) => {
          if (blob) {
            resolve(blob);
          } else {
            reject(new Error("Failed to create blob"));
          }
        },
        type,
        quality,
      );
    });
  }
}
