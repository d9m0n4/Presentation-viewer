import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import type { Presentation, Slide } from "./types.js";

// const NS = {
//   presentation: "http://schemas.openxmlformats.org/presentationml/2006/main",
//   drawing: "http://schemas.openxmlformats.org/drawingml/2006/main",
//   relationships:
//     "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
//   packageRelationships:
//     "http://schemas.openxmlformats.org/package/2006/relationships",
// };

const EMU_PER_PX = 9525;
const EMU_PER_POINT = 12700;
const HUNDRED_THOUSAND = 100000;

type XmlNode = Record<string, any>;

type Point = {
  x: number;
  y: number;
};

type Size = {
  width: number;
  height: number;
};

type Transform = {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
};

type FontScheme = {
  majorLatin?: string;
  minorLatin?: string;
  majorEa?: string;
  minorEa?: string;
  majorCs?: string;
  minorCs?: string;
};

type ColorContext = {
  theme: Record<string, string>;
  fontScheme?: FontScheme;
  placeholderColor?: string;
};

type Relationship = {
  id: string;
  type: string;
  target: string;
};

type Fill = {
  type: "solid" | "gradient" | "pattern" | "image";
  color?: string;
  opacity?: number;
  colors?: Array<{
    color: string;
    position: number;
    opacity?: number;
  }>;
  imageRelId?: string;
};

type Stroke = {
  color: string;
  width: number;
  opacity?: number;
  dash?: string;
  cap?: string;
  join?: string;
};

type TextRun = {
  text: string;
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  opacity?: number;
};

type TextParagraph = {
  text: string;
  runs: TextRun[];
  alignment?: string;
  level?: number;
  bullet?: boolean;
};

type TextContent = {
  paragraphs: TextParagraph[];
  text: string;
  anchor?: string;
  wrap?: string;
  marginLeft?: number;
  marginRight?: number;
  marginTop?: number;
  marginBottom?: number;
};

type ParseContext = {
  colorContext: ColorContext;
  relationships: Map<string, Relationship>;
  zip: JSZip;
  parentTransform?: GroupTransform;
  layoutPlaceholders?: Map<string, Transform>;
};

type GroupTransform = {
  x: number;
  y: number;
  width: number;
  height: number;
  childX: number;
  childY: number;
  childWidth: number;
  childHeight: number;
};

type ParsedShape = {
  id: string;
  name?: string;
  type: string;

  position: Point;
  size: Size;

  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;

  fill?: Fill;
  stroke?: Stroke;

  text?: TextContent;

  geometry?: {
    preset?: string;
  };

  placeholder?: {
    type?: string;
    index?: string;
  };

  children?: ParsedShape[];

  src?: string;
  imageRelId?: string;

  points?: number[];

  raw?: {
    presetGeometry?: string;
  };
};

export class PPTXParser {
  private readonly orderedParser: XMLParser;

  private readonly objectParser: XMLParser;

  private theme: Record<string, string> = {};

  private zip!: JSZip;

  constructor() {
    /*
     * preserveOrder нужен именно для slide/spTree.
     *
     * В PPTX порядок элементов внутри spTree = z-order.
     */
    this.orderedParser = new XMLParser({
      preserveOrder: true,
      ignoreAttributes: false,
      attributeNamePrefix: "",
      textNodeName: "#text",
      trimValues: false,
    });

    /*
     * Object parser удобнее для metadata и некоторых вспомогательных XML.
     */
    this.objectParser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      removeNSPrefix: false,
      trimValues: false,
    });
  }

  async parse(arrayBuffer: ArrayBuffer): Promise<Presentation> {
    this.zip = await JSZip.loadAsync(arrayBuffer);

    await this.parseTheme();

    const presentation = await this.parsePresentationInfo();

    const slides = await this.parseSlides(presentation);

    const metadata = await this.parseMetadata();

    return {
      slides,
      metadata,
    } as Presentation;
  }

  // ---------------------------------------------------------------------------
  // PRESENTATION
  // ---------------------------------------------------------------------------

  private async parsePresentationInfo(): Promise<{
    width: number;
    height: number;
    slideTargets: string[];
  }> {
    const file = this.zip.file("ppt/presentation.xml");

    if (!file) {
      throw new Error("Invalid PPTX: ppt/presentation.xml not found");
    }

    const xml = await file.async("string");
    const data = this.objectParser.parse(xml);

    const presentation = data?.["p:presentation"];

    const slideSize = presentation?.["p:sldSz"];

    const width = this.emuToPx(this.number(slideSize?.["@_cx"]));

    const height = this.emuToPx(this.number(slideSize?.["@_cy"]));

    const relationships = await this.parseRelationships(
      "ppt/_rels/presentation.xml.rels",
    );

    const slideIds = presentation?.["p:sldIdLst"]?.["p:sldId"];

    const ids = Array.isArray(slideIds) ? slideIds : slideIds ? [slideIds] : [];

    const slideTargets: string[] = [];

    for (const slideId of ids) {
      const relId = slideId?.["@_r:id"];

      if (!relId) continue;

      const relationship = relationships.get(relId);

      if (!relationship) continue;

      slideTargets.push(
        this.resolvePath("ppt/presentation.xml", relationship.target),
      );
    }

    return {
      width: width || 1280,
      height: height || 720,
      slideTargets,
    };
  }

  // ---------------------------------------------------------------------------
  // SLIDES
  // ---------------------------------------------------------------------------

  private async parseSlides(presentation: {
    width: number;
    height: number;
    slideTargets: string[];
  }): Promise<Slide[]> {
    const slides: Slide[] = [];

    for (let index = 0; index < presentation.slideTargets.length; index++) {
      const slidePath = presentation.slideTargets[index];

      const slide = await this.parseSlide(
        slidePath,
        index,
        presentation.width,
        presentation.height,
      );

      slides.push(slide);
    }

    return slides;
  }

  private async getLayoutPlaceholders(
    slidePath: string,
    slideRels: Map<string, Relationship>,
  ): Promise<Map<string, Transform>> {
    const layoutRel = [...slideRels.values()].find((r) =>
      r.type.endsWith("/slideLayout"),
    );
    if (!layoutRel) return new Map();

    const layoutPath = this.resolvePath(slidePath, layoutRel.target);
    const layoutFile = this.zip.file(layoutPath);
    if (!layoutFile) return new Map();

    const layoutRoot = this.orderedParser.parse(
      await layoutFile.async("string"),
    );
    const map = this.extractPlaceholderTransforms(layoutRoot, "p:sldLayout");

    // если и в layout нет xfrm — добираем из slideMaster
    const layoutRels = await this.parseRelationships(
      this.relationshipsPath(layoutPath),
    );
    const masterRel = [...layoutRels.values()].find((r) =>
      r.type.endsWith("/slideMaster"),
    );
    if (masterRel) {
      const masterPath = this.resolvePath(layoutPath, masterRel.target);
      const masterFile = this.zip.file(masterPath);
      if (masterFile) {
        const masterRoot = this.orderedParser.parse(
          await masterFile.async("string"),
        );
        const masterMap = this.extractPlaceholderTransforms(
          masterRoot,
          "p:sldMaster",
        );
        for (const [k, v] of masterMap) if (!map.has(k)) map.set(k, v);
      }
    }
    return map;
  }

  private extractPlaceholderTransforms(
    root: XmlNode,
    containerTag: string,
  ): Map<string, Transform> {
    const map = new Map<string, Transform>();
    const container = this.findChild(root, containerTag);
    const spTree =
      container &&
      this.findChild(this.findChild(container, "p:cSld")!, "p:spTree");
    if (!spTree) return map;

    for (const sp of this.elementChildren(spTree)) {
      if (this.elementName(sp) !== "p:sp") continue;
      const ph = this.parsePlaceholder(this.findChild(sp, "p:nvSpPr"));
      const spPr = this.findChild(sp, "p:spPr");
      if (!ph || !spPr) continue;
      const t = this.parseTransform(spPr);
      if (t.width && t.height) {
        map.set(`${ph.type ?? ""}:${ph.index ?? ""}`, t);
        map.set(`${ph.type ?? ""}:`, t);
      }
    }
    return map;
  }

  private async parseSlide(
    slidePath: string,
    index: number,
    slideWidth: number,
    slideHeight: number,
  ): Promise<Slide> {
    const file = this.zip.file(slidePath);

    if (!file) {
      throw new Error(`Slide not found: ${slidePath}`);
    }

    const xml = await file.async("string");

    const root = this.orderedParser.parse(xml);

    const slideNode = this.findChild(root, "p:sld");

    if (!slideNode) {
      throw new Error(`Invalid slide XML: ${slidePath}`);
    }

    /*
     * Relationships принадлежат конкретному slide.
     */
    const relPath = this.relationshipsPath(slidePath);

    const relationships = await this.parseRelationships(relPath);

    const layoutPlaceholders = await this.getLayoutPlaceholders(
      slidePath,
      relationships,
    );

    const colorContext: ColorContext = {
      theme: this.theme,
      fontScheme: this.fontScheme,
    };

    const context: ParseContext = {
      colorContext,
      relationships,
      zip: this.zip,
      layoutPlaceholders,
    };

    const cSld = this.findChild(slideNode, "p:cSld");

    const spTree = cSld ? this.findChild(cSld, "p:spTree") : undefined;

    const shapes: ParsedShape[] = [];

    if (spTree) {
      /*
       * ВАЖНО:
       * не делаем отдельно p:sp, p:pic, p:cxnSp.
       *
       * Иначе теряем z-order.
       */
      for (const child of this.elementChildren(spTree)) {
        const type = this.elementName(child);

        let shape: ParsedShape | null = null;

        switch (type) {
          case "p:sp":
            shape = this.parseShape(child, context);
            break;

          case "p:pic":
            shape = this.parsePicture(child, context);
            break;

          case "p:cxnSp":
            shape = this.parseConnector(child, context);
            break;

          case "p:grpSp":
            shape = this.parseGroup(child, context);
            break;

          case "p:graphicFrame":
            shape = this.parseGraphicFrame(child);
            break;
        }

        if (shape) {
          shapes.push(shape);
        }
      }
    }

    const background = this.parseBackground(cSld, context);

    await this.resolveImages(shapes, slidePath);

    return {
      id: this.slideIdFromPath(slidePath, index),
      shapes,
      background,
      width: slideWidth,
      height: slideHeight,
    } as Slide;
  }

  // ---------------------------------------------------------------------------
  // Images
  // ---------------------------------------------------------------------------

  private async resolveImages(shapes: any[], slidePath: string): Promise<void> {
    const relationships = await this.parseRelationships(
      this.relationshipsPath(slidePath),
    );

    for (const shape of shapes) {
      if (shape.type === "image" && shape.imageRelId) {
        const relationship = relationships.get(shape.imageRelId);

        if (!relationship) {
          continue;
        }

        const imagePath = this.resolvePath(slidePath, relationship.target);

        const imageFile = this.zip.file(imagePath);

        if (!imageFile) {
          continue;
        }

        const base64 = await imageFile.async("base64");

        const extension = imagePath.split(".").pop()?.toLowerCase();

        const mime =
          extension === "jpg" || extension === "jpeg"
            ? "image/jpeg"
            : extension === "svg"
              ? "image/svg+xml"
              : extension === "gif"
                ? "image/gif"
                : "image/png";

        shape.src = `data:${mime};base64,${base64}`;
      }

      if (shape.type === "group" && shape.children) {
        await this.resolveImages(shape.children, slidePath);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // SHAPES
  // ---------------------------------------------------------------------------

  private parseShape(node: XmlNode, context: ParseContext): ParsedShape | null {
    const nvSpPr = this.findChild(node, "p:nvSpPr");
    const cNvPr = nvSpPr ? this.findChild(nvSpPr, "p:cNvPr") : undefined;
    const placeholder = this.parsePlaceholder(nvSpPr);

    const spPr = this.findChild(node, "p:spPr");

    const txBody = this.findChild(node, "p:txBody");
    const text = this.parseTextBody(txBody, context);

    let transform: Transform = spPr
      ? this.parseTransform(spPr)
      : {
          x: 0,
          y: 0,
          width: 0,
          height: 0,
          rotation: 0,
          flipH: false,
          flipV: false,
        };

    if ((!transform.width || !transform.height) && placeholder) {
      const key = `${placeholder.type ?? ""}:${placeholder.index ?? ""}`;
      const fallback =
        context.layoutPlaceholders?.get(key) ??
        context.layoutPlaceholders?.get(`${placeholder.type ?? ""}:`);
      if (fallback) transform = fallback;
    }

    if (!transform.width && !transform.height && !text) {
      // совсем нечего рендерить
      return null;
    }

    const geometry = spPr ? this.findChild(spPr, "a:prstGeom") : undefined;
    const preset = geometry ? this.attr(geometry, "prst") : undefined;

    const fill = spPr ? this.parseFill(spPr, node, context) : undefined;
    const stroke = spPr ? this.parseStroke(spPr, node, context) : undefined;

    let type = this.mapPresetGeometry(preset);

    if (!type) {
      type = text ? "text" : "shape";
    }

    const shape: ParsedShape = {
      id: this.attr(cNvPr, "id") || "unknown",
      name: this.attr(cNvPr, "name"),
      type,
      position: { x: this.emuToPx(transform.x), y: this.emuToPx(transform.y) },
      size: {
        width: this.emuToPx(transform.width),
        height: this.emuToPx(transform.height),
      },
      rotation: transform.rotation || undefined,
      flipH: transform.flipH || undefined,
      flipV: transform.flipV || undefined,
      fill,
      stroke,
      text,
      geometry: { preset },
      placeholder,
      raw: { presetGeometry: preset },
    };

    return this.applyParentTransform(shape, context.parentTransform);
  }

  // ---------------------------------------------------------------------------
  // GROUP
  // ---------------------------------------------------------------------------

  private parseGroup(node: XmlNode, context: ParseContext): ParsedShape | null {
    const nvGrpSpPr = this.findChild(node, "p:nvGrpSpPr");

    const cNvPr = nvGrpSpPr ? this.findChild(nvGrpSpPr, "p:cNvPr") : undefined;

    const grpSpPr = this.findChild(node, "p:grpSpPr");

    const xfrm = grpSpPr ? this.findChild(grpSpPr, "a:xfrm") : undefined;

    const off = xfrm ? this.findChild(xfrm, "a:off") : undefined;

    const ext = xfrm ? this.findChild(xfrm, "a:ext") : undefined;

    const chOff = xfrm ? this.findChild(xfrm, "a:chOff") : undefined;

    const chExt = xfrm ? this.findChild(xfrm, "a:chExt") : undefined;

    const groupTransform: GroupTransform = {
      x: this.number(this.attr(off, "x")),
      y: this.number(this.attr(off, "y")),
      width: this.number(this.attr(ext, "cx")),
      height: this.number(this.attr(ext, "cy")),
      childX: this.number(this.attr(chOff, "x")),
      childY: this.number(this.attr(chOff, "y")),
      childWidth: this.number(this.attr(chExt, "cx")),
      childHeight: this.number(this.attr(chExt, "cy")),
    };

    const childContext: ParseContext = {
      ...context,
      parentTransform: groupTransform,
    };

    const children: ParsedShape[] = [];

    for (const child of this.elementChildren(node)) {
      const type = this.elementName(child);

      let shape: ParsedShape | null = null;

      switch (type) {
        case "p:sp":
          shape = this.parseShape(child, childContext);
          break;

        case "p:pic":
          shape = this.parsePicture(child, childContext);
          break;

        case "p:cxnSp":
          shape = this.parseConnector(child, childContext);
          break;

        case "p:grpSp":
          shape = this.parseGroup(child, childContext);
          break;

        case "p:graphicFrame":
          shape = this.parseGraphicFrame(child);
          break;
      }

      if (shape) {
        children.push(shape);
      }
    }

    return {
      id: this.attr(cNvPr, "id") || "group",
      name: this.attr(cNvPr, "name"),
      type: "group",

      position: {
        x: this.emuToPx(groupTransform.x),
        y: this.emuToPx(groupTransform.y),
      },

      size: {
        width: this.emuToPx(groupTransform.width),
        height: this.emuToPx(groupTransform.height),
      },

      children,
    };
  }

  // ---------------------------------------------------------------------------
  // IMAGE
  // ---------------------------------------------------------------------------

  private parsePicture(
    node: XmlNode,
    context: ParseContext,
  ): ParsedShape | null {
    const nvPicPr = this.findChild(node, "p:nvPicPr");

    const cNvPr = nvPicPr ? this.findChild(nvPicPr, "p:cNvPr") : undefined;

    const blipFill = this.findChild(node, "p:blipFill");

    const blip = blipFill ? this.findChild(blipFill, "a:blip") : undefined;

    const relId = this.attr(blip, "r:embed");

    const spPr = this.findChild(node, "p:spPr");

    if (!spPr) return null;

    const transform = this.parseTransform(spPr);

    let src: string | undefined;

    if (relId) {
      src = awaitablePlaceholder();

      /*
       * Реальный Data URL нельзя получить синхронно.
       *
       * Поэтому parsePicture возвращает relationship ID.
       * Ниже parseSlides заменяет его после asynchronous resolve.
       *
       * Для удобства viewer можно использовать imageRelId.
       */
    }

    const shape: ParsedShape = {
      id: this.attr(cNvPr, "id") || "unknown",
      name: this.attr(cNvPr, "name"),
      type: "image",

      position: {
        x: this.emuToPx(transform.x),
        y: this.emuToPx(transform.y),
      },

      size: {
        width: this.emuToPx(transform.width),
        height: this.emuToPx(transform.height),
      },

      rotation: transform.rotation || undefined,
      flipH: transform.flipH || undefined,
      flipV: transform.flipV || undefined,

      imageRelId: relId,
      src,
    };

    /*
     * Нам нужно реально разрешить изображение.
     *
     * Это делаем позже в parseSlides.
     */
    return this.applyParentTransform(shape, context.parentTransform);
  }

  // ---------------------------------------------------------------------------
  // CONNECTOR
  // ---------------------------------------------------------------------------

  private parseConnector(
    node: XmlNode,
    context: ParseContext,
  ): ParsedShape | null {
    const nvCxnSpPr = this.findChild(node, "p:nvCxnSpPr");

    const cNvPr = nvCxnSpPr ? this.findChild(nvCxnSpPr, "p:cNvPr") : undefined;

    const spPr = this.findChild(node, "p:spPr");

    if (!spPr) return null;

    const transform = this.parseTransform(spPr);

    const geometry = this.findChild(spPr, "a:prstGeom");

    const preset = geometry ? this.attr(geometry, "prst") : undefined;

    const stroke = this.parseStroke(spPr, node, context);

    const shape: ParsedShape = {
      id: this.attr(cNvPr, "id") || "unknown",
      name: this.attr(cNvPr, "name"),
      type: "line",

      position: {
        x: this.emuToPx(transform.x),
        y: this.emuToPx(transform.y),
      },

      size: {
        width: this.emuToPx(transform.width),
        height: this.emuToPx(transform.height),
      },

      rotation: transform.rotation || undefined,
      flipH: transform.flipH || undefined,
      flipV: transform.flipV || undefined,

      stroke,

      geometry: {
        preset,
      },

      points: [
        0,
        0,
        this.emuToPx(transform.width),
        this.emuToPx(transform.height),
      ],
    };

    return this.applyParentTransform(shape, context.parentTransform);
  }

  // ---------------------------------------------------------------------------
  // GRAPHIC FRAME
  // ---------------------------------------------------------------------------

  private parseGraphicFrame(
    node: XmlNode,
    // context: ParseContext,
  ): ParsedShape | null {
    const nvGraphicFramePr = this.findChild(node, "p:nvGraphicFramePr");

    const cNvPr = nvGraphicFramePr
      ? this.findChild(nvGraphicFramePr, "p:cNvPr")
      : undefined;

    const xfrm = this.findChild(node, "p:xfrm");

    if (!xfrm) {
      return null;
    }

    const off = this.findChild(xfrm, "a:off");
    const ext = this.findChild(xfrm, "a:ext");

    const x = this.number(this.attr(off, "x"));
    const y = this.number(this.attr(off, "y"));
    const width = this.number(this.attr(ext, "cx"));
    const height = this.number(this.attr(ext, "cy"));

    /*
     * Пока сохраняем graphicFrame как отдельный объект.
     *
     * Внутри может быть:
     *
     * - таблица
     * - chart
     * - SmartArt
     */
    return {
      id: this.attr(cNvPr, "id") || "unknown",
      name: this.attr(cNvPr, "name"),
      type: "graphicFrame",

      position: {
        x: this.emuToPx(x),
        y: this.emuToPx(y),
      },

      size: {
        width: this.emuToPx(width),
        height: this.emuToPx(height),
      },
    };
  }

  // ---------------------------------------------------------------------------
  // TRANSFORM
  // ---------------------------------------------------------------------------

  private parseTransform(spPr: XmlNode): Transform {
    const xfrm = this.findChild(spPr, "a:xfrm");

    if (!xfrm) {
      return {
        x: 0,
        y: 0,
        width: 0,
        height: 0,
        rotation: 0,
        flipH: false,
        flipV: false,
      };
    }

    const off = this.findChild(xfrm, "a:off");
    const ext = this.findChild(xfrm, "a:ext");

    return {
      x: this.number(this.attr(off, "x")),
      y: this.number(this.attr(off, "y")),
      width: this.number(this.attr(ext, "cx")),
      height: this.number(this.attr(ext, "cy")),

      /*
       * PPTX rotation = 1/60000 degree.
       */
      rotation: this.number(this.attr(xfrm, "rot")) / 60000,

      flipH: this.attr(xfrm, "flipH") === "1",
      flipV: this.attr(xfrm, "flipV") === "1",
    };
  }

  private applyParentTransform(
    shape: ParsedShape,
    parent?: GroupTransform,
  ): ParsedShape {
    if (!parent) {
      return shape;
    }

    if (!parent.childWidth || !parent.childHeight) {
      return shape;
    }

    const scaleX = parent.width / parent.childWidth;
    const scaleY = parent.height / parent.childHeight;

    const localX = this.pxToEmu(shape.position.x);
    const localY = this.pxToEmu(shape.position.y);

    const localWidth = this.pxToEmu(shape.size.width);
    const localHeight = this.pxToEmu(shape.size.height);

    const x = parent.x + (localX - parent.childX) * scaleX;

    const y = parent.y + (localY - parent.childY) * scaleY;

    shape.position = {
      x: this.emuToPx(x),
      y: this.emuToPx(y),
    };

    shape.size = {
      width: this.emuToPx(localWidth * scaleX),
      height: this.emuToPx(localHeight * scaleY),
    };

    return shape;
  }

  // ---------------------------------------------------------------------------
  // GEOMETRY
  // ---------------------------------------------------------------------------

  private mapPresetGeometry(preset?: string): string | undefined {
    if (!preset) return undefined;

    const map: Record<string, string> = {
      rect: "rectangle",
      roundRect: "roundedRectangle",

      ellipse: "ellipse",

      triangle: "triangle",
      rtTriangle: "rightTriangle",

      diamond: "diamond",

      parallelogram: "parallelogram",
      trapezoid: "trapezoid",

      pentagon: "pentagon",
      hexagon: "hexagon",
      heptagon: "heptagon",
      octagon: "octagon",

      homePlate: "homePlate",

      chevron: "chevron",
      pie: "pie",
      chord: "chord",
      arc: "arc",

      leftArrow: "leftArrow",
      rightArrow: "rightArrow",
      upArrow: "upArrow",
      downArrow: "downArrow",

      leftRightArrow: "leftRightArrow",
      upDownArrow: "upDownArrow",

      plus: "plus",
      minus: "minus",

      heart: "heart",
      cloud: "cloud",
      lightningBolt: "lightningBolt",

      star5: "star5",
      star6: "star6",
      star7: "star7",
      star8: "star8",
      star10: "star10",
      star12: "star12",
      star16: "star16",
      star24: "star24",
      star32: "star32",

      flowChartProcess: "flowChartProcess",
      flowChartDecision: "flowChartDecision",
      flowChartTerminator: "flowChartTerminator",

      line: "line",
      lineInv: "line",
    };

    return map[preset] || preset;
  }

  // ---------------------------------------------------------------------------
  // FILL
  // ---------------------------------------------------------------------------

  private parseFill(
    spPr: XmlNode,
    shapeNode: XmlNode,
    context: ParseContext,
  ): Fill | undefined {
    /*
     * 1. Explicit fill.
     */

    const solidFill = this.findChild(spPr, "a:solidFill");

    if (solidFill) {
      const color = this.parseColor(solidFill, context.colorContext);

      if (color) {
        return {
          type: "solid",
          color: color.color,
          opacity: color.opacity,
        };
      }
    }

    /*
     * 2. Gradient.
     */

    const gradFill = this.findChild(spPr, "a:gradFill");

    if (gradFill) {
      const gsLst = this.findChild(gradFill, "a:gsLst");

      const stops = gsLst
        ? this.elementChildren(gsLst).filter(
            (node) => this.elementName(node) === "a:gs",
          )
        : [];

      const colors = [];

      for (const stop of stops) {
        const color = this.parseColor(stop, context.colorContext);

        if (!color) continue;

        colors.push({
          color: color.color,
          opacity: color.opacity,
          position: this.number(this.attr(stop, "pos")) / HUNDRED_THOUSAND,
        });
      }

      if (colors.length) {
        return {
          type: "gradient",
          colors,
        };
      }
    }

    /*
     * 3. Pattern.
     */

    const pattFill = this.findChild(spPr, "a:pattFill");

    if (pattFill) {
      const fgClr = this.findChild(pattFill, "a:fgClr");

      const color = fgClr
        ? this.parseColor(fgClr, context.colorContext)
        : undefined;

      return {
        type: "pattern",
        color: color?.color,
        opacity: color?.opacity,
      };
    }

    /*
     * 4. Image fill.
     */

    const blipFill = this.findChild(shapeNode, "p:blipFill");

    if (blipFill) {
      const blip = this.findChild(blipFill, "a:blip");

      const relId = this.attr(blip, "r:embed");

      if (relId) {
        return {
          type: "image",
          imageRelId: relId,
        };
      }
    }

    /*
     * 5. Explicit noFill.
     */

    if (this.findChild(spPr, "a:noFill")) {
      return undefined;
    }

    /*
     * 6. p:style/a:fillRef.
     *
     * Это как раз твой прямоугольник.
     */

    const style = this.findChild(shapeNode, "p:style");

    const fillRef = style ? this.findChild(style, "a:fillRef") : undefined;

    if (fillRef) {
      const color = this.parseColor(fillRef, context.colorContext);

      if (color) {
        return {
          type: "solid",
          color: color.color,
          opacity: color.opacity,
        };
      }
    }

    /*
     * 7. Не придумываем дефолтную синюю заливку.
     */
    return undefined;
  }

  // ---------------------------------------------------------------------------
  // STROKE
  // ---------------------------------------------------------------------------

  private parseStroke(
    spPr: XmlNode,
    shapeNode: XmlNode,
    context: ParseContext,
  ): Stroke | undefined {
    const line = this.findChild(spPr, "a:ln");

    if (line) {
      const widthEMU = this.number(this.attr(line, "w"));

      const solidFill = this.findChild(line, "a:solidFill");

      const color = solidFill
        ? this.parseColor(solidFill, context.colorContext)
        : undefined;

      if (color) {
        return {
          color: color.color,
          opacity: color.opacity,
          width: widthEMU > 0 ? widthEMU / EMU_PER_POINT : 1,
          dash: this.parseDash(line),
          cap: this.attr(line, "cap"),
        };
      }

      /*
       * Даже если цвет не указан явно,
       * линия существует.
       *
       * Но не подставляем случайный цвет.
       */
    }

    /*
     * Style line reference.
     */

    const style = this.findChild(shapeNode, "p:style");

    const lnRef = style ? this.findChild(style, "a:lnRef") : undefined;

    if (lnRef) {
      const color = this.parseColor(lnRef, context.colorContext);

      if (color) {
        return {
          color: color.color,
          opacity: color.opacity,
          width: 1,
        };
      }
    }

    return undefined;
  }

  private parseDash(line: XmlNode): string | undefined {
    const dash = this.findChild(line, "a:prstDash");

    return dash ? this.attr(dash, "val") : undefined;
  }

  // ---------------------------------------------------------------------------
  // COLOR
  // ---------------------------------------------------------------------------

  private parseColor(
    node: XmlNode,
    context: ColorContext,
  ):
    | {
        color: string;
        opacity: number;
      }
    | undefined {
    if (!node) return undefined;

    const srgb = this.findChild(node, "a:srgbClr");

    if (srgb) {
      const value = this.attr(srgb, "val");

      if (value) {
        return this.applyColorTransforms(`#${value}`, srgb);
      }
    }

    const scheme = this.findChild(node, "a:schemeClr");

    if (scheme) {
      const value = this.attr(scheme, "val");

      if (value) {
        const base = context.theme[value] || this.defaultThemeColor(value);

        return this.applyColorTransforms(base, scheme);
      }
    }

    const system = this.findChild(node, "a:sysClr");

    if (system) {
      const lastColor = this.attr(system, "lastClr");

      if (lastColor) {
        return this.applyColorTransforms(`#${lastColor}`, system);
      }
    }

    const preset = this.findChild(node, "a:prstClr");

    if (preset) {
      const value = this.attr(preset, "val");

      const color = this.presetColor(value);

      if (color) {
        return this.applyColorTransforms(color, preset);
      }
    }

    return undefined;
  }

  private applyColorTransforms(
    input: string,
    node: XmlNode,
  ): {
    color: string;
    opacity: number;
  } {
    let rgb = this.hexToRgb(input);

    if (!rgb) {
      return {
        color: input,
        opacity: 1,
      };
    }

    let alpha = 1;

    const tint = this.number(this.attr(node, "tint"));

    if (tint) {
      const amount = tint / HUNDRED_THOUSAND;

      rgb = {
        r: rgb.r + (255 - rgb.r) * amount,
        g: rgb.g + (255 - rgb.g) * amount,
        b: rgb.b + (255 - rgb.b) * amount,
      };
    }

    const shade = this.number(this.attr(node, "shade"));

    if (shade) {
      const amount = shade / HUNDRED_THOUSAND;

      rgb = {
        r: rgb.r * amount,
        g: rgb.g * amount,
        b: rgb.b * amount,
      };
    }

    const lumMod = this.number(this.attr(node, "lumMod"));

    if (lumMod) {
      const amount = lumMod / HUNDRED_THOUSAND;

      rgb = {
        r: rgb.r * amount,
        g: rgb.g * amount,
        b: rgb.b * amount,
      };
    }

    const lumOff = this.number(this.attr(node, "lumOff"));

    if (lumOff) {
      const amount = lumOff / HUNDRED_THOUSAND;

      rgb = {
        r: rgb.r + 255 * amount,
        g: rgb.g + 255 * amount,
        b: rgb.b + 255 * amount,
      };
    }

    const alphaValue = this.number(this.attr(node, "alpha"));

    if (alphaValue) {
      alpha = alphaValue / HUNDRED_THOUSAND;
    }

    return {
      color: this.rgbToHex(rgb),
      opacity: Math.max(0, Math.min(1, alpha)),
    };
  }

  // ---------------------------------------------------------------------------
  // TEXT
  // ---------------------------------------------------------------------------

  private parseTextBody(
    txBody: XmlNode | undefined,
    context: ParseContext,
  ): TextContent | undefined {
    if (!txBody) return undefined;

    const bodyPr = this.findChild(txBody, "a:bodyPr");

    const paragraphs = this.elementChildren(txBody).filter(
      (node) => this.elementName(node) === "a:p",
    );

    if (!paragraphs.length) {
      return undefined;
    }

    const result: TextParagraph[] = [];

    for (const paragraph of paragraphs) {
      const runs: TextRun[] = [];

      for (const child of this.elementChildren(paragraph)) {
        const type = this.elementName(child);

        if (type === "a:r") {
          const run = this.parseTextRun(child, context);

          if (run) {
            runs.push(run);
          }
        }

        /*
         * Field.
         */
        if (type === "a:fld") {
          const run = this.parseTextRun(child, context);

          if (run) {
            runs.push(run);
          }
        }

        /*
         * Text inside hyperlink.
         */
        if (type === "a:hlinkClick") {
          // hyperlink metadata can be handled later
        }
      }

      const pPr = this.findChild(paragraph, "a:pPr");

      const text = runs.map((run) => run.text).join("");

      result.push({
        text,
        runs,
        alignment: this.attr(pPr, "algn"),
        level: this.number(this.attr(pPr, "lvl")),
        bullet: !!(
          pPr &&
          (this.findChild(pPr, "a:buChar") ||
            this.findChild(pPr, "a:buAutoNum"))
        ),
      });
    }

    const text = result.map((paragraph) => paragraph.text).join("\n");

    return {
      paragraphs: result,
      text,

      anchor: this.attr(bodyPr, "anchor"),

      wrap: this.attr(bodyPr, "wrap"),

      marginLeft: this.emuToPx(this.number(this.attr(bodyPr, "lIns"))),

      marginRight: this.emuToPx(this.number(this.attr(bodyPr, "rIns"))),

      marginTop: this.emuToPx(this.number(this.attr(bodyPr, "tIns"))),

      marginBottom: this.emuToPx(this.number(this.attr(bodyPr, "bIns"))),
    };
  }

  private resolveFontFamily(
    typeface: string | undefined,
    fontScheme?: FontScheme,
  ): string | undefined {
    if (!typeface) return undefined;

    switch (typeface) {
      case "+mn-lt":
        return fontScheme?.minorLatin;
      case "+mj-lt":
        return fontScheme?.majorLatin;
      case "+mn-ea":
        return fontScheme?.minorEa;
      case "+mj-ea":
        return fontScheme?.majorEa;
      case "+mn-cs":
        return fontScheme?.minorCs;
      case "+mj-cs":
        return fontScheme?.majorCs;
      default:
        return typeface;
    }
  }

  private parseTextRun(node: XmlNode, context: ParseContext): TextRun | null {
    const rPr =
      this.findChild(node, "a:rPr") || this.findChild(node, "a:defRPr");
    const t = this.findChild(node, "a:t");
    const text = t ? this.textContent(t) : "";

    if (!text) {
      return null;
    }

    const solidFill = rPr ? this.findChild(rPr, "a:solidFill") : undefined;

    const color = solidFill
      ? this.parseColor(solidFill, context.colorContext)
      : undefined;

    const latin = rPr ? this.findChild(rPr, "a:latin") : undefined;
    const rawFontFamily = latin ? this.attr(latin, "typeface") : undefined;
    const fontFamily = this.resolveFontFamily(
      rawFontFamily,
      context.colorContext.fontScheme,
    );
    const size = this.number(this.attr(rPr, "sz"));

    return {
      text,
      fontFamily: fontFamily || undefined,
      fontSize: size > 0 ? size / 100 : undefined,
      bold: this.attr(rPr, "b") === "1" ? true : undefined,
      italic: this.attr(rPr, "i") === "1" ? true : undefined,
      underline:
        this.attr(rPr, "u") && this.attr(rPr, "u") !== "none"
          ? true
          : undefined,
      strike:
        this.attr(rPr, "strike") && this.attr(rPr, "strike") !== "noStrike"
          ? true
          : undefined,
      color: color?.color,
      opacity: color?.opacity,
    };
  }

  // ---------------------------------------------------------------------------
  // PLACEHOLDER
  // ---------------------------------------------------------------------------

  private parsePlaceholder(nvSpPr: XmlNode | undefined):
    | {
        type?: string;
        index?: string;
      }
    | undefined {
    if (!nvSpPr) return undefined;

    const nvPr = this.findChild(nvSpPr, "p:nvPr");

    if (!nvPr) return undefined;

    const ph = this.findChild(nvPr, "p:ph");

    if (!ph) return undefined;

    return {
      type: this.attr(ph, "type"),
      index: this.attr(ph, "idx"),
    };
  }

  // ---------------------------------------------------------------------------
  // BACKGROUND
  // ---------------------------------------------------------------------------

  private parseBackground(
    cSld: XmlNode | undefined,
    context: ParseContext,
  ): any {
    if (!cSld) return undefined;

    /*
     * p:bg is a sibling of p:spTree.
     */
    const bg = this.findChild(cSld, "p:bg");

    if (!bg) return undefined;

    const bgPr = this.findChild(bg, "p:bgPr");

    if (!bgPr) return undefined;

    const solidFill = this.findChild(bgPr, "a:solidFill");

    if (solidFill) {
      const color = this.parseColor(solidFill, context.colorContext);

      if (color) {
        return {
          type: "solid",
          color: color.color,
          opacity: color.opacity,
        };
      }
    }

    const gradFill = this.findChild(bgPr, "a:gradFill");

    if (gradFill) {
      return {
        type: "gradient",
      };
    }

    const blipFill = this.findChild(bgPr, "a:blipFill");

    if (blipFill) {
      const blip = this.findChild(blipFill, "a:blip");

      return {
        type: "image",
        imageRelId: this.attr(blip, "r:embed"),
      };
    }

    return undefined;
  }

  // ---------------------------------------------------------------------------
  // THEME
  // ---------------------------------------------------------------------------

  private fontScheme: FontScheme = {};

  private async parseTheme(): Promise<void> {
    const themeFile = this.zip.file("ppt/theme/theme1.xml");

    if (!themeFile) {
      this.theme = {};
      return;
    }

    const xml = await themeFile.async("string");

    const root = this.orderedParser.parse(xml);

    const theme = this.findChild(root, "a:theme");

    const themeElements = theme
      ? this.findChild(theme, "a:themeElements")
      : undefined;

    const clrScheme = themeElements
      ? this.findChild(themeElements, "a:clrScheme")
      : undefined;

    const fontScheme = themeElements
      ? this.findChild(themeElements, "a:fontScheme")
      : undefined;
    const majorFont = fontScheme
      ? this.findChild(fontScheme, "a:majorFont")
      : undefined;
    const minorFont = fontScheme
      ? this.findChild(fontScheme, "a:minorFont")
      : undefined;

    const readLatin = (fontNode?: XmlNode) => {
      const latin = fontNode ? this.findChild(fontNode, "a:latin") : undefined;
      return latin ? this.attr(latin, "typeface") : undefined;
    };

    const readEa = (fontNode?: XmlNode) => {
      const ea = fontNode ? this.findChild(fontNode, "a:ea") : undefined;
      return ea ? this.attr(ea, "typeface") : undefined;
    };
    const readCs = (fontNode?: XmlNode) => {
      const cs = fontNode ? this.findChild(fontNode, "a:cs") : undefined;
      return cs ? this.attr(cs, "typeface") : undefined;
    };

    this.fontScheme = {
      majorLatin: readLatin(majorFont) || "Calibri Light",
      minorLatin: readLatin(minorFont) || "Calibri",
      majorEa: readEa(majorFont),
      minorEa: readEa(minorFont),
      majorCs: readCs(majorFont),
      minorCs: readCs(minorFont),
    };

    const result: Record<string, string> = {};

    if (clrScheme) {
      for (const colorNode of this.elementChildren(clrScheme)) {
        const name = this.elementName(colorNode);

        if (!name) {
          continue;
        }

        const value = this.findChild(colorNode, "a:srgbClr");

        if (value) {
          const hex = this.attr(value, "val");

          if (hex) {
            result[name.replace("a:", "")] = `#${hex}`;
          }

          continue;
        }

        const sysClr = this.findChild(colorNode, "a:sysClr");

        if (sysClr) {
          const hex = this.attr(sysClr, "lastClr");

          if (hex) {
            result[name.replace("a:", "")] = `#${hex}`;
          }
        }
      }
    }

    this.theme = result;
  }

  private defaultThemeColor(value: string): string {
    const defaults: Record<string, string> = {
      dk1: "#000000",
      lt1: "#FFFFFF",
      dk2: "#44546A",
      lt2: "#E7E6E6",

      tx1: "#000000",
      tx2: "#44546A",
      bg1: "#FFFFFF",
      bg2: "#E7E6E6",

      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",

      hlink: "#0563C1",
      folHlink: "#954F72",
    };

    return defaults[value] || "#000000";
  }

  // ---------------------------------------------------------------------------
  // RELATIONSHIPS
  // ---------------------------------------------------------------------------

  private async parseRelationships(
    path: string,
  ): Promise<Map<string, Relationship>> {
    const file = this.zip.file(path);

    if (!file) {
      return new Map();
    }

    const xml = await file.async("string");

    const data = this.objectParser.parse(xml);

    const relationships = data?.Relationships?.Relationship;

    const list = Array.isArray(relationships)
      ? relationships
      : relationships
        ? [relationships]
        : [];

    const map = new Map<string, Relationship>();

    for (const relationship of list) {
      const id = relationship?.["@_Id"];

      const type = relationship?.["@_Type"];

      const target = relationship?.["@_Target"];

      if (!id || !target) continue;

      map.set(id, {
        id,
        type,
        target,
      });
    }

    return map;
  }

  // ---------------------------------------------------------------------------
  // METADATA
  // ---------------------------------------------------------------------------

  private async parseMetadata(): Promise<any> {
    const file = this.zip.file("docProps/core.xml");

    if (!file) return {};

    const xml = await file.async("string");

    const data = this.objectParser.parse(xml);

    const core = data?.["cp:coreProperties"];

    return {
      title: core?.["dc:title"],

      author: core?.["dc:creator"],

      subject: core?.["dc:subject"],

      description: core?.["dc:description"],

      created: core?.["dcterms:created"],

      modified: core?.["dcterms:modified"],
    };
  }

  // ---------------------------------------------------------------------------
  // XML HELPERS
  // ---------------------------------------------------------------------------

  private elementName(node: XmlNode): string | undefined {
    if (!node) return undefined;

    return Object.keys(node).find((key) => key !== ":@" && key !== "#text");
  }

  private elementChildren(node: XmlNode): XmlNode[] {
    if (!node) return [];

    const result: XmlNode[] = [];

    for (const key of Object.keys(node)) {
      if (key === ":@" || key === "#text") {
        continue;
      }

      const value = node[key];

      if (Array.isArray(value)) {
        for (const child of value) {
          if (child && typeof child === "object") {
            result.push(child);
          }
        }
      } else if (value && typeof value === "object") {
        result.push(value);
      }
    }

    return result;
  }

  private findChild(
    node: XmlNode[] | XmlNode | undefined,
    name: string,
  ): XmlNode | undefined {
    if (!node) return undefined;

    const children = Array.isArray(node) ? node : this.elementChildren(node);

    return children.find((child) => this.elementName(child) === name);
  }

  private attr(node: XmlNode | undefined, name: string): string | undefined {
    if (!node) return undefined;

    const attrs = node[":@"];

    if (attrs && name in attrs) {
      return attrs[name];
    }

    return undefined;
  }

  private textContent(node: XmlNode): string {
    if (!node) return "";

    let result = "";

    for (const key of Object.keys(node)) {
      if (key === "#text") {
        result += String(node[key] ?? "");
        continue;
      }

      if (key === ":@") continue;

      const value = node[key];

      if (Array.isArray(value)) {
        for (const child of value) {
          if (child && typeof child === "object") {
            result += this.textContent(child);
          }
        }
      } else if (value && typeof value === "object") {
        result += this.textContent(value);
      }
    }

    return result;
  }

  // ---------------------------------------------------------------------------
  // PATHS
  // ---------------------------------------------------------------------------

  private relationshipsPath(filePath: string): string {
    const slash = filePath.lastIndexOf("/");

    const directory = slash >= 0 ? filePath.slice(0, slash) : "";

    const fileName = slash >= 0 ? filePath.slice(slash + 1) : filePath;

    return directory
      ? `${directory}/_rels/${fileName}.rels`
      : `_rels/${fileName}.rels`;
  }

  private resolvePath(sourcePath: string, target: string): string {
    if (target.startsWith("/")) {
      return target.slice(1);
    }

    const sourceDir = sourcePath.slice(0, sourcePath.lastIndexOf("/"));

    const parts = [...sourceDir.split("/"), ...target.split("/")];

    const result: string[] = [];

    for (const part of parts) {
      if (!part || part === ".") continue;

      if (part === "..") {
        result.pop();
      } else {
        result.push(part);
      }
    }

    return result.join("/");
  }

  // ---------------------------------------------------------------------------
  // COLOR UTILS
  // ---------------------------------------------------------------------------

  private presetColor(value?: string): string | undefined {
    if (!value) return undefined;

    const colors: Record<string, string> = {
      black: "#000000",
      white: "#FFFFFF",
      red: "#FF0000",
      green: "#008000",
      blue: "#0000FF",
      yellow: "#FFFF00",
      cyan: "#00FFFF",
      magenta: "#FF00FF",

      gray: "#808080",
      grey: "#808080",

      orange: "#FFA500",
      purple: "#800080",
      pink: "#FFC0CB",
      brown: "#A52A2A",

      transparent: "#000000",
    };

    return colors[value.toLowerCase()];
  }

  private hexToRgb(hex: string): {
    r: number;
    g: number;
    b: number;
  } | null {
    const normalized = hex.replace("#", "");

    if (normalized.length !== 6 && normalized.length !== 8) {
      return null;
    }

    return {
      r: parseInt(normalized.slice(0, 2), 16),
      g: parseInt(normalized.slice(2, 4), 16),
      b: parseInt(normalized.slice(4, 6), 16),
    };
  }

  private rgbToHex(rgb: { r: number; g: number; b: number }): string {
    const r = Math.round(Math.max(0, Math.min(255, rgb.r)))
      .toString(16)
      .padStart(2, "0");

    const g = Math.round(Math.max(0, Math.min(255, rgb.g)))
      .toString(16)
      .padStart(2, "0");

    const b = Math.round(Math.max(0, Math.min(255, rgb.b)))
      .toString(16)
      .padStart(2, "0");

    return `#${r}${g}${b}`.toUpperCase();
  }

  // ---------------------------------------------------------------------------
  // NUMBERS
  // ---------------------------------------------------------------------------

  private number(value: unknown): number {
    if (value === undefined || value === null || value === "") {
      return 0;
    }

    const parsed = Number(value);

    return Number.isFinite(parsed) ? parsed : 0;
  }

  private emuToPx(value: number): number {
    return value / EMU_PER_PX;
  }

  private pxToEmu(value: number): number {
    return value * EMU_PER_PX;
  }

  private slideIdFromPath(path: string, index: number): string {
    const match = path.match(/slide(\d+)\.xml$/);

    return match ? `slide${match[1]}` : `slide${index + 1}`;
  }
}

/*
 * parsePicture сейчас асинхронно разрешает imageRelId
 * после того, как известен relationship map.
 *
 * Поэтому используем отдельную функцию для Data URL.
 */
function awaitablePlaceholder(): string {
  return "";
}

// -----------------------------------------------------------------------------
// PUBLIC API
// -----------------------------------------------------------------------------

export async function parsePPTX(
  file: File | ArrayBuffer,
): Promise<Presentation> {
  const parser = new PPTXParser();

  const buffer = file instanceof File ? await file.arrayBuffer() : file;

  return parser.parse(buffer);
}
