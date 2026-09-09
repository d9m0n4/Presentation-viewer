export interface Presentation {
  slides: Slide[];
  metadata?: PresentationMetadata;
}

export interface PresentationMetadata {
  title?: string;
  author?: string;
  subject?: string;
  description?: string;
  created?: string;
  modified?: string;
}

export interface Slide {
  id: string;
  shapes: Shape[];
  background?: Background;

  /**
   * Slide dimensions are already converted from EMU to CSS pixels
   * by the parser.
   */
  width: number;
  height: number;
}

export type Shape =
  | TextShape
  | ImageShape
  | RectangleShape
  | GeometryShape
  | LineShape
  | GroupShape
  | GraphicFrameShape;

export interface BaseShape {
  id: string;
  name?: string;

  type: string;

  /**
   * All coordinates are CSS pixels.
   */
  position: Position;
  size: Size;

  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;

  fill?: Fill;
  stroke?: Stroke;

  placeholder?: Placeholder;
}

export interface TextShape extends BaseShape {
  type: "text";

  text: TextContent;

  geometry?: ShapeGeometry;

  raw?: RawShapeData;
}

export interface ImageShape extends BaseShape {
  type: "image";

  src?: string;
  imageRelId?: string;
}

export interface RectangleShape extends BaseShape {
  type: "rectangle";

  text?: TextContent;
  geometry?: ShapeGeometry;

  raw?: RawShapeData;
}

export interface GeometryShape extends BaseShape {
  type:
    | "shape"
    | "roundedRectangle"
    | "ellipse"
    | "triangle"
    | "rightTriangle"
    | "diamond"
    | "parallelogram"
    | "trapezoid"
    | "pentagon"
    | "hexagon"
    | "heptagon"
    | "octagon"
    | "homePlate"
    | "chevron"
    | "pie"
    | "chord"
    | "arc"
    | "leftArrow"
    | "rightArrow"
    | "upArrow"
    | "downArrow"
    | "leftRightArrow"
    | "upDownArrow"
    | "plus"
    | "minus"
    | "heart"
    | "cloud"
    | "lightningBolt"
    | "star5"
    | "star6"
    | "star7"
    | "star8"
    | "star10"
    | "star12"
    | "star16"
    | "star24"
    | "star32"
    | "flowChartProcess"
    | "flowChartDecision"
    | "flowChartTerminator";

  text?: TextContent;
  geometry?: ShapeGeometry;

  raw?: RawShapeData;
}

export interface LineShape extends BaseShape {
  type: "line";

  points?: number[];
}

export interface GroupShape extends BaseShape {
  type: "group";

  children: Shape[];
}

export interface GraphicFrameShape extends BaseShape {
  type: "graphicFrame";
}

export interface Position {
  x: number;
  y: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface ShapeGeometry {
  preset?: string;
}

export interface RawShapeData {
  presetGeometry?: string;
}

export interface Placeholder {
  type?: string;
  index?: string;
}

export type Fill = SolidFill | GradientFill | PatternFill | ImageFill;

export interface SolidFill {
  type: "solid";

  color?: string;
  opacity?: number;
}

export interface GradientFill {
  type: "gradient";

  colors?: GradientStop[];
}

export interface PatternFill {
  type: "pattern";

  color?: string;
  opacity?: number;
}

export interface ImageFill {
  type: "image";

  imageRelId?: string;
}

export interface GradientStop {
  color: string;
  position: number;
  opacity?: number;
}

export interface Stroke {
  color: string;
  width: number;
  opacity?: number;
  dash?: string;
  cap?: string;
  join?: string;
}

export interface TextContent {
  paragraphs: TextParagraph[];

  /**
   * Full text with paragraphs separated by \n.
   */
  text: string;

  anchor?: string;
  wrap?: string;

  marginLeft?: number;
  marginRight?: number;
  marginTop?: number;
  marginBottom?: number;
}

export interface TextParagraph {
  text: string;
  runs: TextRun[];

  alignment?: string;
  level?: number;
  bullet?: boolean;
}

export interface TextRun {
  text: string;

  fontFamily?: string;
  fontSize?: number;

  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;

  color?: string;
  opacity?: number;
}

export interface Background {
  type: "solid" | "gradient" | "image";

  color?: string;
  opacity?: number;

  imageRelId?: string;
  imageUrl?: string;
}
