export interface RenderOptions {
  width?: number;
  height?: number;
  backgroundColor?: string;
  scale?: number;
  debug?: boolean; // Show shape boundaries
}

export interface TextRenderOptions {
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  align?: 'left' | 'center' | 'right';
}
