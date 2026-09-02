/**
 * Core types for presentation structure
 */

export interface Presentation {
  slides: Slide[];
  metadata?: PresentationMetadata;
}

export interface PresentationMetadata {
  title?: string;
  author?: string;
  createdAt?: Date;
  modifiedAt?: Date;
}

export interface Slide {
  id: string;
  shapes: Shape[];
  background?: Background;
}

export interface Shape {
  id: string;
  type: ShapeType;
  position: Position;
  size: Size;
}

export type ShapeType = 'text' | 'image' | 'rectangle' | 'circle' | 'line';

export interface TextShape extends Shape {
  type: 'text';
  text: string;
  style?: TextStyle;
}

export interface ImageShape extends Shape {
  type: 'image';
  src: string;
  alt?: string;
}

export interface RectangleShape extends Shape {
  type: 'rectangle';
  fill?: Fill;
  stroke?: Stroke;
}

export interface Position {
  x: number;
  y: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface TextStyle {
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  align?: 'left' | 'center' | 'right';
}

export interface Fill {
  color: string;
  opacity?: number;
}

export interface Stroke {
  color: string;
  width?: number;
}

export interface Background {
  type: 'solid' | 'gradient' | 'image';
  color?: string;
  imageUrl?: string;
}
