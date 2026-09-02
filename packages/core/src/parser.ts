import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import type { Presentation, Slide } from './types.js';

export class PPTXParser {
  private xmlParser: XMLParser;

  constructor() {
    this.xmlParser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
    });
  }

  async parse(arrayBuffer: ArrayBuffer): Promise<Presentation> {
    const zip = await JSZip.loadAsync(arrayBuffer);

    const slides = await this.parseSlides(zip);
    const metadata = await this.parseMetadata(zip);

    return {
      slides,
      metadata,
    };
  }

  private async parseSlides(zip: JSZip): Promise<Slide[]> {
    const slides: Slide[] = [];
    const slideFiles = Object.keys(zip.files).filter(name =>
      name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
    );

    slideFiles.sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)\.xml/)?.[1] || '0');
      const numB = parseInt(b.match(/slide(\d+)\.xml/)?.[1] || '0');
      return numA - numB;
    });

    for (const slideFile of slideFiles) {
      const content = await zip.file(slideFile)?.async('string');
      if (content) {
        const slide = this.parseSlide(content, slideFile);
        slides.push(slide);
      }
    }

    return slides;
  }

  private parseSlide(xmlContent: string, filename: string): Slide {
    const data = this.xmlParser.parse(xmlContent);

    // Детальное логирование структуры
    const cSld = data?.['p:sld']?.['p:cSld'];
    const spTree = cSld?.['p:spTree'];

    console.log(`Parsing ${filename}:`);
    console.log('- Available elements in spTree:', spTree ? Object.keys(spTree) : 'no spTree');
    console.log('- p:sp (shapes):', spTree?.['p:sp'] ? (Array.isArray(spTree['p:sp']) ? spTree['p:sp'].length : 1) : 0);
    console.log('- p:pic (images):', spTree?.['p:pic'] ? (Array.isArray(spTree['p:pic']) ? spTree['p:pic'].length : 1) : 0);
    console.log('- Background (p:bg):', cSld?.['p:bg']);

    const shapes = this.extractShapes(data);
    const background = this.extractBackground(data);

    return {
      id: filename.replace('ppt/slides/', '').replace('.xml', ''),
      shapes,
      background,
    };
  }

  private extractShapes(slideData: any): any[] {
    const shapes: any[] = [];

    try {
      const spTree = slideData?.['p:sld']?.['p:cSld']?.['p:spTree'];
      if (!spTree) return shapes;

      // Извлекаем текстовые фигуры и прямоугольники (p:sp)
      const textShapes = spTree['p:sp'];
      if (textShapes) {
        const shapeArray = Array.isArray(textShapes) ? textShapes : [textShapes];
        for (const element of shapeArray) {
          const shape = this.parseTextShape(element);
          if (shape) {
            console.log('Parsed shape:', shape);
            shapes.push(shape);
          }
        }
      }

      // Извлекаем изображения (p:pic)
      const pictures = spTree['p:pic'];
      if (pictures) {
        const picArray = Array.isArray(pictures) ? pictures : [pictures];
        for (const element of picArray) {
          const shape = this.parseImageShape(element);
          if (shape) {
            console.log('Parsed image:', shape);
            shapes.push(shape);
          }
        }
      }

      // Извлекаем соединительные линии (p:cxnSp)
      const connectors = spTree['p:cxnSp'];
      if (connectors) {
        const connArray = Array.isArray(connectors) ? connectors : [connectors];
        for (const element of connArray) {
          const shape = this.parseConnectorShape(element);
          if (shape) {
            console.log('Parsed connector:', shape);
            shapes.push(shape);
          }
        }
      }

    } catch (error) {
      console.warn('Error parsing shapes:', error);
    }

    return shapes;
  }

  private parseTextShape(shapeData: any): any | null {
    try {
      // Извлекаем позицию и размер
      const spPr = shapeData['p:spPr'];
      const xfrm = spPr?.['a:xfrm'];
      const off = xfrm?.['a:off'];
      const ext = xfrm?.['a:ext'];

      const position = {
        x: parseInt(off?.['@_x'] || '0'),
        y: parseInt(off?.['@_y'] || '0'),
      };
      const size = {
        width: parseInt(ext?.['@_cx'] || '0'),
        height: parseInt(ext?.['@_cy'] || '0'),
      };

      const shapeId = shapeData['p:nvSpPr']?.['p:cNvPr']?.['@_id'] || 'unknown';

      // Извлекаем текст
      const txBody = shapeData['p:txBody'];
      const text = txBody ? this.extractText(txBody) : '';

      // Если есть текст, возвращаем текстовую фигуру
      if (text) {
        return {
          id: shapeId,
          type: 'text',
          text,
          position,
          size,
        };
      }

      // Если нет текста, проверяем есть ли заливка — это прямоугольник
      const style = shapeData['p:style'];
      let fill = this.extractFillFromSpPr(spPr); // Только явная заливка, без дефолта

      // Если заливки нет, пробуем взять из style reference
      if (!fill && style) {
        console.log('Shape has p:style:', style);
        const fillRef = style['a:fillRef'];
        if (fillRef) {
          console.log('Found fillRef:', fillRef);
          const color = this.extractColor(fillRef);
          if (color) {
            fill = { color, opacity: 1 };
          }
        }
      }

      // Если все еще нет заливки, используем дефолт для фигур с геометрией
      if (!fill) {
        const prstGeom = spPr?.['a:prstGeom'];
        if (prstGeom) {
          console.log('Using default fill for shape with prstGeom');
          fill = {
            color: '#4472c4',
            opacity: 1,
          };
        }
      }

      const stroke = this.extractStroke(spPr);

      if (fill || stroke) {
        console.log(`Rectangle shape ${shapeId}:`, {
          position,
          size,
          fill,
          stroke,
          hasStyle: !!style,
          spPrKeys: spPr ? Object.keys(spPr) : 'no spPr'
        });
        return {
          id: shapeId,
          type: 'rectangle',
          position,
          size,
          fill,
          stroke,
        };
      }

      // Логируем пропущенные фигуры
      console.log(`Skipped shape ${shapeId}: no text, no fill, no stroke`, { position, size });

      return null;
    } catch (error) {
      console.warn('Error parsing text shape:', error);
      return null;
    }
  }

  private parseImageShape(picData: any): any | null {
    try {
      // Извлекаем позицию и размер
      const spPr = picData['p:spPr'];
      const xfrm = spPr?.['a:xfrm'];
      const off = xfrm?.['a:off'];
      const ext = xfrm?.['a:ext'];

      // Извлекаем ссылку на изображение
      const blip = picData['p:blipFill']?.['a:blip'];
      const imageRel = blip?.['@_r:embed'];

      // TODO: разрешить ссылку на реальный файл изображения из архива
      // Пока используем placeholder
      const src = imageRel ? `data:image/png;base64,placeholder_${imageRel}` : '';

      return {
        id: picData['p:nvPicPr']?.['p:cNvPr']?.['@_id'] || 'unknown',
        type: 'image',
        src,
        position: {
          x: parseInt(off?.['@_x'] || '0'),
          y: parseInt(off?.['@_y'] || '0'),
        },
        size: {
          width: parseInt(ext?.['@_cx'] || '0'),
          height: parseInt(ext?.['@_cy'] || '0'),
        },
      };
    } catch (error) {
      console.warn('Error parsing image shape:', error);
      return null;
    }
  }

  private parseConnectorShape(connData: any): any | null {
    try {
      // Извлекаем позицию и размер
      const spPr = connData['p:spPr'];
      const xfrm = spPr?.['a:xfrm'];
      const off = xfrm?.['a:off'];
      const ext = xfrm?.['a:ext'];

      const position = {
        x: parseInt(off?.['@_x'] || '0'),
        y: parseInt(off?.['@_y'] || '0'),
      };
      const size = {
        width: parseInt(ext?.['@_cx'] || '0'),
        height: parseInt(ext?.['@_cy'] || '0'),
      };

      const stroke = this.extractStroke(spPr);

      return {
        id: connData['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id'] || 'unknown',
        type: 'line',
        position,
        size,
        stroke,
      };
    } catch (error) {
      console.warn('Error parsing connector shape:', error);
      return null;
    }
  }

  private extractText(txBody: any): string {
    try {
      const paragraphs = txBody['a:p'];
      if (!paragraphs) return '';

      const pArray = Array.isArray(paragraphs) ? paragraphs : [paragraphs];
      const textParts: string[] = [];

      for (const p of pArray) {
        const runs = p['a:r'];
        if (!runs) continue;

        const runArray = Array.isArray(runs) ? runs : [runs];
        for (const run of runArray) {
          const text = run['a:t'];
          if (text) {
            textParts.push(typeof text === 'string' ? text : text.toString());
          }
        }
      }

      return textParts.join(' ').trim();
    } catch (error) {
      console.warn('Error extracting text:', error);
      return '';
    }
  }

  private extractBackground(slideData: any): any {
    try {
      const cSld = slideData?.['p:sld']?.['p:cSld'];
      const bg = cSld?.['p:bg'];

      if (!bg) return undefined;

      // Проверяем заливку фона
      const bgPr = bg['p:bgPr'];
      if (bgPr) {
        // Сплошная заливка
        const solidFill = bgPr['a:solidFill'];
        if (solidFill) {
          const color = this.extractColor(solidFill);
          if (color) {
            return {
              type: 'solid',
              color,
            };
          }
        }

        // Градиентная заливка
        const gradFill = bgPr['a:gradFill'];
        if (gradFill) {
          // TODO: парсить градиенты
          return {
            type: 'gradient',
          };
        }

        // Заливка изображением
        const blipFill = bgPr['a:blipFill'];
        if (blipFill) {
          // TODO: парсить изображения фона
          return {
            type: 'image',
          };
        }
      }

      return undefined;
    } catch (error) {
      console.warn('Error extracting background:', error);
      return undefined;
    }
  }

  private extractColor(fillData: any): string | undefined {
    try {
      // sRGB цвет
      const srgbClr = fillData['a:srgbClr'];
      if (srgbClr) {
        const val = srgbClr['@_val'];
        if (val) {
          return `#${val}`;
        }
      }

      // Схемный цвет (theme color)
      const schemeClr = fillData['a:schemeClr'];
      if (schemeClr) {
        const val = schemeClr['@_val'];
        // Маппинг базовых цветов схемы
        const colorMap: Record<string, string> = {
          'bg1': '#ffffff',
          'bg2': '#f0f0f0',
          'tx1': '#000000',
          'tx2': '#444444',
          'accent1': '#4472c4',
          'accent2': '#ed7d31',
          'accent3': '#a5a5a5',
          'accent4': '#ffc000',
          'accent5': '#5b9bd5',
          'accent6': '#70ad47',
          'lt1': '#ffffff',
          'lt2': '#e7e6e6',
          'dk1': '#000000',
          'dk2': '#44546a',
          'folHlink': '#0563c1',
          'hlink': '#467886',
        };
        const mappedColor = colorMap[val] || '#4472c4';

        // Проверяем модификаторы цвета
        console.log(`Scheme color ${val}, raw schemeClr:`, schemeClr);

        console.log(`Mapped scheme color ${val} -> ${mappedColor}`);
        return mappedColor;
      }

      // Preset color
      const prstClr = fillData['a:prstClr'];
      if (prstClr) {
        const val = prstClr['@_val'];
        // Некоторые preset цвета
        const presetMap: Record<string, string> = {
          'black': '#000000',
          'white': '#ffffff',
          'red': '#ff0000',
          'green': '#00ff00',
          'blue': '#0000ff',
          'yellow': '#ffff00',
          'cyan': '#00ffff',
          'magenta': '#ff00ff',
        };
        return presetMap[val] || '#000000';
      }

      return undefined;
    } catch (error) {
      console.warn('Error extracting color:', error);
      return undefined;
    }
  }

  private extractFill(spPr: any): any {
    try {
      // Сплошная заливка
      const solidFill = spPr?.['a:solidFill'];
      if (solidFill) {
        const color = this.extractColor(solidFill);
        if (color) {
          return {
            color,
            opacity: 1,
          };
        }
      }

      // Градиентная заливка - берем первый цвет
      const gradFill = spPr?.['a:gradFill'];
      if (gradFill) {
        const gsLst = gradFill['a:gsLst'];
        const gs = gsLst?.['a:gs'];
        if (gs) {
          const firstStop = Array.isArray(gs) ? gs[0] : gs;
          const color = this.extractColor(firstStop);
          if (color) {
            return {
              color,
              opacity: 1,
            };
          }
        }
      }

      // Нет заливки (явно указано)
      const noFill = spPr?.['a:noFill'];
      if (noFill !== undefined) {
        return undefined;
      }

      // Если нет явной заливки, проверяем есть ли preset geometry
      const prstGeom = spPr?.['a:prstGeom'];
      if (prstGeom) {
        console.log('Shape has prstGeom but no explicit fill, using default');
        // У фигуры есть геометрия, возвращаем дефолтную заливку
        return {
          color: '#4472c4', // Accent1 - стандартный цвет PowerPoint
          opacity: 1,
        };
      }

      return undefined;
    } catch (error) {
      console.warn('Error extracting fill:', error);
      return undefined;
    }
  }

  private extractFillFromSpPr(spPr: any): any {
    try {
      // Сплошная заливка
      const solidFill = spPr?.['a:solidFill'];
      if (solidFill) {
        const color = this.extractColor(solidFill);
        if (color) {
          return {
            color,
            opacity: 1,
          };
        }
      }

      // Градиентная заливка - берем первый цвет
      const gradFill = spPr?.['a:gradFill'];
      if (gradFill) {
        const gsLst = gradFill['a:gsLst'];
        const gs = gsLst?.['a:gs'];
        if (gs) {
          const firstStop = Array.isArray(gs) ? gs[0] : gs;
          const color = this.extractColor(firstStop);
          if (color) {
            return {
              color,
              opacity: 1,
            };
          }
        }
      }

      // Нет заливки (явно указано)
      const noFill = spPr?.['a:noFill'];
      if (noFill !== undefined) {
        return undefined;
      }

      // Не возвращаем дефолт - только явная заливка
      return undefined;
    } catch (error) {
      console.warn('Error extracting fill from spPr:', error);
      return undefined;
    }
  }

  private extractStroke(spPr: any): any {
    try {
      const ln = spPr?.['a:ln'];
      if (!ln) return undefined;

      const width = ln['@_w'] ? parseInt(ln['@_w']) / 12700 : 1; // EMU to points
      const solidFill = ln['a:solidFill'];

      if (solidFill) {
        const color = this.extractColor(solidFill);
        if (color) {
          return {
            color,
            width,
          };
        }
      }

      return undefined;
    } catch (error) {
      console.warn('Error extracting stroke:', error);
      return undefined;
    }
  }

  private async parseMetadata(zip: JSZip): Promise<any> {
    const coreFile = zip.file('docProps/core.xml');
    if (!coreFile) return {};

    const content = await coreFile.async('string');
    const data = this.xmlParser.parse(content);

    return {
      title: data?.['cp:coreProperties']?.['dc:title'],
      author: data?.['cp:coreProperties']?.['dc:creator'],
    };
  }
}

export async function parsePPTX(file: File | ArrayBuffer): Promise<Presentation> {
  const parser = new PPTXParser();
  const buffer = file instanceof File ? await file.arrayBuffer() : file;
  return parser.parse(buffer);
}
