# Presentation Viewer

Модульная библиотека для работы с PPTX файлами в браузере.

## Структура проекта

```
packages/
  core/       - Парсер и сериализатор PPTX формата (без зависимостей от фреймворков)
  renderer/   - Движок рендеринга (Canvas/SVG)
  react/      - React компоненты
  editor/     - Логика редактирования
```

## Начало работы

```bash
pnpm install
pnpm build
```

## Пакеты

### @presentation-viewer/core

Ядро библиотеки - парсинг PPTX файлов.

```typescript
import { parsePPTX } from '@presentation-viewer/core';

const file = await fetch('presentation.pptx').then(r => r.arrayBuffer());
const presentation = await parsePPTX(file);
console.log(presentation.slides);
```

## Roadmap

- [x] Monorepo setup
- [x] Core parser базовая структура
- [ ] Парсинг текстовых элементов
- [ ] Парсинг изображений
- [ ] Парсинг фигур
- [ ] Renderer пакет
- [ ] React компоненты
- [ ] Editor функционал
# Presentation-viewer
