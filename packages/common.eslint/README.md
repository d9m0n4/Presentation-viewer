# @presentation-viewer/common.eslint

Общие ESLint конфигурации для всех пакетов.

## Использование

### Для TypeScript пакетов

```js
// eslint.config.js
import config from '@presentation-viewer/common.eslint/typescript.js';

export default config;
```

### Для React пакетов

```js
// eslint.config.js
import config from '@presentation-viewer/common.eslint/react.js';

export default config;
```

### Базовый конфиг (JS)

```js
// eslint.config.js
import config from '@presentation-viewer/common.eslint/base.js';

export default config;
```

## Правила

- **base.js** — базовые правила для JavaScript
- **typescript.js** — правила для TypeScript проектов
- **react.js** — правила для React + TypeScript проектов
