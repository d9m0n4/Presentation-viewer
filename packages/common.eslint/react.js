import tseslint from 'typescript-eslint';
import baseConfig from './typescript.js';

export default tseslint.config(...baseConfig, {
  files: ['**/*.jsx', '**/*.tsx'],
  rules: {
    'react/react-in-jsx-scope': 'off',
    'react/prop-types': 'off',
  },
  settings: {
    react: {
      version: 'detect',
    },
  },
});
