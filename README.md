# Batería de indicadores DANE · Frontend Putumayo

Este paquete contiene únicamente el frontend Vite del dashboard visual. La base de indicadores está integrada en `client/src/data/indicadores.json`; la interfaz la importa directamente al iniciar, sin conexión a una base de datos y sin selección manual de un archivo Excel.

## Requisitos

- Node.js 18 o superior
- npm 9 o pnpm 8 o superior

## Uso local

Desde esta carpeta, instala las dependencias y ejecuta el servidor de desarrollo:

```bash
npm install
npm run dev
```

Vite mostrará la dirección local en la terminal. Para generar una compilación de producción:

```bash
npm run build
npm run preview
```

El proyecto conserva la estructura solicitada: `client/src/`, `client/index.html`, `vite.config.ts` y `client/src/data/indicadores.json`. Los recursos visuales locales están en `client/public/assets/`.
