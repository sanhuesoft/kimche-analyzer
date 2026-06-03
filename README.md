# Kimche Analyzer

Aplicación SPA en Next.js para cargar un CSV local de observaciones escolares y visualizar métricas, gráficos y detalle filtrable en el navegador.

## Requisitos

- Node.js 20+
- npm

## Instalación

```bash
npm install
```

## Scripts

- `npm run dev`: entorno local.
- `npm run lint`: validación ESLint.
- `npm run build`: build de producción.

## Formato esperado del CSV

El archivo debe incluir estos headers:

- `Curso`
- `No. Lista`
- `Primer Apellido Estudiante`
- `Segundo Apellido Estudiante`
- `Nombre Estudiante`
- `Fecha`
- `Tipo de observación`
- `Descripción`
- `Categorización`
