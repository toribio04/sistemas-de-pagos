# Banco Robles - Sistema de Pagos

Sistema de pagos para Banco Robles que permite registrar transacciones y guardarlas en un archivo Excel.

## Características

- Formulario de pago con validación
- Guardado automático en archivo Excel
- Diseño moderno y responsive
- Interfaz similar al diseño original

## Requisitos

- Node.js (v18 o superior)
- npm o yarn

## Instalación

1. Instalar dependencias:
```bash
npm install
```

## Ejecutar la aplicación

```bash
npm start
```

La aplicación se abrirá en `http://localhost:4200`

## Funcionalidades

- **Formulario de pago**: Permite ingresar nombre, apellido, empresa, importe y forma de pago
- **Guardado en Excel**: Cada pago se guarda automáticamente en un archivo Excel que se descarga
- **Persistencia**: Los datos se guardan en localStorage para mantener un historial

## Estructura del proyecto

```
src/
├── app/
│   ├── components/
│   │   └── payment-form/
│   │       ├── payment-form.component.ts
│   │       ├── payment-form.component.html
│   │       └── payment-form.component.css
│   ├── services/
│   │   └── excel.service.ts
│   └── app.module.ts
└── styles.css
```

## Tecnologías utilizadas

- Angular 17
- TypeScript
- XLSX (para manejo de archivos Excel)
- CSS3

## Notas

- Los datos se guardan en localStorage y se descargan en un archivo Excel cada vez que se confirma un pago
- El archivo Excel se genera con el nombre `pagos_banco_robles.xlsx`
- Todos los campos son obligatorios excepto que se indique lo contrario







