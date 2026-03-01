# Espacio_En_3D 🌌📐

Este repositorio contiene el sistema de gestión y simulación de **Entornos Tridimensionales**, desarrollado en Visual Basic 6.0.

## 🚀 Descripción

`Espacio_En_3D` es un proyecto diseñado para establecer y manipular un espacio de coordenadas tridimensionales. A diferencia de un simple motor de renderizado, este sistema se centra en la lógica de la "escena": la definición del origen, la orientación de los ejes cartesianos y la ubicación relativa de múltiples objetos dentro de un universo virtual.

Es la base estructural necesaria para situar complejos modelos matemáticos (como la esfera de Riemann o teseractos) dentro de un contexto espacial coherente y navegable.

## 🛠️ Especificaciones Técnicas

- **Lenguaje:** Visual Basic 6.0.
- **Arquitectura:** Gestor de escena basado en vectores de posición y orientación.
- **Funcionalidades Clave:**
  - **Sistemas de Coordenadas:** Gestión de ejes $X, Y, Z$ con soporte para rejillas de referencia (*grids*).
  - **Transformaciones de Espacio:** Lógica para trasladar el punto de origen y rotar el universo completo.
  - **Manejo de Objetos:** Capacidad para insertar y rastrear múltiples entidades dentro del mismo espacio 3D.
  - **Cálculo Vectorial:** Implementación de operaciones de producto punto, producto cruz y normalización de vectores.

## 📂 Estructura del Repositorio

Tras la reestructuración profesional:
- `/src`: Código fuente del gestor de espacio (.vbp, .frm, .bas, .cls).
- `/scripts`: Rutinas para la generación de entornos y mallas de referencia.
- `/docs`: Documentación sobre geometría del espacio y álgebra vectorial aplicada.

## ⚙️ Instalación y Uso

1. Clona el repositorio:
   ```bash
   git clone [https://github.com/MiguelQuinteiro/Espacio_En_3D.git](https://github.com/MiguelQuinteiro/Espacio_En_3D.git)
   