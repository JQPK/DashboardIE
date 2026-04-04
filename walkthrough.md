# Dashboard Académico Listo

El Dashboard ha sido construido exitosamente utilizando tecnologías web nativas (HTML5 puro, CSS moderno y Vanilla Javascript con librerías por CDN) como respuesta directa a la incompatibilidad con Node.js en su sistema, garantizando que puedas usar y desplegar la herramienta de inmediato sin tener que instalar nada.

## Arquitectura y Lógica Resolutiva
* **Diseño UI:** Se implementó una paleta basada en blanco (fondo), celeste (Logrado), amarillo (Proceso) y rojo (Inicio/Alerta) según tu requerimiento de diseño moderno con uso de sombras suaves.
* **Procesador Engine (`app.js`):** La app espera un CSV subido por drag and drop. 
* **Reglas de Procesamiento:** 
    * El sistema normaliza valores reconociendo texto como "INICIO", "PROCESO", "LOGRADO", pero es resiliente para interpretar escalas numéricas puras.
    * Agrupa por alumno y calcula *el promedio matemático estricto por curso* eliminando casillas vacías automáticamente.
    * Muestra mensajes de validación si faltan columnas esenciales como `NOMBRES` o `AREA` o notas vacías.
* **Detectando Debilidades Dinámicas:** 
    * Si la app detecta que al menos el **30%** de los alumnos evaluados por cada curso tiene un estado general de *Inicio o Proceso*, lanza una alerta ("Debilidad"). Las métricas por curso se reportan automáticamente en las tarjetas de insights.

## Interfaz Interesante
* **KPIs dinámicos:** Muestra la totalidad de alumnos únicos, cursos evaluados, porcentaje general de logro y cursos en riesgo.
* **Filtros combinables:** Arriba del dashboard podrás aislar un área específica (ej. "Religión") o un Docente en particular y los gráficos se reacomodarán a ese contexto.

Para visualizar el resultado, solo debes hacer doble clic en el archivo `index.html` que está en tu carpeta del proyecto.
