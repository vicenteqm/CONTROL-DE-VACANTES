# Control de Vacantes INVEX

Dashboard estático para GitHub Pages con KPIs, gráficas, filtros y exportación CSV.

## Publicar en GitHub Pages
1. Crea un repositorio público en GitHub, por ejemplo `control-vacantes-invex`.
2. Sube `index.html`, `styles.css`, `app.js`, `data.js` y `README.md` a la raíz del repositorio.
3. Entra a **Settings → Pages**.
4. En **Build and deployment**, selecciona **Deploy from a branch**.
5. Selecciona la rama `main` y la carpeta `/(root)`, guarda.
6. Espera a que GitHub Pages genere la URL. Normalmente será `https://TU-USUARIO.github.io/control-vacantes-invex/`.

## Reglas de negocio aplicadas
- Los KPIs de cobertura incluyen únicamente tiendas cuyo estatus en la base es `OPERANDO`.
- La plantilla requerida por tienda se calcula como `PROMOTORES REQUERIDOS + SUBLIDER`.
- El KPI de vacantes estimadas es la diferencia neta entre plantilla requerida y activos; la tabla calcula faltantes por perfil sin compensar excedentes de un perfil con faltantes del otro.
- La cobertura incluye únicamente registros de asistencia con estatus exactamente `ACTIVO` y perfil `PROMOTOR DE CREDITO` o `SUBLIDER`.
- El cruce de archivos usa la combinación de socio y determinante.
- Las tiendas con otros estatus se muestran por separado y no afectan los KPIs de cobertura.
- Se incluyen filtros por región, búsqueda y filtro por brecha; la tabla se puede exportar a CSV.
- No se exportan nombres ni datos personales de empleados.

## Actualización de datos
Cuando recibas archivos nuevos, reemplaza los datos de origen y vuelve a generar `data.js` con el proceso de transformación. El sitio no lee Excel directamente en GitHub Pages; publica el archivo agregado `data.js`.

## Consideraciones de calidad
La fuente tiene algunas tiendas OPERANDO que no encontraron registros coincidentes con personal elegible al cruzar socio + determinante. El dashboard las contabiliza con cero activos y permite identificarlas en la sección de detalle. Antes de usar los indicadores como reporte oficial, valida esas coincidencias y las claves de las tiendas.
