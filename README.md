# Control de Vacantes INVEX

Dashboard estático para GitHub Pages, con KPIs, gráficas, búsqueda, filtros por región, exportación CSV y resumen diario de códigos de asistencia.

## Cómo funciona la actualización diaria

El sitio publicado en GitHub Pages no puede leer directamente una carpeta privada de tu computadora. Para actualizarlo sin volver a editar el código, este proyecto incluye un proceso local gratuito que:

1. Busca el Excel vigente de base de tiendas y el Excel vigente de asistencia en las carpetas que selecciones.
2. Detecta las columnas de días de asistencia por sus encabezados numéricos (1–31), así que no necesita cambios cuando se agrega el día 9, 10, etc.
3. Cruza socio + determinante y calcula de nuevo los KPIs.
4. Actualiza `data.js` con datos agregados. No copia nombres, nóminas, teléfonos ni registros individuales de empleados.
5. Si el proyecto está dentro de un repositorio Git clonado y configurado, crea un commit únicamente de `data.js` y ejecuta `git push`.
6. Escribe el resultado en `logs/actualizacion.log`.

## Preparación inicial en Windows

**Importante:** el proceso automático debe instalarse en una carpeta local vinculada al repositorio de GitHub (un clon del repositorio), no en la carpeta temporal donde se descargó el ZIP.

1. Instala Python 3 para Windows desde https://www.python.org/downloads/windows/. Si el instalador lo ofrece, activa la opción para agregar Python al PATH.
2. Instala Git for Windows desde https://git-scm.com/download/win y autentícate con GitHub. Para publicar un repositorio existente también puedes usar GitHub Desktop, pero la orden `git` debe estar disponible en la terminal.
3. Clona tu repositorio `control-vacantes-invex` a tu computadora. Si aún no existe, crea/publica el repositorio y copia estos archivos a la raíz del repositorio local, sin borrar su carpeta oculta `.git`.
4. Copia los archivos de este proyecto a la raíz del repositorio local. No copies los Excel a la carpeta del repositorio ni borres la carpeta oculta `.git`.
5. Abre GitHub Desktop, revisa que aparezcan los archivos nuevos/modificados, crea un commit inicial del proyecto y pulsa **Push origin**. Esto publica la nueva interfaz y los scripts; las ejecuciones diarias posteriores solo publican `data.js`.
6. Haz doble clic en `instalar.bat` para instalar la dependencia gratuita que lee los Excel.
7. Haz doble clic en `configurar_carpetas.bat`. Selecciona la carpeta donde está el Excel de la base y después la carpeta donde está el Excel de asistencia. Se guarda un `config.json` local que no se publica en GitHub.
8. Abre `actualizar.bat` para hacer la primera prueba. La primera publicación con Git puede requerir autenticación. Si el proceso indica que falta la sesión, abre GitHub Desktop, termina el inicio de sesión y prueba de nuevo.
9. Comprueba que el dashboard actualizado esté publicado y que la ejecución termine sin errores en `logs/actualizacion.log`.

### Cómo identifica los archivos de Excel

La configuración busca el archivo `.xlsx` modificado más recientemente que coincida con estos patrones:

- Base de tiendas: `*BASE*TIENDAS*.xlsx`
- Asistencia: `*ASISTENCIA*.xlsx`

Por eso, conserva en cada carpeta el archivo de trabajo vigente o evita dejar copias antiguas que también coincidan con el patrón y se modifiquen después. Si los nombres de tus archivos son distintos, puedes cambiar `base_pattern` o `attendance_pattern` en `config.json`.

## Programar el proceso diario

1. Primero ejecuta `actualizar.bat` manualmente y confirma que termina sin errores y que GitHub recibe el cambio.
2. Ejecuta `programar_actualizacion.bat` y escribe la hora de ejecución (formato de 24 horas, por ejemplo `09:00`). Esto crea una tarea diaria en el Programador de tareas de Windows.
3. Deja la computadora encendida, conectada a internet y con acceso a las carpetas cuando se ejecute la tarea. Los Excel deben estar guardados y sincronizados antes de la hora programada.
4. Revisa `logs/actualizacion.log` si alguna ejecución falla. También puedes abrir `actualizar.bat` en cualquier momento para lanzar una actualización manual.

El proceso evita crear commits repetidos cuando los datos no han cambiado. Si el equipo está apagado, no tiene conexión o GitHub no acepta la publicación, el sitio conservará los datos publicados anteriormente hasta que una ejecución termine correctamente.

## Reglas de negocio

- Los KPIs de cobertura consideran únicamente tiendas con estatus `OPERANDO`.
- La plantilla requerida es `PROMOTORES REQUERIDOS + SUBLIDER`.
- Para cobertura solo cuentan filas con estatus exactamente `ACTIVO` y perfil `PROMOTOR DE CREDITO` o `SUBLIDER`.
- El cruce de asignación utiliza la combinación de socio y determinante.
- La gráfica diaria cuenta códigos en las columnas numéricas de día para ese mismo personal elegible asignado a tiendas `OPERANDO`.
- Códigos identificados: `A` asistencia, `F` falta, `R` retardo, `I` incapacidad, `V` vacaciones y `D` descanso (observado en el archivo fuente). Cualquier código corto no catalogado se agrupa en `Otros códigos` sin asumir su significado.
- Los conteos de la gráfica de asistencia son por empleado y día: una persona puede sumar una observación en distintos días. No son personas únicas del mes.
- Las tiendas con otros estatus se muestran por separado y no afectan los KPIs de cobertura.
- La exportación de la tabla contiene datos agregados por tienda, no registros individuales de empleados.

## Publicar o actualizar GitHub Pages

En el repositorio, entra a **Settings → Pages** y selecciona **Deploy from a branch**, rama `main` y carpeta `/(root)`. La dirección suele tener este formato: `https://TU-USUARIO.github.io/control-vacantes-invex/`.

GitHub Pages sirve el dashboard de forma gratuita para repositorios públicos. La actualización automática descrita aquí también usa herramientas gratuitas, pero depende de que la computadora local esté encendida, conectada y pueda acceder a las carpetas de origen.

## Protección de datos

No publiques los archivos originales de Excel si contienen nombres, números de nómina, teléfonos u otros datos personales. Este proyecto solo genera y publica el archivo agregado `data.js`; `config.json` y los logs se mantienen locales mediante `.gitignore`.
