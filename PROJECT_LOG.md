# Project Log

Este archivo resume el estado del proyecto y registra los cambios importantes para poder retomar trabajo sin reconstruir todo el contexto desde cero.

## Como usar este archivo

- Actualizar este log cada vez que se haga una tanda relevante de cambios.
- Priorizar decisiones, comportamiento actual, riesgos y tareas pendientes.
- Evitar detalle excesivo de implementacion cuando no aporte.

## Estado general

- Proyecto: Plan de Visitas Maxiseguridad
- Produccion: `https://visitasms.online`
- Stack principal:
  - Frontend estatico en `public/`
  - Backend Node.js en `backend/server.js`
  - Base PostgreSQL en produccion
  - PM2 + Nginx en el VPS

## Reglas de negocio relevantes

- `Configuracion` solo debe estar visible y accesible para:
  - Nombre: `Nicolas Beguelman`
  - Email: `nicolas@maxiseguridad.com`
- El rol `Ejecutivo` solo puede ver y operar sus propias companias y sus reuniones.
- El resto de los roles puede ver todo.
- CRM / Pipeline deben permanecer ocultos en la rama principal por ahora.

## Comportamiento funcional actual

- Reuniones:
  - Se pueden crear, editar y eliminar.
  - La eliminacion es logica: las reuniones van a `Papelera` y pueden recuperarse.
- Configuracion:
  - Organizada en sub-vistas / modulos.
  - Incluye catalogos, importaciones, papelera y auditoria.
- Auditoria:
  - Debe registrar movimientos relevantes de usuarios.
- Importaciones:
  - Hay importacion desde Excel para clientes.
  - La importacion de usuarios esta ubicada dentro de Gestion de usuarios.
- Visitas:
  - Existen vistas Lista, Calendario, Grilla y Estadisticas.
- Estadisticas:
  - Hay una vista con filtros por periodo.
  - Muestra metricas por usuario y por tipo de visita.

## Consideraciones tecnicas importantes

- No resembrar usuarios demo ni resetear tipos de reunion en cada deploy.
- Evitar cualquier cambio que pise datos cargados manualmente en produccion.
- Antes de deployar a produccion, conviene hacer backup de PostgreSQL.
- El VPS requiere SSH interactivo con password; desde Codex no siempre se puede completar ese paso.

## Deploy seguro a produccion

Comandos sugeridos en el VPS:

```bash
mkdir -p /var/backups/plan-de-visitas
sudo -u postgres pg_dump -Fc planvisitas > /var/backups/plan-de-visitas/planvisitas-$(date +%F-%H%M%S).dump
cd /var/www/plan-de-visitas
git pull
pm2 restart plan-de-visitas
pm2 status
curl -I https://visitasms.online
```

## Bitacora de cambios

### 2026-04-14

- Se crea este archivo para conservar contexto operativo del proyecto entre sesiones.
- Queda establecido que cada nueva tanda de cambios relevantes debe reflejarse aca.

### 2026-04-15

- La `Función del contacto` en reuniones deja de ser texto libre y pasa a ser un catálogo administrable.
- Se agregan funciones de contacto predefinidas:
  - `Usuario`
  - `Comprador`
  - `Nuevo contacto`
- Estas opciones se pueden crear, editar y eliminar desde `Configuración`.
- Se agrega importador de `Sucursales desde Excel` dentro de `Configuración > Importaciones`.
- La importación de sucursales asocia cada fila a una compañía existente por nombre.
- Se agregan `Reglas de visita` editables en la ficha de compañía y sucursal.
- Cada regla guarda:
  - periodicidad en días
  - función del contacto a visitar
  - objetivo libre
- Las reglas quedan persistidas para seguir luego con lógica automática basada en ellas.
- Las reglas ya disparan reuniones automáticas:
  - tipo `Visita Comercial`
  - motivo `Acercamiento al cliente`
  - modalidad `Presencial`
  - participante: ejecutivo de la cuenta
  - estado `Agendada`
  - sin oportunidad vinculada
- Cuando una reunión automática se completa, el sistema genera la siguiente según la periodicidad de la regla.
- En `Estadísticas` se suma un resumen específico de visitas automáticas por regla:
  - total a ejecutar
  - total ejecutadas
  - porcentaje de cumplimiento

### 2026-04-17

- Se agregan campos preparatorios para futura integración con Bitrix en entidades ya existentes.
- Usuarios:
  - nuevo campo `ID usuario Bitrix`
  - visible y editable desde `Configuración > Gestión de usuarios`
- Compañías:
  - nuevo campo `Bitrix Lead ID`
  - nuevo campo `Bitrix Company ID`
  - visibles y editables en la ficha de compañía
- Sucursales:
  - no usan estos campos por ahora
  - los inputs quedan ocultos cuando se edita o crea una sucursal para no mezclar datos
- Objetivo de este paso:
  - dejar cargables los identificadores externos antes de avanzar con automatizaciones o sincronización con Bitrix
- Se agrega en `Configuración > Control` una prueba manual de Bitrix:
  - permite pegar un webhook o URL de `user.get`
  - consulta usuarios reales de Bitrix desde el backend
  - muestra listado para validar IDs, mails y cargos antes de mapear integraciones
- Decisión tomada:
  - por ahora no se guarda el webhook en base ni en código
  - el objetivo en esta etapa es validar acceso y estructura de datos sin comprometer credenciales
- La integración Bitrix evoluciona con tres capacidades nuevas en `Configuración > Control > Bitrix`:
  - mapeo automático de usuarios de la app contra usuarios Bitrix usando `bitrixUserId`
  - mapeo automático de compañías de la app contra `crm.company` o `crm.lead` usando `Bitrix Company ID` y `Bitrix Lead ID`
  - creación de una tarea de prueba en Bitrix desde la app con responsable seleccionable
- El webhook sigue siendo temporal y manual:
  - se pega en la pantalla al momento de probar
  - no se persiste todavía en base de datos ni en archivos
- Cambio posterior:
  - el webhook de Bitrix pasa a guardarse en base de datos (`app_settings`)
  - al entrar a `Configuración > Control > Bitrix` se carga automáticamente
  - también se guarda automáticamente al actualizar la vista de Bitrix
- Login:
  - se agrega la opción `Recordarme en este dispositivo`
  - si se activa, la sesión queda persistida por 30 días
  - si no se activa, la sesión sigue siendo de navegador
- Reglas de visita:
  - el campo `objetivo` deja de ser texto libre
  - pasa a ser un selector cerrado con estas opciones:
    - `Desarrollo de cuentas`
    - `Nuevo negocio`

### 2026-04-20

- En `Estadísticas de visitas` se agrega un `semáforo de cumplimiento` para visitas automáticas creadas por reglas.
- El semáforo se calcula por `regla activa`:
  - `Blanco`: todavía no tuvo ninguna visita realizada y no está vencida
  - `Verde`: la regla viene cumplida y al día
  - `Amarillo`: la visita pendiente se atrasó hasta 10 días
  - `Rojo`: la visita pendiente supera 10 días de atraso
- La vista de estadísticas ahora permite segmentar también por `Ejecutivo comercial`.
- Ese filtro impacta tanto en las tablas generales como en el nuevo semáforo.
- En la ficha de visitas, el campo visible `Status sobre el servicio` pasa a llamarse `Reclamos`.
- Por ahora el cambio es solo de nombre en la UI; no se altera la estructura interna del dato ni se integra todavía con Bitrix.
- Luego se agrega un segundo bloque independiente en la visita llamado `Status del servicio`.
- Quedan separados conceptualmente:
  - `Status del servicio`: estado general, desvíos y feedback del servicio
  - `Reclamos`: reclamos específicos, pensado a futuro para integración con Bitrix
- Se incorpora automatización para `Reclamos`:
  - al guardar una visita con texto en `Reclamos`, el sistema crea o actualiza una tarea en Bitrix
  - grupo fijo Bitrix: `46`
  - responsables posibles desde desplegable:
    - `Gonzalo Garcia (44)`
    - `Guillermo Saralegui (10542)`
    - `Guido Avalo Ceci (1720)`
  - título: `Reclamo - [Cliente] - [Fecha de la visita]`
  - descripción: contenido textual del bloque `Reclamos`
  - vencimiento: fecha actual + 7 días
  - se guarda el `taskId` de Bitrix en la reunión para evitar duplicados y actualizar la misma tarea en futuras ediciones
- Se mejora el mapeo manual con Bitrix:
  - usuarios: el vínculo pasa a elegirse desde un buscador con usuarios activos de Bitrix, en vez de cargar el ID a mano
  - compañías: el vínculo con Bitrix pasa a elegirse desde un buscador de compañías de Bitrix
  - sucursales: también guardan su propio `bitrixCompanyId` y lo eligen desde ese mismo buscador
- Se reemplaza la implementación de esos buscadores Bitrix basada en `datalist` por un dropdown propio con filtro en vivo.
- Motivo:
  - en el navegador embebido local no se abría ni mostraba resultados de manera consistente
- Resultado:
  - al enfocar el campo ya se listan opciones disponibles
  - al escribir, los resultados se filtran en vivo
  - al hacer click en una opción se guarda correctamente el ID oculto correspondiente
- Ajuste posterior sobre Bitrix:
  - el campo `ACTIVE` de `user.get` puede venir como booleano y no solo como `Y/N`
  - se corrige esa lectura para que los usuarios activos aparezcan efectivamente en el buscador
- Nuevo ajuste sobre usuarios Bitrix:
  - `user.get` en este portal devuelve los usuarios con `ACTIVE=false` aunque existan activos
  - para poblar el selector de usuarios se pasa a usar `user.search` con filtro `ACTIVE=true`
- También se mejora el feedback del directorio Bitrix:
  - si el webhook no tiene permisos CRM para leer compañías, la UI ahora lo informa explícitamente
  - deja de verse como si simplemente “no hubiera compañías”
- Mejora en el selector de compañías Bitrix:
  - la precarga inicial no alcanza para cubrir todo el padrón cuando hay miles de compañías
  - se agrega búsqueda remota en vivo contra Bitrix por nombre (`crm.company.list` con filtro por `TITLE`)
  - esto permite encontrar compañías como `OBRAS METALICAS S.A.` aunque no estén en las primeras páginas cargadas al iniciar
- Ajuste en reclamos Bitrix:
  - al crear una tarea de `Reclamo`, el sistema intenta usar como creador al `ejecutivo de la compañía` donde se hizo la visita
  - para eso toma el `bitrixUserId` mapeado en el usuario ejecutivo local
  - si la cuenta no tiene ese mapeo, la creación cae al comportamiento estándar de Bitrix para no bloquear el reclamo
- Se agrega un switch de módulo para `Reglas automáticas de visitas`.
- El switch vive en `Configuración` y permite:
  - mostrar u ocultar el editor de reglas en clientes y sucursales
  - mostrar u ocultar las estadísticas propias de reglas
  - activar o desactivar la lógica automática asociada a reglas
- Comportamiento al desactivar:
  - no se borran las reglas ya cargadas
  - simplemente dejan de verse en UI y no se ejecuta la lógica automática de creación/regeneración
  - las estadísticas de reglas quedan ocultas

### 2026-04-23

- Se agregan assets de identidad para navegador e instalación como app:
  - `favicon.ico`
  - `favicon.svg`
  - `apple-touch-icon.png`
  - `app-icon-192.png`
  - `app-icon-512.png`
  - `site.webmanifest`
- El ícono usa una composición alegórica a `Plan de Visitas`:
  - fondo rojo institucional
  - ruta de visitas con nodos
  - calendario/check de cumplimiento
- Se conectan los íconos en el `head` del HTML con soporte para favicon, Apple touch icon y manifest web app.
