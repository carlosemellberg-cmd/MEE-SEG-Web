# MEE SEG Web v1.10.1 — GitHub Pages

Versión web pública del Comité de Seguridad MEE. El repositorio contiene únicamente el programa; no incluye el maestro, personas, tareas, submissions, contraseñas ni tokens.

## Cambios v1.10.1

- Apuntes rápidos locales con guardado automático.
- Opción **Traer desde apunte** dentro de Nueva tarea.
- Administrador, supervisor, coordinador, líder y jefe pueden crear, modificar y eliminar.
- Operadores, integrantes y analistas no administradores quedan en modo solo lectura.
- Botón **Actualizar estado** para seleccionar el maestro actualizado y confirmar el ACK sin salir del centro de archivos.
- Correo de Power Automate configurado de forma predeterminada.
- Corrección para que Chrome en una tablet Android sea tratado como navegador web y no como APK.
- Caché PWA renovada a `mee-seg-web-v1.10.1`.

## Flujo actual

1. Cargar `MEE_DATOS_COMITE_MASTER.json`.
2. Elegir el usuario.
3. Registrar tareas o apuntes según permisos.
4. Preparar el envío seguro.
5. Descargar/compartir `MEE_SUBMISSION_<versión>_<UUID>.json`.
6. Verificar el adjunto en Outlook y enviar.
7. Descargar el maestro actualizado desde SharePoint.
8. Usar **Actualizar estado** y seleccionar el maestro para confirmar el ACK.

## Seguridad

- Los datos se almacenan en `localStorage` del navegador.
- Los apuntes quedan solamente en el navegador hasta convertirse en una tarea.
- GitHub Pages no recibe ni almacena el archivo maestro.
- La web no escribe SharePoint directamente.
- El envío automático directo queda pendiente del segundo flujo de Power Automate.
