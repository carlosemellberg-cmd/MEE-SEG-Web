# MEE SEG Web — GitHub Pages

Versión pública del programa, sin datos corporativos incluidos.

## Funcionamiento

1. Abrir el enlace.
2. Cargar `MEE_DATOS_COMITE_MASTER.json`.
3. Elegir el usuario local.
4. Realizar cambios.
5. Preparar el envío seguro.
6. Descargar o compartir `MEE_SUBMISSION_<versión>_<UUID>.json`.
7. Adjuntarlo al correo de Power Automate con asunto `MEE_SEG_SYNC`.
8. Descargar el maestro actualizado desde SharePoint.
9. Confirmar el ACK seleccionando ese maestro.

## Seguridad

- GitHub Pages aloja únicamente HTML, CSS, JavaScript e imágenes.
- El repositorio no contiene personas, tareas, maestro, submissions, contraseñas ni tokens.
- Los datos quedan en `localStorage` del navegador usado.
- La web no sobrescribe SharePoint ni el archivo maestro.
- La sincronización sigue pasando por Power Automate.

## Limitación actual

El navegador no puede adjuntar silenciosamente un archivo en Outlook. La web descarga/compartirá el submission y abrirá el cliente de correo; el usuario debe verificar que el JSON esté adjunto antes de enviar.
