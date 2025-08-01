Visor MSG
  Un visor de archivos .msg (correos de Outlook) desarrollado en Python con PyQt5. Permite visualizar el cuerpo del correo (HTML o texto plano), imágenes incrustadas, y adjuntos, con la capacidad de descargarlos.
Características

Muestra el asunto, remitente y destinatarios.
Renderiza el cuerpo del correo en HTML o texto plano.
Muestra vistas previas de adjuntos de imagen.
Permite descargar adjuntos.
Interfaz inspirada en Outlook con una sección de adjuntos compacta.
Fuente de botones ajustada a 14px para mayor legibilidad.

Instalación

Descarga el ejecutable desde Releases.
(Opcional) Instala las dependencias para ejecutar el código fuente:pip install extract-msg PyQt5


Ejecuta el programa:python visor_msg.py



Uso

Haz doble clic en visor_msg.exe o ejecuta python visor_msg.py.
Selecciona un archivo .msg para visualizarlo.
Usa los botones "Descargar: {nombre_adjunto}" para guardar adjuntos.

Notas

El programa genera archivos temporales en %TEMP% para imágenes incrustadas y vistas previas, que se eliminan al cerrar la ventana.
La asociación de archivos .msg con el ejecutable se puede configurar en Windows.
