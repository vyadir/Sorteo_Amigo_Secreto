"""
    Elaborado por: Yadir Vega Espinoza
    Objetivo: Realizar sorteo de amigo secreto y enviar por correo electrónico a cada participante su respectivo amigo secreto.
    Fecha:2023-11-07
"""

import smtplib
import random
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from dotenv import load_dotenv

# Cargar las variables de entorno desde un archivo .env para mantener las credenciales seguras.
load_dotenv()

class GestorCorreo:
    """Clase para gestionar la conexión SMTP y enviar correos electrónicos."""
    def __init__(self):
        # Inicializar la clase con las credenciales y configuración del servidor SMTP.
        self.servidor_smtp = os.getenv('SMTP_SERVIDOR')
        self.puerto = int(os.getenv('SMTP_PUERTO'))
        self.usuario = os.getenv('SMTP_USUARIO')
        self.contraseña = os.getenv('SMTP_CONTRASEÑA')
        # Establecer la conexión inicial con el servidor SMTP.
        self.servidor = smtplib.SMTP(self.servidor_smtp, self.puerto)
        self.conectar()

    def conectar(self):
        # Intentar establecer una conexión segura con el servidor SMTP.
        try:
            self.servidor.connect(self.servidor_smtp, self.puerto)
            self.servidor.ehlo()  # Verificar la conexión con el servidor.
            self.servidor.starttls()  # Iniciar el modo TLS para una conexión segura.
            self.servidor.ehlo()  # Verificar la conexión nuevamente.
            self.servidor.login(self.usuario, self.contraseña)  # Iniciar sesión en el servidor SMTP.
        except smtplib.SMTPException as e:
            print(f"Error al conectar al servidor SMTP: {e}")
            raise e

    def enviar_correo(self, destinatario, asunto, cuerpo, ruta_imagen):
        # Crear un mensaje de correo electrónico con un asunto, remitente y destinatario.
        mensaje = MIMEMultipart()
        mensaje['Subject'] = asunto
        mensaje['From'] = self.usuario
        mensaje['To'] = destinatario

        # Adjuntar el cuerpo del correo como HTML para poder incluir imágenes y formateo.
        mensaje.attach(MIMEText(cuerpo, 'html'))

        # Intentar abrir y adjuntar la imagen al mensaje.
        try:
            with open(ruta_imagen, 'rb') as archivo_imagen:
                img = MIMEImage(archivo_imagen.read())
                # Agregar un encabezado para identificar la imagen en el cuerpo HTML.
                img.add_header('Content-ID', '<{}>'.format(os.path.basename(ruta_imagen)))
                mensaje.attach(img)
        except IOError as e:
            print(f"No se pudo abrir la imagen: {e}")
            raise

        # Intentar enviar el mensaje a través del servidor SMTP y confirmar el envío.
        try:
            self.servidor.send_message(mensaje)
            print(f"Correo enviado con éxito a {destinatario}")
        except smtplib.SMTPException as e:
            print(f"Error al enviar correo a {destinatario}: {e}")

    def cerrar_conexion(self):
        # Cerrar la conexión con el servidor SMTP de forma segura.
        try:
            self.servidor.quit()
        except smtplib.SMTPException as e:
            print(f"Error al cerrar la conexión SMTP: {e}")

class AmigoSecreto:
    """Clase para manejar la lógica del sorteo del amigo secreto."""
    def __init__(self, participantes):
        # Inicializar con un diccionario de participantes y sus correos electrónicos.
        self.participantes = participantes

    def asignar_amigos(self):
        # Asignar amigos secretos de manera aleatoria.
        amigos = list(self.participantes.keys())
        random.shuffle(amigos)
        # Asegurarse de que nadie se asigne a sí mismo.
        while any(nombre == amigo for nombre, amigo in zip(self.participantes, amigos)):
            random.shuffle(amigos)
        # Devolver un diccionario con las asignaciones.
        return dict(zip(self.participantes, amigos))

# Ejemplo de uso
if __name__ == '__main__':
    # Diccionario de participantes con sus correos electrónicos.
    participantes = {
        'Alice': 'y.vega.x@gmail.com',
        'Bob': 'yadir.sve@outlook.com',
        'Charlie': 'yadir.vega@estudiantec.cr',
    }
    # Crear instancias de las clases y realizar el sorteo.
    gestor_correo = GestorCorreo()
    juego_amigo_secreto = AmigoSecreto(participantes)
    parejas = juego_amigo_secreto.asignar_amigos()
    # Ruta a la imagen que se incluirá en el correo.
    ruta_imagen = 'Sorteo_Amigo_Secreto/regalo.jpg'
    # Enviar un correo a cada participante con su amigo secreto asignado.
    for persona, amigo in parejas.items():
        asunto = "Tu amigo secreto!"
        # Crear el cuerpo del correo como HTML, incluyendo una referencia a la imagen adjunta.
        cuerpo = f"""
        <html>
            <body>
                <p>Hola {persona},</p>
                <p>Tu amigo secreto para el regalo es: {amigo}</p>
                <p><img src="cid:{os.path.basename(ruta_imagen)}"></p>
            </body>
        </html>
        """
        # Llamar al método para enviar el correo.
        gestor_correo.enviar_correo(participantes[persona], asunto, cuerpo, ruta_imagen)
    # Cerrar la conexión con el servidor SMTP.
    gestor_correo.cerrar_conexion()
