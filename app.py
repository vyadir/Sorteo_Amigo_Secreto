"""
    Elaborado por: Yadir Vega Espinoza
    Objetivo: Realizar sorteo de amigo secreto y enviar por correo electrónico a cada participante su respectivo amigo secreto.
    Fecha:2023-11-07
"""

import smtplib
import random
import os
from tkinter import *
from tkinter import messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from dotenv import load_dotenv

# Cargar las variables de entorno desde .env
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

# Interfaz gráfica de usuario
class Aplicacion:
    def __init__(self, master):
        # Configuración de la ventana principal
        self.master = master
        master.title('Sorteo Amigo Secreto')
        master.attributes('-fullscreen', False)
        master.state('zoomed')  
        # Inicialización de las variables de control para las entradas de texto
        self.nombre_var = StringVar(master)
        self.correo_var = StringVar(master)
        # Inicializar el diccionario de participantes
        self.participantes = {}
    
        # Colores del esquema análogo
        color_fondo_principal = "#34568B"  # Azul oscuro
        color_fondo = "#9ABAD9"  # Azul claro
        color_acento = "#5975A4"  # Azul intermedio
        color_texto = "#FFFFFF"  # Blanco para suficiente contraste
        color_botones = "#D9E2EC"  # Azul muy claro

        # Configurar colores y estilos
        fuente_texto = ("Helvetica", 14)
        fuente_titulo = ("Helvetica", 16, "bold")
        master.configure(bg=color_fondo_principal)

        # Título de la aplicación
        self.titulo = Label(master, text="Sorteo Amigo Secreto", font=fuente_titulo, bg=color_fondo_principal, fg=color_texto)
        self.titulo.pack(pady=20)

        # Frame para el formulario de entrada
        self.frame_formulario = Frame(master, bg=color_fondo)
        self.frame_formulario.pack(pady=10)

        # Etiqueta y campo de entrada para el nombre
        self.etiqueta_nombre = Label(self.frame_formulario, text="Nombre", font=fuente_texto, bg=color_fondo, fg=color_texto)
        self.etiqueta_nombre.grid(row=0, column=0, padx=10, sticky=W)
        self.entrada_nombre = Entry(self.frame_formulario,textvariable=self.nombre_var, font=fuente_texto, width=30)
        self.entrada_nombre.grid(row=0, column=1, padx=10, pady=10)

        # Etiqueta y campo de entrada para el correo
        self.etiqueta_correo = Label(self.frame_formulario, text="Correo", font=fuente_texto, bg=color_fondo, fg=color_texto)
        self.etiqueta_correo.grid(row=1, column=0, padx=10, sticky=W)
        self.entrada_correo = Entry(self.frame_formulario,textvariable=self.correo_var, font=fuente_texto, width=30)
        self.entrada_correo.grid(row=1, column=1, padx=10, pady=10)

        # Botón para agregar participante
        self.boton_agregar = Button(self.frame_formulario, text="Agregar Participante", font=fuente_texto, bg=color_botones, fg=color_acento, command=self.agregar_participante)
        self.boton_agregar.grid(row=2, column=0, columnspan=2, pady=10)

        # Frame para la lista de participantes y botones de acción
        self.frame_lista = Frame(master, bg=color_fondo)
        self.frame_lista.pack(fill=BOTH, expand=True, pady=10)

        # Área de texto para mostrar los participantes
        self.texto_participantes = Text(self.frame_lista, font=fuente_texto, height=10, width=50)
        self.texto_participantes.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=10)

        # Frame para los botones de acción
        self.frame_botones = Frame(self.frame_lista, bg=color_fondo)
        self.frame_botones.pack(fill=Y, expand=False, pady=10)

        # Botón para realizar sorteo y enviar correos
        self.boton_enviar = Button(self.frame_botones, text="Realizar Sorteo y Enviar Correos", font=fuente_texto, bg=color_botones, fg=color_acento, command=self.realizar_sorteo_enviar_correos)
        self.boton_enviar.pack(fill=X, padx=10, pady=5)

        # Botón para cerrar la aplicación
        self.boton_salir = Button(self.frame_botones, text="Cerrar", font=fuente_texto, bg=color_botones, fg=color_acento, command=master.quit)
        self.boton_salir.pack(fill=X, padx=10, pady=5)

    def agregar_participante(self):
        nombre = self.nombre_var.get()
        correo = self.correo_var.get()
        if nombre and correo:
            self.participantes[nombre] = correo
            self.texto_participantes.insert(END, f"{nombre}: {correo}\n")
            self.nombre_var.set('')
            self.correo_var.set('')
        else:
            print("Nombre o correo no puede estar vacío")

    def realizar_sorteo_enviar_correos(self):
        juego_amigo_secreto = AmigoSecreto(self.participantes)
        parejas = juego_amigo_secreto.asignar_amigos()
        gestor_correo = GestorCorreo()
        ruta_imagen = 'Sorteo_Amigo_Secreto/Img/regalo.jpg'
        mensajes_exitosos = []  # Lista para almacenar mensajes de confirmación
        for persona, amigo in parejas.items():
            asunto = "Tu amigo secreto 2023, Familia Vega Espinoza"
            cuerpo = f"""
                        <html>
                            <body>
                                <p>Hola, espero que estés de maravilla {persona},</p>
                                <p>Tu amigo secreto para el regalo es: {amigo}</p>
                                <p><img src="cid:{os.path.basename(ruta_imagen)}"></p>
                            </body>
                        </html>
                        """
            try:
                gestor_correo.enviar_correo(self.participantes[persona], asunto, cuerpo, ruta_imagen)
                mensajes_exitosos.append(f"Correo enviado con éxito a {persona}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo enviar el correo a {persona}: {e}")
                
        gestor_correo.cerrar_conexion()
        # Mostrar un mensaje pop-up con los resultados del envío
        if mensajes_exitosos:  # Si hay mensajes exitosos para mostrar
            messagebox.showinfo("Correos Enviados", "\n".join(mensajes_exitosos) + "\nSorteo realizado y correos enviados.")



# Ejemplo de uso
if __name__ == '__main__':
    root = Tk()
    app = Aplicacion(root)
    root.mainloop()

