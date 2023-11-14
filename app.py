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
from abc import ABC, abstractmethod
import random
# Cargar las variables de entorno desde .env
load_dotenv()

# Estrategia Base
class SorteoEstrategia(ABC):
    @abstractmethod
    def asignar_amigos(self, participantes, restricciones):
        pass

# Estrategia Concreta A
class SorteoSimpleEstrategia(SorteoEstrategia):
    def asignar_amigos(self, participantes, restricciones):
        amigos = list(participantes.keys())
        random.shuffle(amigos)
        return dict(zip(amigos, amigos))

class SorteoComplejoEstrategia(SorteoEstrategia):
    def asignar_amigos(self, participantes, restricciones):
        # Crear una copia de los participantes para trabajar con ella
        amigos = list(participantes.keys())
        random.shuffle(amigos)

        asignaciones = {}
        intentos = 0
        max_intentos = 1000

        while len(asignaciones) < len(amigos):
            for donante in amigos:
                posibles_receptores = [r for r in amigos if r != donante and r not in asignaciones.values() and r not in restricciones.get(donante, [])]
                
                if not posibles_receptores:
                    # Si no hay posibles receptores, reiniciar el proceso
                    asignaciones.clear()
                    random.shuffle(amigos)
                    intentos += 1
                    break
                
                receptor = random.choice(posibles_receptores)
                asignaciones[donante] = receptor

            if intentos > max_intentos:
                raise Exception("No se puede generar una asignación válida con las restricciones dadas.")

        return asignaciones


class AmigoSecreto:
    def __init__(self, estrategia: SorteoEstrategia, participantes, restricciones=None):
        self.estrategia = estrategia
        self.participantes = participantes
        self.restricciones = restricciones or {}

    def asignar_amigos(self):
        return self.estrategia.asignar_amigos(self.participantes, self.restricciones)


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
        mensaje['From'] = "Amigo Secreto 2023"
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

    def __init__(self, participantes, restricciones=None):
        # Inicializar con un diccionario de participantes y sus correos electrónicos.
        # Añadir un diccionario de restricciones donde la clave es quien no puede regalar a quien.
        self.participantes = participantes
        self.restricciones = restricciones or {}

    def verificar_restricciones(self, asignaciones):
        # Verificar si las asignaciones actuales violan alguna restricción.
        for donante, receptor in asignaciones.items():
            if receptor in self.restricciones.get(donante, []):
                return False
        return True

    def detectar_ciclos(self):
        # Función para detectar ciclos en las restricciones.
        visitado = set()
        
        def visitar(nodo, camino):
            if nodo in camino:
                return True  # Ciclo detectado.
            camino.add(nodo)
            for vecino in self.restricciones.get(nodo, []):
                if visitar(vecino, camino):
                    return True
            camino.remove(nodo)
            return False
        
        for nodo in self.restricciones:
            if visitar(nodo, set()):
                return True  # Hay un ciclo.
        return False  # No hay ciclos.

    def asignar_amigos(self):
        # Asignar amigos secretos de manera aleatoria.
        amigos = list(self.participantes.keys())
        random.shuffle(amigos)
        asignaciones = dict(zip(amigos, amigos))

        # Asegurarse de que las asignaciones cumplan con las restricciones.
        while not self.verificar_restricciones(asignaciones):
            random.shuffle(amigos)
            asignaciones = dict(zip(amigos, amigos))

        # Verificar que nadie se auto-asigne y que se respeten las restricciones.
        intentos = 0
        max_intentos = 1000
        while any(nombre == amigo for nombre, amigo in asignaciones.items()) or not self.verificar_restricciones(asignaciones):
            if intentos > max_intentos:
                raise Exception("No se puede generar una asignación válida con las restricciones dadas.")
            random.shuffle(amigos)
            asignaciones = dict(zip(amigos, amigos))
            intentos += 1

        # Devolver el diccionario con las asignaciones.
        return asignaciones


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

        # Agregar interfaz para restricciones
        self.restricciones = {}  # Diccionario para almacenar restricciones
        self.restricciones_var = StringVar(master)

        # Etiqueta y campo de entrada para las restricciones
        self.etiqueta_restricciones = Label(self.frame_formulario, text="Restricciones (formato: Luis no Juan)", font=fuente_texto, bg=color_fondo, fg=color_texto)
        self.etiqueta_restricciones.grid(row=3, column=0, padx=10, sticky=W)
        self.entrada_restricciones = Entry(self.frame_formulario, textvariable=self.restricciones_var, font=fuente_texto, width=30)
        self.entrada_restricciones.grid(row=3, column=1, padx=10, pady=10)

        # Botón para agregar restricciones
        self.boton_agregar_restriccion = Button(self.frame_formulario, text="Agregar Restricción", font=fuente_texto, bg=color_botones, fg=color_acento, command=self.agregar_restriccion)
        self.boton_agregar_restriccion.grid(row=4, column=0, columnspan=2, pady=10)

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

    def agregar_restriccion(self):
        # Obtener la restricción del campo de entrada y actualizar el diccionario de restricciones
        restriccion = self.restricciones_var.get()
        if restriccion:
            donante, _, receptor = restriccion.partition(' no ')
            if donante and receptor:
                if donante in self.restricciones:
                    self.restricciones[donante].append(receptor)
                else:
                    self.restricciones[donante] = [receptor]
                self.texto_participantes.insert(END, f"Restricción agregada: {donante} no puede regalar a {receptor}\n")
                self.restricciones_var.set('')
            else:
                messagebox.showwarning("Advertencia", "Formato de restricción incorrecto")
        else:
            messagebox.showwarning("Advertencia", "La restricción no puede estar vacía")

    
    def realizar_sorteo_enviar_correos(self):
        # Seleccionar la estrategia basada en si hay restricciones definidas o no.
        estrategia = SorteoComplejoEstrategia() if self.restricciones else SorteoSimpleEstrategia()
        juego_amigo_secreto = AmigoSecreto(estrategia, self.participantes, self.restricciones)
        
        try:
            parejas = juego_amigo_secreto.asignar_amigos()
            gestor_correo = GestorCorreo()
            ruta_imagen = 'Sorteo_Amigo_Secreto/Img/regalo.jpg'
            mensajes_exitosos = []
            for persona, amigo in parejas.items():
                asunto = "Tu amigo secreto 2023"
                cuerpo = f"""
                            <html>
                                <body>
                                    <p>Hola, espero que estés de maravilla {persona},</p>
                                    <p>Tu amigo secreto es: {amigo}</p>
                                    <p><img src="cid:{os.path.basename(ruta_imagen)}"></p>
                                    <p>¡Saludos!</p>
                                </body>
                            </html>
                            """
            try:
                gestor_correo.enviar_correo(self.participantes[persona], asunto, cuerpo, ruta_imagen)
                mensajes_exitosos.append(f"Correo enviado con éxito a {persona}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo enviar el correo a {persona}: {e}")
            gestor_correo.cerrar_conexion()
            if mensajes_exitosos:
                messagebox.showinfo("Correos Enviados", "\n".join(mensajes_exitosos) + "\nSorteo realizado y correos enviados.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo realizar el sorteo: {e}")

if __name__ == '__main__':
    root = Tk()
    app = Aplicacion(root)
    # Aquí debes definir las restricciones, por ejemplo:
    restricciones = {
        'Luis': ['Juan'],  # Luis no puede regalar a Juan
        # Puedes agregar más restricciones aquí
    }
    juego_amigo_secreto = AmigoSecreto(app.participantes, restricciones)
    root.mainloop()

