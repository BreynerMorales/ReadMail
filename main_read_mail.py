import imaplib
import email
from email.header import decode_header
import time
import re
import os

# ===============================
# LIBRARY FROM SOUND READ
import asyncio
import edge_tts
import io
import pygame
# ===============================
import winsound
import imaplib, ssl, socket
import numpy as np

pygame.mixer.init(frequency=44100, size=-16, channels=2)  # estéreo

def beep(frequency, duration_ms):
    sample_rate = 44100
    n_samples = int(sample_rate * duration_ms / 1000)
    t = np.linspace(0, duration_ms/1000, n_samples, False)

    # Genera onda senoidal
    wave = (np.sin(2 * np.pi * frequency * t) * 32767).astype(np.int16)

    # Convierte a 2D (estéreo: L y R iguales)
    wave_stereo = np.column_stack((wave, wave))

    # Crea el sonido
    sound = pygame.sndarray.make_sound(wave_stereo)
    sound.play()

    # Espera a que termine
    pygame.time.delay(duration_ms)

# Configuración de la cuenta
IMAP_SERVER = "imap.gmail.com"   # servidor IMAP (para Gmail)
EMAIL_ACCOUNT = "user_test@gmail.com"
PASSWORD = "your application pass"
def conectar():
    """Conecta al servidor IMAP"""
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    #mail.select("inbox")
    return mail

def buscar_correos(mail,type_mail):
    """Busca correos cuyo asunto contenga R-123 o I-123, pero sin marcarlos como leídos"""
    # try:
    
        
    #mail.select("inbox")  # bandeja de entrada
    mail.select(type_mail)  # bandeja de entrada
    status, mensajes = mail.search(None, "UNSEEN")  # solo correos no leídos
    if status != "OK":
        return []

    correos = []
    for num in mensajes[0].split():
        # Usamos BODY.PEEK[] para que no cambie el estado de no leído
        status, data = mail.fetch(num, "(BODY.PEEK[])")
        if status != "OK":
            continue

        msg = email.message_from_bytes(data[0][1])

        # Obtener asunto
        subject, encoding = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8", errors="ignore")

        # Verificar si contiene R-123 o I-123
        PATRON_ASUNTO = re.search(r'([RrIi]-\d+)', subject)
        if PATRON_ASUNTO:
            codigo = PATRON_ASUNTO.group(1)  

            # Obtener cuerpo
            cuerpo = ""
            if msg.is_multipart():
                for part in msg.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get("Content-Disposition"))

                    if ctype == "text/plain" and "attachment" not in cdispo:
                        cuerpo = part.get_payload(decode=True).decode(errors="ignore")
                        break
            else:
                cuerpo = msg.get_payload(decode=True).decode(errors="ignore")

            correos.append((subject.strip(), cuerpo.strip(), codigo))

    return correos
pygame.mixer.init()
async def hablar(texto, voz="es-PE-AlvaroNeural", rate="-5%", pitch="-50Hz"):
    communicate = edge_tts.Communicate(
        texto, 
        voice=voz, 
        rate=rate, 
        pitch=pitch
    )

    stream = io.BytesIO()

    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            stream.write(chunk["data"])

    # Ir al inicio del buffer
    stream.seek(0)
    # Reproducir con pygame
    pygame.mixer.music.load(stream, "mp3")
    pygame.mixer.music.play()

    while pygame.mixer.music.get_busy():
        pygame.time.Clock().tick(1)
def read_sound(text):
    asyncio.run(hablar(
        text,
        voz="es-ES-AlvaroNeural",
        rate="+5%",
        pitch="-50Hz"
    ))

def main(Type_bandeja):
    mail = conectar()
    try:
        while True:
            for type_mail, description_type in Type_bandeja.items():
                correos = buscar_correos(mail,type_mail)
                if len(correos)>1:
                    text_print_mails = f"Se recibieron {len(correos)} nuevos correos en la {description_type}"
                    print(text_print_mails)
                    read_sound(text_print_mails)
                elif len(correos)==0:
                    print(f"".center(50, "-"))
                    print(f" > {description_type.upper()} < ".center(50, "="))
                    text_print_mails = f'No se encontraron nuevos mensajes en la {description_type}'
                    print(f"".center(50, "-"))
                    print(text_print_mails)
                    
                else:
                    text_print_mails = f"Se recibio un nuevo correo"
                    print(text_print_mails)
                    read_sound(text_print_mails)
                pivot_count = 1
                for asunto, cuerpo, codigo in correos:
                    print(f" Correo N°{pivot_count} ".center(50, "="))
                    # print("Nuevo correo:")
                    print("CODIGO ENCONTRADO:", codigo)
                    print("Asunto:", asunto)
                    print("Cuerpo:", cuerpo)
                    print(f"".center(50, "="))
                    print("")
                    pivot_count+=1
                
                if len(correos) != 0:
                    read_sound(text_print_mails)
                    for _ in range(2):  # número de ciclos de subida/bajada
                        # Subida
                        for f in range(1000, 2000, 100):
                            # winsound.Beep(f, 100)  # frecuencia, duración (ms)
                            beep(f, 100)
                        # Bajada
                        for f in range(2000, 1000, -100):
                            # winsound.Beep(f, 100)
                            beep(f, 100)
                
                for second in range(10,0,-1):
                
                    var_null = " "*(10-second)
                    print(f"{var_null}".ljust(10, "*"),": [",f"{second}".rjust(2, " "),"]", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                    if len(correos) != 0:
                        winsound.Beep(1000, 150)
                    else:
                        time.sleep(0.3)  # espera 0.1s antes de volver a buscar
            text_print_mails = "\n=== > > > Volviendo a reiniciar la busqueda"
            print(text_print_mails)
            time.sleep(3)
            os.system("cls" if os.name == "nt" else "clear")
    except (imaplib.IMAP4.abort, ssl.SSLEOFError, OSError) as e:
        print(f"⚠️ Conexión perdida: {e}, reiniciando main()...")
        read_sound("ADVERTENCIA! Conexión IMAP perdida. Se volvera a reiniciar en unos segundos")
        # time.sleep(3)
        try:
            mail.logout()
        except:
            pass
        beep(1500, 500)
    except KeyboardInterrupt:
        print("⛔ Script detenido por el usuario.")
        raise 
    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        read_sound("Error inesperado")
        for _ in range(2):  # número de ciclos de subida/bajada
            # Subida
            for f in range(1000, 2000, 100):
                # winsound.Beep(f, 100)  # frecuencia, duración (ms)
                beep(f, 100)
            # Bajada
            for f in range(2000, 1000, -100):
                # winsound.Beep(f, 100)
                beep(f, 100)
        
if __name__ == "__main__":
    Type_bandeja = {"INBOX"        :"Bandeja Principal",
                    "[Gmail]/Spam" :"Bandeja Spam"}
    winsound.PlaySound("C:\\Windows\\Media\\tada.wav", winsound.SND_FILENAME)
    while True:
        try:
            read_sound("Bienvenido, El sistema de rastreo de correos en tiempo real ha sido activado")
            main(Type_bandeja)
        except (socket.error) as e:
            print(f"⚠️ Conexión perdida: {e}, reiniciando main()...")
            beep(1000, 400)
            beep(1000, 400)
            print("Ha ocurrido un error")
        except Exception as e:
            print("Error inseperado",e)
            break
