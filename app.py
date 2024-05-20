import speech_recognition as sr
import pyttsx3, pywhatkit, subprocess, pyjokes, wikipedia, datetime, webbrowser, os, pathlib, pyautogui
from time import sleep
from unidecode import unidecode

# Programas
photoshop = pathlib.Path("C:/Program Files/Adobe/Adobe Photoshop 2023", "Photoshop.exe")
registro_word = str(pathlib.Path("C:/Users/jesus/Dropbox/Desarollo","perfil.docx"))
registro_amubaweb = str(pathlib.Path("C:/Users/jesus/Dropbox/Desarollo","ambubaweb.docx"))
registro_proyectos = str(pathlib.Path("C:/Users/jesus/Dropbox/Desarollo","proyectos.xlsx"))

# Escuchar nuestro microfono y transformarlo a texto
def audio_to_text():

    # almacenar recognizer
    r = sr.Recognizer()

    # Configurar microfono
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        # Asignar un tiemp de espera
        r.pause_threshold = .5

        # Mensaje de comienzo
        print("Esperando instruccion...")
    
        # almacenar el audio
        audio = r.listen(source)
    try:
        # Transcribir el audio a texto
        valor = r.recognize_google(audio, language= "es-es")
        print("Creo que dijiste: " + valor)
        return valor
    except sr.UnknownValueError:
        print("No pude enternder lo que dijiste")
    except sr.RequestError as e:
        print("Error de tipo request; {0}".format(e))

def speak(mensaje):

    engine = pyttsx3.init()
    engine.setProperty("voice", "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_ES-ES_HELENA_11.0")

    engine.say(mensaje)
    engine.runAndWait()

#Helpers
def limpiar_entrada(entrada,texto):
    entrada = entrada.replace(texto,"")
    return entrada

#Funcionalidades
def pedir_dia():

    dia = datetime.date.today()

    dia_semana = dia.weekday()

    calendario = { 0:"Lunes",
                1: "Martes",
                2: "Miércoles",
                3: "Jueves",
                4: "Viernes",
                5: "Sábado",
                6: "Domingo"}

    speak(f"Hoy es {calendario[dia_semana]}")

def copilot(pedido): # patro--- preguntale a copilot | busca en copilot | pregunta a copilot ---
    pedido = unidecode(pedido)
    if "preguntale a copilot" in pedido or "pregunta a copilot" in pedido or "busca en copilot" in pedido:
        pedido = limpiar_entrada(pedido,"preguntale a copilot")
        pedido = limpiar_entrada(pedido,"pregunta a copilot")
        pedido = limpiar_entrada(pedido,"busca en copilot")
        pyautogui.hotkey("win","c")
        sleep(1)
        pyautogui.write(pedido)
        pyautogui.press("enter")

def pedir_hora():
        
        dia = datetime.datetime.now()

        hora = f"Son las {dia.hour} y {dia.minute} minutos"

        speak(hora)

def msg_inicio():
    speak("Iniciando, por favor espere")

def chiste():

    chiste = pyjokes.get_joke(language="es")

    speak(F"bien, el chiste es: {chiste}")

def about_me():

    about = " Soy la primera version de lu. una inteligencia artificial creada para optimizar tiempo al usuario"

    speak(about)

def procesos_abrir(pedido):
    if "abre youtube" in pedido or "youtube" in pedido or "you tube" in pedido:
        webbrowser.open("https://www.youtube.com")
    elif "abre udemy" in pedido or "udemy" in pedido:
        webbrowser.open("https://www.udemy.com")
    elif "abre chatgpt" in pedido or "abre chat gpt" in pedido or "chatgpt" in pedido or "chat gpt" in pedido:
        webbrowser.open("https://chatgpt.com/c/0db011f1-16fd-4ba6-8675-e8d5fe720788")
    elif "abre jkanime" in pedido or "abre jk anime" in pedido or "jk anime" in pedido or "jkanime" in pedido:
        webbrowser.open( "https://jkanime.net/")
    elif "abre photoshop" in pedido or "abrir photoshop" in pedido:
        subprocess.Popen([str(photoshop)])
    elif "abre la calculadora" in pedido or "abre calculador" in pedido or "abrir calculadora" in pedido or "abrir la calculadora" in pedido:
        subprocess.run("calc.exe")
    elif "muéstrame los repositorios" in pedido or "abre el repositorio" in pedido or "abre repositorios" in pedido or "abre repositorio" in pedido:
        webbrowser.open("https://github.com/trending")
    elif "jugar ajedrez" in pedido or "abrir ajedrez" in pedido:
        webbrowser.open("https://www.chess.com/play/computer/ChessDiver_Scanner")

def redactar(pedido):
    if "redacta y busca" in pedido:
        pedido = pedido.replace("redacta y busca", "")
        pedido = unidecode(pedido)
        pyautogui.write(pedido)
        pyautogui.press("enter")
    elif "redacta" in pedido:
        pedido = pedido.replace("redacta", "")
        pedido = unidecode(pedido)
        pyautogui.write(pedido)

def funcionalidades(pedido):
    pedido = unidecode(pedido)
    if "f5" in pedido:
        pyautogui.press("f5")
    elif "abre el libro de registros" in pedido or "abrir el libro de registros" in pedido or "abrir libro de registros" in pedido:
        print(registro_word)
        subprocess.run(['start', '', registro_word], shell=True)
    elif "registro catorce" in pedido or "abrir el libro de registros catorce" in pedido or "abrir libro de registros catorce" in pedido or "registro catorce" in pedido or "registro 14" in pedido:
        print(registro_word)
        subprocess.run(['start', '', registro_amubaweb], shell=True)
    elif "abre el libro de proyectos" in pedido or "abrir el libro de proyectos" in pedido or "abrir libro de proyectos" in pedido:
        print(registro_word)
        subprocess.run(['start', '', registro_proyectos], shell=True)
    elif "escribe la fecha de hoy" in pedido or "escribe fecha de hoy" in pedido or "coloca la fecha de hoy" in pedido:
        fecha = datetime.date.today().strftime("%d/%m/%y")
        pyautogui.write(fecha)
    elif "cerrar programa" in pedido:
        print("programa cerrado")
        pyautogui.hotkey("alt","f4")
    elif "parar musica" in pedido or "reanudar musica" in pedido:
        print("musica pausada/reanudada")
        pyautogui.hotkey("fn","f7")
    elif "limpia la consola" in pedido or "limpiar consola" in pedido:
        os.system("cls")
    elif "busca en wikipedia" in pedido or "buscar en wikipedia" in pedido:
        speak("Buscando...")
        pedido = pedido.replace("busca en wikipedia", '')
        pedido = pedido.replace("buscar en wikipedia", '')
        wikipedia.set_lang("es")
        resultado = wikipedia.summary(pedido, sentences=1)
        speak("Consegui esto...")
        speak(resultado)
    elif "busca en internet" in pedido or "buscar en internet" in pedido:
        pedido = pedido.replace("busca en internet", "")
        pedido = pedido.replace("buscar en internet", "")
        pywhatkit.search(pedido)
        speak("Busqueda lista")
    elif "reproducir" in pedido or "reproduce" in pedido:
        pedido = pedido.replace("reproducir", "")
        pedido = pedido.replace("reproduce", "")
        pywhatkit.playonyt(pedido)
    elif "aumenta la velocidad a 4" in pedido or "aumentar velocidad" in pedido or "aumenta la velocidad del video" in pedido or "aumenta la velocidad" in pedido or "aumentar la velocidad" in pedido:
        pyautogui.press("d")
        pyautogui.press("d")
        pyautogui.press("d")
        pyautogui.press("d")
    elif "reduce la velocidad a 4" in pedido or "reducir velocidad" in pedido or "reduce la velocidad del video" in pedido or "reduce la velocidad" in pedido or "reducir la velocidad" in pedido:
        pyautogui.press("s")
        pyautogui.press("s")
        pyautogui.press("s")
        pyautogui.press("s")

def engine():
    """Nucleo del asistente, encargado de unir todas las funcionalidades"""
    msg_inicio()

    comenzar = True

    while comenzar:

        try:
            
            pedido = audio_to_text().lower()
            # Funcionalidades deacuerdo al texto del pedido
            procesos_abrir(pedido)
            redactar(pedido)
            funcionalidades(pedido)
            copilot(pedido)
            if "dime un chiste" in pedido:
                chiste()
            elif "dame la hora" in pedido or "qué hora es" in pedido:
                pedir_hora()
            elif "qué día es hoy" in pedido or "qué dia es" in pedido:
                pedir_dia()
            elif "quién eres" in pedido:
                about_me()

            elif "cerrar bucle" in pedido or "eso es todo" in pedido or "cierra bucle" in pedido:
                speak("Terminando conversacion")
                os._exit(1)

        except KeyboardInterrupt:
            os._exit(1)
        except:
            print("ERROR TEMPORAL, REINTENTANDO...")

if __name__ == "__main__":
    engine()