# VBSInjection

`[INFO]` Imagen:

![VBSInjection](https://github.com/ZombieGeeK0/VBSInjection/assets/158185295/c91c9fff-d100-40fc-86c5-d156febc299d)

<hr>

`[INFO]` Este es un proyecto para `Windows,` pero también se pueden generar los `scripts en Linux.`

`[INFO]` Este proyecto combina `Python, BATCH y VBS (Visual Basic Scripting).`

`[INFO]` Especiales agradecimientos a <a href="https://www.github.com/Euronymou5">`Euronymou5`</a> sin el que este proyecto habría sido inviable.

`[INFO]` Sistemas operativos `soportados:`

| Sistema operativo  | Soporte |
| ------------- | ------------- |
| Windows  | ✔️  |
| Linux  | ✔️  |
| MacOS | :x: |
| Android | :x: |
| Apple & IOS | :x: |

<hr>

`[INFO]` Instalación en `Windows:`

`[1]` Descarga el archivo `.zip`

`[2]` Abre la carpeta y ejecuta el `install.bat` gráficamente o desde el terminal con `start install.bat`

<hr>

`[INFO]` Instalación en `Linux:`

`[1]` Instalación con `comandos:`

    sudo apt update && sudo apt upgrade && git clone https://www.github.com/VBSInjection && cd VBSInjection && chmod +x setup.sh && chmod 777 setup.sh && sudo bash setup.sh

<hr>

`[INFO]` Hotkeys `soportadas:`

`[INFO]` Source: https://ss64.com/vb/sendkeys.html

| Hotkey  | Manera |
| ------------- | ------------- |
| SHIFT| (+ prefix) wshshell.sendkeys "`%`{F6}" |
| CTRL | (^ PREFIX) wshshell.sendkeys "`^`{escape}" |
| ALT | (% prefix) wshshell.sendkeys "`%`{F4}" |

| Hotkey | Soporte para la v1 |
| ------------- | ------------- |
| ENTER | ✔️ |
| ESCAPE | ✔️ |
| BACKSPACE | ✔️ |
| BREAK | ✔️ |
| CAPSLOCK | ✔️ |
| CLEAR | ✔️ |
| DELETE | ✔️ |
| INSERT | ✔️ |
| LEFT/RIGHT/UP/DOWN | ✔️ |
| END | ✔️ |
| F1-F16 | ✔️ |
| HELP | ✔️ |
| HOME | ✔️ |
| NUMLOCK | ✔️ |
| PGUP/PGDN | ✔️ |
| PRTSC (print screen) | ✔️ |
| SCROLLOCK | ✔️ |
| TAB | ✔️ |

<hr>

`[INFO]` Comandos:

    $ python3 main.py --help
    usage: main.py [-h] [--open OPEN] [--begin BEGIN] [--sleep SLEEP] [--write WRITE] [--keyhot KEYHOT] [--appactivate APPACTIVATE] [--message MESSAGE]
                   [--close CLOSE]
    
    options:
      -h, --help            show this help message and exit
      --open OPEN, -o OPEN  Indica el nombre y la extensión del archivo a abrir
      --begin BEGIN, -b BEGIN
                            Empieza creando el objeto de VBS en el archivo. Este argumento es necesario para empezar a crear el script
      --sleep SLEEP, -s SLEEP
                            Crea una pausa en el script de VBS del tiempo especificado en el argumento
      --write WRITE, -w WRITE
                            Escribe una tecla normal en el script VBS para que sea inyectada cuando se ejecute
      --keyhot KEYHOT, -k KEYHOT
                            Escribe una hotkey especificada en mayúscuñas en el archivo VBS
      --appactivate APPACTIVATE, -a APPACTIVATE
                            Avtiva la app indicada
      --message MESSAGE, -m MESSAGE
                            Indica el texto a mostrar en el mensaje
      --close CLOSE, -c CLOSE
                            Cierra el archivo 

<hr>

`[INFO]` Funcionamiento:

`[INFO]` Se importan las `librerías:`

```python
import os, argparse, sys
from colorama import Fore, Back
```

`[INFO]` Creamos los argumentos de `parser:`

```python
parser = argparse.ArgumentParser()

parser.add_argument('--open', '-o',
                    required=False,
                    help="Indica el nombre y la extensión del archivo a abrir")

args = parser.parse_args()
```

`[INFO]` Para crear los colores del `texto:`

```python
class color:
    RED = Fore.RED + Back.RESET
    RESET = Fore.RESET + Back.RESET
```

`[INFO]` Procesamos los argumentos con una `condicional:`

```python
if args.begin == 'true':
    file = open('inject.vbs', 'a')
    file.write('Set WshShell = WScript.CreateObject("WScript.Shell")\n')
    file.write('strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )\n')
    sys.exit()
```

<hr>

`[INFO]` Comandos de `VBS soportados:`

`[INFO]` MsgBox:

```vbs
MsgBox"Hello World",48,"Mensaje del sistema"
```

`[INFO]` Inciar el código creando el `objeto:`

```vbs
Set WshShell = WScript.CreateObject("WScript.Shell")
strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
```

`[INFO]` Iniciar una app:

```vbs
wshshell.run "cmd.exe"
wshshell.AppActivate "cmd.exe"
```

`[INFO]` Tiempo de `pausa.` En el argumento de `Python` se indica el tiempo en segundos y se `traduce a milisegundos` en el programa de `VBS:`

```vbs
WScript.sleep 1000
```

`[INFO]` Inyección de `pulsaciones del teclado:`

```vbs
wshshell.sendkeys "hello"
```

`[INFO]` Inyección de hotkeys:

```vbs
wshshell.sendkeys "{ENTER}"
```

<hr>

`[INFO]` Ejemplos de `ejecución:`

    python3 main.py -o true
    python3 main.py -b true
    python3 main.py -s 1
    python3 main.py -a cmd.exe
    python3 main.py -s 1
    python3 main.py -m Ejecutando...
    python3 main.py -w ipconfig
    python3 main.py -k ENTER
    python3 main.py -s 1
    python3 main.py -w exit
    python3 main.py -k ENTER
    python3 main.py -c true

`[INFO]` Esto generaría un programa de output `(inject.vbs)` donde estaría el siguiente `contenido:`

    Set WshShell = WScript.CreateObject("WScript.Shell")
    strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
    WScript.sleep 1000
    wshshell.run "cmd.exe"
    wshshell.AppActivate "cmd.exe"
    WScript.sleep 1000
    MsgBox"Ejecutando...",48,"Mensaje del sistema"
    wshshell.sendkeys "ipconfig"
    wshshell.sendkeys "{ENTER}"
    WScript.sleep 1000
    wshshell.sendkeys "exit"
    wshshell.sendkeys "{ENTER}"

`[INFO]` En este caso el `programa` serviría para abrir el `CMD,` ejecutar el comando `ipconfig` y salir del `terminal.`

<hr>

![imagen](https://github.com/ZombieGeeK0/VBSInjection/assets/158185295/4a05d544-0e7e-4813-b29d-684b25bcd4ad)

<hr>

## 🎖️ CONTRIBUIDORES 🎖️

- <a href="https://www.github.com/Euronymou5">Euronymou5</a>
- <a href="https://www.github.com/ZombieGeek0">ZombieGeek0</a>
    
<hr>

`[ 📬 ]` Contacta conmigo a través de `Discord` mandando una invitación a `qwfkr.`

    qwfkr
`[ 📬 ]` Si lo prefieres, mándame un correo a `3xpl017.contact@proton.me.`

    3xpl017.contact@proton.me.
