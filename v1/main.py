# importamos las librerías necesarias
import os, argparse, sys
from colorama import Fore, Back

# cremos los argumentos de parser
parser = argparse.ArgumentParser()

parser.add_argument('--open', '-o',
                    required=False,
                    help="Indica el nombre y la extensión del archivo a abrir")

parser.add_argument('--begin', '-b',
                    required=False,
                    help="Empieza creando el objeto de VBS en el archivo. Este argumento es necesario para empezar a crear el script")

parser.add_argument('--sleep', '-s',
                    required=False,
                    help="Crea una pausa en el script de VBS del tiempo especificado en el argumento")

parser.add_argument('--write', '-w',
                    required=False,
                    help="Escribe una tecla normal en el script VBS para que sea inyectada cuando se ejecute")

parser.add_argument('--keyhot', '-k',
                    required=False,
                    help="Escribe una hotkey especificada en mayúscuñas en el archivo VBS")

parser.add_argument('--appactivate', '-a',
                    required=False,
                    help="Avtiva la app indicada")

parser.add_argument('--message', '-m',
                    required=False,
                    help="Indica el texto a mostrar en el mensaje")

parser.add_argument('--close', '-c',
                    required=False,
                    help="Cierra el archivo")

args = parser.parse_args()

# hacemos la clase color
class color:
    RED = Fore.RED + Back.RESET
    RESET = Fore.RESET + Back.RESET

# creamos la variable del archivo
file = ''

# hacemos la función principal
def main():
    os.system('cls')

    if args.open == 'true':
        file = open('inject.vbs', 'x')
        sys.exit()

    if args.begin == 'true':
        file = open('inject.vbs', 'a')
        file.write('Set WshShell = WScript.CreateObject("WScript.Shell")\n')
        file.write('strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )\n')
        sys.exit()

    elif args.sleep:
        wait = int(args.sleep) * 1000
        file = open('inject.vbs', 'a')
        file.write(f'WScript.sleep {wait}\n')
        sys.exit()

    elif args.write:
        file = open('inject.vbs', 'a')
        file.write(f'wshshell.sendkeys "{args.write}"\n')
        sys.exit()

    elif args.keyhot:
        file = open('inject.vbs', 'a')
        file.write('wshshell.sendkeys "{' + args.keyhot + '}"\n')
        sys.exit()

    elif args.appactivate:
        file = open('inject.vbs', 'a')
        file.write(f'wshshell.run "{args.appactivate}"\n')
        file.write(f'wshshell.AppActivate "{args.appactivate}"\n')
        sys.exit()

    elif args.message:
      file = open('inject.vbs', 'a')
      file.write('MsgBox"",46,""')
      sys.exit()

    elif args.close:
        file.close()
        sys.exit()

    else:
        print(f'{color.RED}\n[>]: Error: Verifica si has añadido todos los parámetros necesarios\n')
        sys.exit()


# inciamos el menú
main()
