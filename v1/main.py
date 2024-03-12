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

parser.add_argument('--close', '-c',
                    required=False,
                    help="Cierra la edición del archivo")

args = parser.parse_args()

# hacemos la clase color
class color
  RED = Fore.RED + Back.RESET
  RESET = Fore.RESET + Back.RESET

# hacemos la función principal
def main():
  os.system('cls')
  title = '''
 ▌ ▐·▄▄▄▄· .▄▄ · ▪   ▐ ▄  ▐▄▄▄▄▄▄ . ▄▄· ▄▄▄▄▄▪         ▐ ▄ 
▪█·█▌▐█ ▀█▪▐█ ▀. ██ •█▌▐█  ·██▀▄.▀·▐█ ▌▪•██  ██ ▪     •█▌▐█   By ZombieGeek0, Copyright© 2024
▐█▐█•▐█▀▀█▄▄▀▀▀█▄▐█·▐█▐▐▌▪▄ ██▐▀▀▪▄██ ▄▄ ▐█.▪▐█· ▄█▀▄ ▐█▐▐▌   GitHub: https://www.github.com/ZombieGeek0/VBSInjection
 ███ ██▄▪▐█▐█▄▪▐█▐█▌██▐█▌▐▌▐█▌▐█▄▄▌▐███▌ ▐█▌·▐█▌▐█▌.▐▌██▐█▌
. ▀  ·▀▀▀▀  ▀▀▀▀ ▀▀▀▀▀ █▪ ▀▀▀• ▀▀▀ ·▀▀▀  ▀▀▀ ▀▀▀ ▀█▄▀▪▀▀ █▪
'''
    
    if args.open == 'true':
      file = open('inject.vbs', 'w')
      sys.exit()

    if args.begin:
      file.write('Set WshShell = WScript.CreateObject("WScript.Shell")\n')
      file.write('strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )\n')
      sys.exit()

    elif args.sleep:
      file.write(f'WScript.sleep {args.sleep}\n')
      sys.exit()

    elif args.write:
      file.write(f'wshshell.sendkeys "{args.write}"\n')
      sys.exit()

    elif args.keyhot:
      file.write('wshshell.sendkeys "{' + args.keyhoy + '}"\n')
      sys.exit()

    elif args.appactivate:
      file.write(f'wshshell.run "{args.appactivate}"\n')
      file.write(f'wshshell.AppActivate "{args.appactivate}"\n')
      sys.exit()

    elif args.close == 'true':
      file.close()
      sys.exit()

    else:
      print(f'{color.RED}\n[>]: Error: La expresión {choice} es inválida o no existe\n')



main()
