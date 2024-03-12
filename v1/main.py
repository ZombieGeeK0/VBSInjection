# importamos las librerías necesarias
import os, argparse
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

parser.add_argument('--hotkey', '-h',
                    required=False,
                    help="Escribe una hotkey especificada en mayúscuñas en el archivo VBS")

parser.add_argument('--close', '-c',
                    required=False,
                    help="Cierra la edición del archivo")

args = parser.parse_args()

# args.open

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
  print(f'{color.RED}{title}')

  file = open('injection.vbs', 'w')
    
  while True:

    choice = input(f'\n{color.RED}[>]:  ')
    
    if 'start' in choice:
      file.write('Set WshShell = WScript.CreateObject("WScript.Shell")\n')
      file.write('strName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )'
      

    else:
      print(f'{color.RED}[>]: Error: La expresión {choice} es inválida o no existe')


  file.close()



main()
