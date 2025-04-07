import pyautogui as pg
import time

print("Mueve el mouse sobre el botón de adjuntar (📎) y espera...")
time.sleep(5)  # Te da 5 segundos para mover el mouse

x, y = pg.position()
print(f"Coordenadas del botón 📎: {x}, {y}")

print("Ahora mueve el mouse sobre la opción 'Documento' y espera...")
time.sleep(5)  # Te da otros 5 segundos

x, y = pg.position()
print(f"Coordenadas de 'Documento': {x}, {y}")
