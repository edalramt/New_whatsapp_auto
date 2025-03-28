
---

## **6. Preparar Visual Studio Code para Live Share**

### **a. Instalar la extensión Live Share**
1. Abre Visual Studio Code.
2. Ve a la pestaña de extensiones (Ctrl+Shift+X).
3. Busca "Live Share" e instálala.

### **b. Configurar `launch.json`**
Si deseas que los usuarios puedan ejecutar el proyecto directamente desde Visual Studio Code, configura el archivo `.vscode/launch.json`. Este archivo permite ejecutar y depurar el proyecto fácilmente.

#### **Pasos para configurar `launch.json`:**
1. Crea una carpeta llamada `.vscode` en la raíz del proyecto (si no existe).
2. Dentro de `.vscode`, crea un archivo llamado `launch.json`.
3. Agrega la siguiente configuración al archivo:

```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Python: wsp_esteroides_pro",
            "type": "python",
            "request": "launch",
            "program": "${workspaceFolder}/wps_esteroides_pro/wsp_esteroides_pro.py",
            "console": "integratedTerminal"
        }
    ]
}