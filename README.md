# Py_SecretSanta

**Autor:** Hugo Quispe

## Descripción

Py_SecretSanta es una herramienta de consola en Python diseñada para realizar un sorteo de intercambio de regalos de Navidad entre familiares, amigos o compañeros de trabajo. El sistema garantiza que se cumplan todas las reglas del sorteo, evitando que un participante se regale a sí mismo y minimizando ciclos cortos entre los participantes.

## Requisitos

- Python instalado en la máquina local.

## Instalación

Para descargar e instalar el proyecto desde el repositorio público en GitHub, sigue estos pasos:

1. Clona el repositorio en tu máquina local:
    ```bash
    git clone https://github.com/hquispeon/Py_SecretSanta.git
    ```

2. Navega al directorio del proyecto:
    ```bash
    cd Py_SecretSanta
    ```

## Uso

Ejecuta el programa desde la consola de Python. Asegúrate de que tienes permisos de lectura y escritura en el directorio donde se ha descargado el proyecto, ya que este generará automáticamente la base de datos en un archivo local.

El código del proyecto te ofrece las siguientes alternativas:
```python
print("\n--- Menú Principal ---")
print("1. Agregar o editar participantes")
print("2. Listar participantes actual")
print("3. Ver sorteos por año")
print("4. Preguntar por participante")
print("5. Realizar Sorteo", anio_actual)
print("6. Ilustrar amigos por año")
print("7. Salir")
opcion = input("Elige una opción: ")
```

## Instalación Automática de Librerías
El proyecto reconocerá automáticamente las librerías que no tienes instaladas e intentará instalarlas. Las librerías utilizadas son:
openpyxl
PIL
os
pytest
unittest.mock

## Ejecución de Pruebas Unitarias
El proyecto incluye un archivo de pruebas unitarias test_sorteo.py que realiza las pruebas más importantes. Para ejecutar las pruebas, ingresa el siguiente comando en la consola:
    ```bash
	pytest -s test_sorteo.py
    ```

## Sugerencias para Futuros Trabajos
Algunas ideas para mejorar y extender el proyecto incluyen:
- Desarrollar una Aplicación Web: Transforma la herramienta de consola en una aplicación web accesible desde cualquier dispositivo.
- Integración con WhatsApp: Permite consultar y realizar sorteos a través de WhatsApp utilizando la API de WhatsApp.
- Mejoras en la Interfaz de Usuario: Implementa una interfaz gráfica para hacer la experiencia del usuario más amigable.
- Soporte Multilenguaje: Añadir soporte para múltiples idiomas, adaptando la interfaz y las respuestas según la localización del usuario.

## Agradecimientos
Este proyecto fue desarrollado con el apoyo de herramientas de inteligencia artificial para la generación de código. Sin embargo, toda la lógica, el modelo y el refinamiento del código fueron realizados por mí persona.

## Licencia
Este proyecto está licenciado bajo la Licencia MIT. Para más detalles, mira el archivo LICENSE.md

## Contacto
Para más información o consultas, puedes contactar al autor.
Proyecto en GitHub: https://github.com/hquispeon
