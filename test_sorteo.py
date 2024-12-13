# Ejecutar el archivo con el siguiente comando
# pytest -s test_sorteo.py

import subprocess
import Sorteo
from datetime import datetime

def verificar_e_instalar_libreria(nombre_libreria):
    try:
        # Intentar importar la librería
        __import__(nombre_libreria)
        # print(f"La librería '{nombre_libreria}' ya está instalada.")
    except ImportError:
        print(f"La librería '{nombre_libreria}' no está instalada. Procediendo a instalarla...")
        # Intentar instalar la librería con pip
        subprocess.check_call([sys.executable, "-m", "pip", "install", nombre_libreria])
        print(f"La librería '{nombre_libreria}' ha sido instalada exitosamente.")

# Verificar e instalar 'openpyxl' si es necesario
verificar_e_instalar_libreria("pytest")
verificar_e_instalar_libreria("unittest.mock")

import pytest
from unittest.mock import MagicMock, patch

@pytest.fixture
def mock_workbook():
    workbook = MagicMock()
    workbook.sheetnames = []
    participante_sheet = MagicMock()
    workbook.create_sheet.return_value = participante_sheet
    workbook.__getitem__.return_value = participante_sheet
    return workbook, participante_sheet

@pytest.fixture
def mock_workbook_sorteo():
    workbook = MagicMock()
    participante_sheet = MagicMock()
    historia_sheet = MagicMock()
    workbook.__getitem__.side_effect = lambda x: participante_sheet if x == Sorteo.config.PesParticipante else historia_sheet
    return workbook, participante_sheet, historia_sheet

@patch('builtins.input', side_effect=["Ana, María, Juan Carlos, Julia"])
def test_agregar_participantes(mock_input, mock_workbook):
    workbook, participante_sheet = mock_workbook

    Sorteo.agregar_o_editar_participantes(workbook, "mocked_file.xlsx")

    workbook.create_sheet.assert_called_once_with(Sorteo.config.PesParticipante)
    participante_sheet.append.assert_any_call([Sorteo.config.ColNombre])
    participante_sheet.append.assert_any_call(["Ana"])
    participante_sheet.append.assert_any_call(["María"])
    participante_sheet.append.assert_any_call(["Juan Carlos"])
    participante_sheet.append.assert_any_call(["Julia"])
    workbook.save.assert_called_once_with("mocked_file.xlsx")
    
    print("Prueba unitaria de agregar_participantes exitosa")

@patch('builtins.input', side_effect=["Ana, María, Juan Carlos, Julia"])
def test_listar_participantes(mock_input, mock_workbook):
    workbook, participante_sheet = mock_workbook

    Sorteo.agregar_o_editar_participantes(workbook, "mocked_file.xlsx")

    participante_sheet.iter_rows.return_value = [
        ('Ana',),
        ('María',),
        ('Juan Carlos',),
        ('Julia',),
    ]

    nombres = Sorteo.listar_participantes(workbook)
    assert nombres == ["Ana", "María", "Juan Carlos", "Julia"]

    print("Prueba unitaria de listar_participantes exitosa")

@patch('builtins.input', side_effect=["s"])  # Simular confirmación de guardar
def test_realizar_sorteo(mock_input, mock_workbook_sorteo):
    workbook, participante_sheet, historia_sheet = mock_workbook_sorteo

    # Configurar datos de prueba
    participante_sheet.iter_rows.return_value = [
        ('Ana',),
        ('María',),
        ('Juan Carlos',),
        ('Julia',),
    ]

    anio_actual = datetime.now().year

    historia_sheet.iter_rows.return_value = [
        (anio_actual - 1, 'Ana', 'María'),
        (anio_actual - 1, 'Juan Carlos', 'Julia'),
    ]

    resultados = Sorteo.realizar_sorteo(workbook, anio_actual, "mocked_file.xlsx")

    # Validar que los participantes no se regalen a sí mismos
    for participante, amigo in resultados:
        assert participante != amigo, f"Error: {participante} se regaló a sí mismo."

    # Validar que no haya ciclos de dos o tres personas
    regalos = dict(resultados)
    for participante in regalos:
        amigo1 = regalos[participante]
        amigo2 = regalos[amigo1]
        if amigo2 == participante:
            assert False, f"Error: Se detectó un ciclo de 2 personas entre {participante} y {amigo1}."
        amigo3 = regalos.get(amigo2)
        if amigo3 == participante:
            assert False, f"Error: Se detectó un ciclo de 3 personas entre {participante}, {amigo1} y {amigo2}."

    print("Prueba unitaria de realizar_sorteo exitosa")

@patch('builtins.input', side_effect=["2023"])  # Preguntar por un año que no está en el historial
def test_ver_sorteos_por_anio_no_historial(mock_input, mock_workbook_sorteo):
    workbook, _, historia_sheet = mock_workbook_sorteo

    anio_actual = datetime.now().year

    historia_sheet.iter_rows.return_value = [
        (Sorteo.config.ColAnio, Sorteo.config.ColParticipante, Sorteo.config.ColAmigo),  # Fila de encabezado
        (anio_actual - 2, 'Ana', 'María'),
        (anio_actual - 2, 'María', 'Juan Carlos'),
        (anio_actual - 2, 'Juan Carlos', 'Ana'),
    ]

    resultado = Sorteo.ver_sorteos_por_anio(workbook, 2023, datetime.now().year)
    assert resultado == [], "El resultado debería ser una lista vacía, ya que no hay sorteos para el año 2023."

    print("Prueba unitaria de ver_sorteos_por_anio para año no en el historial exitosa")

@patch('builtins.input', side_effect=[str(datetime.now().year)])  # Preguntar por el año actual
def test_ver_sorteos_por_anio_anio_actual(mock_input, mock_workbook_sorteo):
    workbook, _, historia_sheet = mock_workbook_sorteo

    anio_actual = datetime.now().year

    historia_sheet.iter_rows.return_value = [
        (Sorteo.config.ColAnio, Sorteo.config.ColParticipante, Sorteo.config.ColAmigo),  # Fila de encabezado
        (anio_actual, 'Ana', 'María'),
        (anio_actual, 'María', 'Juan Carlos'),
        (anio_actual, 'Juan Carlos', 'Ana'),
    ]

    resultado = Sorteo.ver_sorteos_por_anio(workbook, anio_actual, anio_actual)
    expected_result = [(anio_actual, 'Ana', 'María'), (anio_actual, 'María', 'Juan Carlos'), (anio_actual, 'Juan Carlos', 'Ana')]
    assert resultado == expected_result, f"El resultado debería ser la lista de sorteos para el año {anio_actual}."

    print("Prueba unitaria de ver_sorteos_por_anio para el año actual exitosa")

@patch('builtins.input', side_effect=[str('Manuel')]) # Preguntar por historial de Ana
def test_preguntar_por_participante_no_existe(mock_imput, mock_workbook_sorteo):
    workbook, _, historia_sheet = mock_workbook_sorteo

    anio_actual = datetime.now().year

    historia_sheet.iter_rows.return_value = [
        (Sorteo.config.ColAnio, Sorteo.config.ColParticipante, Sorteo.config.ColAmigo),  # Fila de encabezado
        (anio_actual - 2, 'Ana', 'María'),
        (anio_actual - 2, 'María', 'Juan Carlos'),
        (anio_actual - 2, 'Juan Carlos', 'Ana'),
        (anio_actual - 1, 'Juan Carlos', 'María'),
        (anio_actual - 1, 'María', 'Ana'),
        (anio_actual - 1, 'Ana', 'Juan Carlos'),
    ]

    resultado = Sorteo.preguntar_por_participante(workbook)
    expected_result = []
    assert resultado == expected_result, f"El resultado debería ser una lista vacía, ya que no hay sorteos para el año 2023."

    print("Prueba unitaria de preguntar_por_participante para un participante que no existe exitosa")

@patch('builtins.input', side_effect=[str('Ana')]) # Preguntar por historial de Ana
def test_preguntar_por_participante_no_existe(mock_imput, mock_workbook_sorteo):
    workbook, _, historia_sheet = mock_workbook_sorteo

    anio_actual = datetime.now().year

    historia_sheet.iter_rows.return_value = [
        (Sorteo.config.ColAnio, Sorteo.config.ColParticipante, Sorteo.config.ColAmigo),  # Fila de encabezado
        (anio_actual - 2, 'Ana', 'María'),
        (anio_actual - 2, 'María', 'Juan Carlos'),
        (anio_actual - 2, 'Juan Carlos', 'Ana'),
        (anio_actual - 1, 'Juan Carlos', 'María'),
        (anio_actual - 1, 'María', 'Ana'),
        (anio_actual - 1, 'Ana', 'Juan Carlos'),
    ]

    resultado = Sorteo.preguntar_por_participante(workbook)
    expected_result = [(anio_actual - 2, 'María'), (anio_actual - 1, 'Juan Carlos')]
    assert resultado == expected_result, f"El resultado debería ser una lista con los amigos a los que regaló Ana."

    print("Prueba unitaria de preguntar_por_participante para un participante que existe exitosa")

if __name__ == '__main__':
    pytest.main(["-s"])