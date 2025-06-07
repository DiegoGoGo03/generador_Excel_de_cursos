import pandas as pd
from io import BytesIO
import os

#Lista de diccionarios fija (si luego quieres que esto venga dinámico, podemos adapartarlo)
enrollmentGroups = [
    {
        "courseId": 149,
        "courseName": "Fundamentos Tekla Structures Acero",
        "groups": [
            {"groupId": 1618, "groupName": "PROGRAMA ESTUDIANTES 2023"},
            {"groupId": 1619, "groupName": "PROGRAMA PROFESORES 2023"},
            {"groupId": 2569, "groupName": "PROGRAMA CLIENTES TEKLA 2024"},
            {"groupId": 2567, "groupName": "PROGRAMA ESTUDIANTES 2024"},
            {"groupId": 2581, "groupName": "PROGRAMA PROFESORES 2024"},
            {"groupId": 3785, "groupName": "PROGRAMA ESTUDIANTES 2025"},
            {"groupId": 3786, "groupName": "PROGRAMA PROFESORES 2025"}
        ]
    },
    {
        "courseId": 98,
        "courseName": "Fundamentos Tekla Structures Hormigón (Antiguo)",
        "groups": [
            {"groupId": 1621, "groupName": "PROGRAMA ESTUDIANTES 2023"},
            {"groupId": 1620, "groupName": "PROGRAMA PROFESORES 2023"},
            {"groupId": 2570, "groupName": "PROGRAMA CLIENTES TEKLA 2024"},
            {"groupId": 2566, "groupName": "PROGRAMA ESTUDIANTES 2024"},
            {"groupId": 2582, "groupName": "PROGRAMA PROFESORES 2024"}
        ]
    },
    {
        "courseId": 195,
        "courseName": "Fundamentos Tekla Structures Hormigón",
        "groups": [
            {"groupId": 2576, "groupName": "PROGRAMA CLIENTES TEKLA 2024"},
            {"groupId": 2568, "groupName": "PROGRAMA ESTUDIANTES 2024"},
            {"groupId": 2840, "groupName": "PROGRAMA PROFESORES 2024"}
        ]
    },
    {
        "courseId": 240,
        "courseName": "Curso Modelado de Estructuras con Tekla Structures - EUDE",
        "groups": [
            {"groupId": 3483, "groupName": "PROGRAMA EUDE 2025"}
        ]
    },
    {
        "courseId": 76,
        "courseName": "Teoría y cálculo de uniones metálicas con IDEA STATICA",
        "groups": [
            {"groupId": 1622, "groupName": "PROGRAMA ESTUDIANTES 2023"},
            {"groupId": 1623, "groupName": "PROGRAMA PROFESORES 2023"},
            {"groupId": 2561, "groupName": "PROGRAMA ESTUDIANTES 2024"},
            {"groupId": 2583, "groupName": "PROGRAMA PROFESORES 2024"}
        ]
    },
    {
        "courseId": 141,
        "courseName": "Teoría y cálculo de elementos HA con IDEA STATICA",
        "groups": [
            {"groupId": 1625, "groupName": "PROGRAMA ESTUDIANTES 2023"},
            {"groupId": 1624, "groupName": "PROGRAMA PROFESORES 2023"},
            {"groupId": 2563, "groupName": "PROGRAMA ESTUDIANTES 2024"},
            {"groupId": 2584, "groupName": "PROGRAMA PROFESORES 2024"}
        ]
    },
    {
        "courseId": 135,
        "courseName": "Análisis y diseño de edificaciones con Tekla Structural Designer",
        "groups": [
            {"groupId": 1626, "groupName": "PROGRAMA ESTUDIANTES 2023"},
            {"groupId": 1627, "groupName": "PROGRAMA PROFESORES 2023"},
            {"groupId": 2562, "groupName": "PROGRAMA ESTUDIANTES 2024"},
            {"groupId": 2585, "groupName": "PROGRAMA PROFESORES 2024"}
        ]
    },
    {
        "courseId": 81,
        "courseName": "Common Data Environment con Trimble Connect (antiguo)",
        "groups": [
            {"groupId": 1629, "groupName": "PROGRAMA ESTUDIANTES 2023"},
            {"groupId": 1628, "groupName": "PROGRAMA PROFESORES 2023"},
            {"groupId": 1628, "groupName": "PROGRAMA PROFESORES 2023"}
        ]
    },
    {
        "courseId": 96,
        "courseName": "Optimización de flujos BIM con Trimble Connect",
        "groups": [
            {"groupId": 1630, "groupName": "PROGRAMA ESTUDIANTES 2023"},
            {"groupId": 1631, "groupName": "PROGRAMA PROFESORES 2023"},
            {"groupId": 2578, "groupName": "PROGRAMA CLIENTES TEKLA 2024"},
            {"groupId": 2564, "groupName": "PROGRAMA ESTUDIANTES 2024"},
            {"groupId": 2587, "groupName": "PROGRAMA PROFESORES 2024"}
        ]
    },
    {
        "courseId": 174,
        "courseName": "Common Data Environment con Trimble Connect",
        "groups": [
            {"groupId": 2518, "groupName": "23_ADM"},
            {"groupId": 2517, "groupName": "23_CON"},
            {"groupId": 2519, "groupName": "23_FULL"},
            {"groupId": 2520, "groupName": "23_OWN"},
            {"groupId": 2565, "groupName": "PROGRAMA ESTUDIANTES 2024"},
            {"groupId": 2586, "groupName": "PROGRAMA PROFESORES 2024"}
        ]
    },
    {
        "courseId": 116,
        "courseName": "Detallado de Elementos Prefabricados en Tekla Structures",
        "groups": [
            {"groupId": 2577, "groupName": "PROGRAMA CLIENTES TEKLA 2024"}
        ]
    },
    {
        "courseId": 89,
        "courseName": "Componentes Personalizados en Tekla Structures",
        "groups": [
            {"groupId": 2572, "groupName": "PROGRAMA CLIENTES TEKLA 2024"}
        ]
    },
    {
        "courseId": 113,
        "courseName": "Editor de cuadros en Tekla Structures",
        "groups": [
            {"groupId": 2573, "groupName": "PROGRAMA CLIENTES TEKLA 2024"}
        ]
    },
    {
        "courseId": 114,
        "courseName": "Gestión de la numeración de Tekla Structures",
        "groups": [
            {"groupId": 2574, "groupName": "PROGRAMA CLIENTES TEKLA 2024"}
        ]
    },
    {
        "courseId": 115,
        "courseName": "Macros de Construsoft para Tekla Structures",
        "groups": [
            {"groupId": 2575, "groupName": "PROGRAMA CLIENTES TEKLA 2024"}
        ]
    },
    {
        "courseId": 2,
        "courseName": "Strusite",
        "groups": [
            {"groupId": 33, "groupName": "PROGRAMA ESTUDIANTES 2023"},
            {"groupId": 34, "groupName": "PROGRAMA PROFESORES 2023"}       
        ]
    }
]

def generar_excel_en_memoria():
  table_data = []
  for course in enrollmentGroups:
    for group in course["groups"]:
        table_data.append({
           "courseId": course["courseId"],
           "courseName": course["courseName"],
           "groupId": group["groupId"],
           "groupName": group["groupName"]
        })

  df = pd.DataFrame(table_data)

  #Guardar en memoria
  output = BytesIO()
  df.to_excel(output, index=False)
  output.seek(0)
  return output


def procesar_archivo(ruta_archivo: str) -> str:
  #Leer el Excel de entrada
  df = pd.read_excel(ruta_archivo)

  #Aquí puedes aplicar tu lógica de procesamiento.
  df["NuevaColumna"] = "Procesado"

  #Crear carpeta de salida
  output_dir = "output"
  os.makedirs(output_dir, exist_ok=True)
  output_path = os.path.join(output_dir, "resultado.xlsx")

  #Guardar resultado
  df.to_excel(output_path, index=False)
  return output_path