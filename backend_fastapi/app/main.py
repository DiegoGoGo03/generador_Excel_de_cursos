from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, StreamingResponse
import shutil
import os
from generar_excel import procesar_archivo, generar_excel_en_memoria

app = FastAPI()

@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
  file_location = f"temp/{file.filename}"
  os.makedirs("temp", exist_ok=True)
  with open(file_location, "wb") as buffer:
    shutil.copyfileobj(file.file, buffer)
    
  #Llamar la función que genera el Excel
  output_excel = procesar_archivo(file_location)
  return FileResponse(output_excel, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename="resultado.xlsx")

@app.get("/generar-excel/")
def generar_excel_directo():
  output = generar_excel_en_memoria()
  return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": "attachment; filename=enrollment_groups.xlsx"})