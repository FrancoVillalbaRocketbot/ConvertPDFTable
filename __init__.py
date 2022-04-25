# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
import os
import sys
base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'ConvertPDFTable' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)

import tabula
from PyPDF2 import PdfFileReader
"""
    Obtengo el modulo que fueron invocados
"""
module = GetParams("module")

if module == "pdftocsv":
    try:
        pdf_file = GetParams("pdf")
        csv_path = GetParams("path")
        paginas = GetParams("page")
        coordenadas = GetParams("coordinates")
        area = [coordenadas]
        reader = PdfFileReader(pdf_file)
        number_of_pages = reader.numPages
        page = reader.pages[0]
        size = page.mediaBox

        if not paginas:
            paginas = "all"
            print(paginas)

        if coordenadas:
            dfs=tabula.read_pdf(pdf_file,area = area, pages=paginas, stream = True,guess=True, silent=True, encoding = "latin-1")
            if dfs:
                dfs[0].to_csv(csv_path, mode = "a", index=False, encoding='latin-1')
        else:
            dfs1 = tabula.read_pdf(pdf_file, area=[0,0,size.getHeight()/3,size.getWidth()], pages=paginas, stream = True,guess=False, silent=True, encoding = "latin-1")
            dfs2 = tabula.read_pdf(pdf_file, area=[size.getHeight()/3,0,2*size.getHeight()/3,size.getWidth()], pages=paginas, stream = True, guess=True, silent=True, encoding = "latin-1")
            dfs3 = tabula.read_pdf(pdf_file, area=[2*size.getHeight()/3,0,3*size.getHeight()/3,size.getWidth()], pages=paginas, stream = True, guess=True, silent=True, encoding = "latin-1")
            for dfs in [dfs3,dfs2,dfs1]:
                if dfs:
                    for d in dfs:
                        d.to_csv(csv_path, mode = "a", index=False, encoding='latin-1')
    except Exception as e:
            PrintException()
            raise e

if module == "pdftotsv":
    pdf_file = GetParams("pdf")
    tsv_path = GetParams("path")
    paginas = GetParams("page")

    if not tsv_path.endswith(".tsv"):
        tsv_path += ".tsv"
    
    if paginas is None or paginas == "":
        paginas = "all"

    pdf_file.replace('\\', '/')
    tsv_path.replace('\\', '/')

    try:
        tabula.convert_into(pdf_file, tsv_path, output_format="tsv", pages=paginas)
    except Exception as e:
        PrintException()
        raise e

if module == "pdftojson":
    pdf_file = GetParams("pdf")
    json_path = GetParams("path")
    paginas = GetParams("page")
    _var = GetParams("_var")

    if not json_path.endswith(".json"):
        json_path += ".json"
    try:
        if paginas is None or len(paginas) == 0:
            paginas = "all"

        pdf_file.replace('\\', '/')
        json_path.replace('\\', '/')

        if len(_var)>0:
            df = tabula.read_pdf(pdf_file, output_format="json", pages=paginas)
            SetVar(_var, df)

        tabula.convert_into(pdf_file, json_path, output_format="json", pages=paginas)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\u2193\x1B[" + "0m")
        PrintException()
        raise e
