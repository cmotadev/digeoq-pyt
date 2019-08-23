# -*- coding: latin-1 -*-
import arcpy
import io
import re
from datetime import datetime


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Ferramentas DIGEOQ"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [GeradorEtiqueta]


class GeradorEtiqueta(object):
    """
    Ferramenta para gerar etiquetas Zebra
    """

    __TEMPLATE__ = u"""^XA

^FX Borda da etiqueta (para referência)
^FO1,1^GB597,235,1^FS

^FX Texto em UTF-8
^CI28

^FX CPRM Logo
^FO60,40^GFA,1190,1190,14,,N038M07,N07EL01F8,M01FFL03FE,M07FFCK0IF8,M0IFEJ01IFC,L03JF8I07JF,L0KFE001KFC,K01LF803LF,K07LFC0MF8,J01NF3MFE,J03NF3NF8,J0NFC0NFE,I03NF003NF8,I07MFE001NFC,001NF8I07NF,007MFEJ01NF8,01NFCK0NFE,03NFL03NF,0NFCM0NFC,3NF8M07NF,7NFN03NF8,::::::::::::::::::::::::::::::::::::::::::3NFN03NF,0NFCM0NFE,07NFL03NF,01NFCK0NFE,007MFEJ01NF8,003NF8I07NF,I0NFE001NFC,I03NF003NF8,J0NFC0NFC,J03NF1NF,J01NF3MFE,K07LFE1MF8,K01LF803LF,L0LF001KFC,L07JFCI0KF8,M0IFEJ01IFE,M07FFCK0IF8,M01FFL03FE,N07EM0F8,N038M07,,^FS

^FX Logo Caption
^CF0,20,15
^FO15,140^FB205,2,5,C,0^FDSERVIÇO GEOLÓGICO DO BRASIL CPRM^FS

^FX Data da Coleta
^CF0,20,20
^FO15,200^FB205,1,5,C,0^FD%(analysis_date)s^FS

^FX Third section with barcode.
^BY2,2,150
^FO240,40^B3^FD%(num_lab)s+%(weight)s^FS

^XZ"""

    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Gerador de Etiquetas"
        self.description = "Ferramenta para, a partir dos dados das alíquotas, gerar arquivos ZPL para impressão em equipamentos Zebra"
        self.canRunInBackground = True

    def getParameterInfo(self):
        """Define parameter definitions"""
        # First parameter
        param0 = arcpy.Parameter(
            displayName="Tabela ou feição vetorial",
            name="in_table",
            datatype="GPTableView",
            parameterType="Required",
            direction="Input")

        # Second parameter
        param1 = arcpy.Parameter(
            displayName="Coluna do número de laboratório",
            name="in_numlab_field",
            datatype="Field",
            parameterType="Required",
            direction="Input")

        param1.parameterDependencies = [param0.name]
        param1.filter.list = ["Text"]

        # Third parameter
        param2 = arcpy.Parameter(
            displayName="Coluna do peso da amostra (em gramas)",
            name="in_wght_field",
            datatype="Field",
            parameterType="Required",
            direction="Input")

        param2.parameterDependencies = [param0.name]
        param2.filter.list = ["Short", "Long"]

        # Fourth parameter
        param3 = arcpy.Parameter(
            displayName="Data do Pedido",
            name="in_analysis_date",
            datatype="GPDate",
            parameterType="Required",
            direction="Input")

        param3.value = datetime.now().strftime("%x %X")

        # Fifith parameter (Output)
        param4 = arcpy.Parameter(
            displayName="Output ZPL File",
            name="out_file",
            datatype="DEFile",
            parameterType="Required",
            direction="Output")

        # To define a file filter that includes .csv and .txt extensions,
        #  set the filter list to a list of file extension names
        param4.filter.list = ['zpl']

        return [param0, param1, param2, param3, param4]

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""

        # Messages
        try:
            datetime.strptime(parameters[3].valueAsText.strip(), "%x")

        except ValueError:
            parameters[3].setErrorMessage("Não é permitido inserir apenas hora. Selecione Data, apenas")

        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        in_table = parameters[0].valueAsText
        in_numlab_field = parameters[1].valueAsText
        in_wght_field = parameters[2].valueAsText
        in_analysis_date = parameters[3].valueAsText
        out_file = parameters[4].valueAsText

        # Let´s do the magic
        cnt = arcpy.GetCount_management(in_table)

        messages.addMessage(in_overwrite)

        if cnt == 0:
            messages.addMessage("Não há registros selecionados")
            return

        messages.addMessage("%s amostras selecionadas" % cnt)

        # Iteration
        with io.open(out_file, "w", encoding="utf-8") as output:
            with arcpy.da.SearchCursor(in_table,[in_numlab_field, in_wght_field]) as rows:
                for row in rows:
                    _date = in_analysis_date
                    _num_lab = re.sub('[\s\-\_]+', '', row[0])
                    _weight =  unicode(row[1]).zfill(4)

                    output.write(self.__TEMPLATE__ % ({
                        "analysis_date": _date,
                        "num_lab": _num_lab,
                        "weight": _weight
                    }))

                    output.write(u"\r\n")

        return