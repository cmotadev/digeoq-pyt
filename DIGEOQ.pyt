# -*- coding: latin-1 -*-
import arcpy
import io
import re
import win32print as prn

from datetime import datetime
from os import path
from tempfile import TemporaryFile


def get_zebra_printers():
    """
    :return:
    """
    return [name for flags, descr, name, cmnts in prn.EnumPrinters(prn.PRINTER_ENUM_LOCAL) if
            name.find("ZDesigner") >= 0]


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

        param4 = arcpy.Parameter(
            displayName="Selecionar Padrão de Etiqueta",
            name="in_template",
            datatype="string",
            parameterType="Required",
            direction="Input")

        param4.filter.type = "ValueList"
        param4.filter.list = ["digeoq.zpl"]

        param5 = arcpy.Parameter(
            displayName="Selecionar Impressora",
            name="in_printer",
            datatype="string",
            parameterType="Required",
            direction="Input")

        param5.filter.type = "ValueList"
        param5.filter.list = get_zebra_printers()

        return [param0, param1, param2, param3, param4, param5]

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
            try:
                datetime.strptime(parameters[3].valueAsText.strip(), "%x %X")

            except ValueError:
                parameters[3].setErrorMessage("Não é permitido inserir apenas hora. Selecione Data ou Data e Hora")

        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        in_table = parameters[0].valueAsText
        in_numlab_field = parameters[1].valueAsText
        in_wght_field = parameters[2].valueAsText
        in_analysis_date = datetime.strftime(parameters[3].value, "%d/%m/%Y")
        in_template = path.join(path.dirname(path.realpath(__file__)), parameters[4].valueAsText)
        in_printer = parameters[5].valueAsText

        # Let´s do the magic
        if not path.exists(in_template):
            messages.setErrorMessage("Arquivo de template [%s] ZPL inexistente. Entrar em contato com a DIGEOP")
            return

        cnt = arcpy.GetCount_management(in_table)

        if cnt == 0:
            messages.addMessage("Não há registros selecionados")
            return

        messages.addMessage("%s amostras selecionadas" % cnt)

        # Lê o template
        with io.open(in_template, "r", encoding="utf-8") as template:
            _template = "".join(template.readlines())
            raw_data = []

            # Adiciona os atributos e grava no tempfile
            with arcpy.da.SearchCursor(in_table, [in_numlab_field, in_wght_field]) as rows:
                for row in rows:
                    # Data da análise
                    _date = in_analysis_date

                    # Tratar número de laboratório
                    _num_lab = re.sub('[\s\-\_]+', '', row[0]).upper()

                    if not re.search("^[A-Z]{3}[0-9]{3}$", _num_lab):
                        messages.addWarningMessage(u"A alíquota %s está como nome fora de padrão. O padrão aceito é AAA111. A etiqueta não será produzida" % _num_lab)
                        continue

                    # Tratar peso da amostra
                    if int(row[1]) >= 10000:
                        messages.addWarningMessage(u"A alíquota %s possui peso maior que 9999 gramas. A etiqueta não será produzida" % _num_lab)
                        continue

                    _weight = unicode(row[1]).zfill(4)

                    # Imprimir etiqueta
                    raw_data.append(_template % ({"analysis_date": _date, "num_lab": _num_lab, "weight": _weight}))

            # Print this
            try:
                _printer = prn.OpenPrinter(in_printer)
                job = prn.StartDocPrinter(_printer, 1, ("ZPLII data from ArcMap", None, "RAW"))

                try:
                    prn.StartPagePrinter(_printer)
                    prn.WritePrinter(_printer, u"\r\n".join(raw_data))
                    prn.EndPagePrinter(_printer)
                finally:
                    prn.EndDocPrinter(_printer)
            finally:
                prn.ClosePrinter(_printer)

        return
