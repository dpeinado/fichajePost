# -*- coding: utf-8 -*-
__author__ = "dpm"

import datetime

import os
import xlwt
import xlrd

class EntradasImpares(Exception):
    pass

class NoEmpiezaPorEntrada(Exception):
    pass

class NoAcabaPorEntrada(Exception):
    pass

class MalFormadoSinRemedio(Exception):
    pass

ZERO_TIME = datetime.timedelta()


class fichajesPost(object):
    def __init__(self, outFile, inPath):
        self.outFile = outFile
        self.inPath = inPath
        self.workbooks = [ os.path.join(inPath, _) for _ in os.listdir(inPath) if _.endswith('.xls')]
        self.outWorkbook = xlwt.Workbook()
        self.outSheet = self.outWorkbook.add_sheet('Salida')

    def printWorkbooks(self):
        print self.workbooks

    def limpiaEntradas(self, miEnt):
        limpias = []
        if (len(miEnt) == 1) and ('Sin marcajes' in miEnt[0]):
            return
        for ent in miEnt:
            partes = ent.split()
            val = (" ".join(partes[0:-1])).upper()
            if "ENTRADA" in val:
                tipo = "A"
            elif "COMIDA" in val:
                tipo = "B"
            elif "83" in val:
                tipo = "C"
            elif "VISITA" in val:
                tipo = "D"
            elif "TRABAJO" in val:
                tipo = "E"
            elif "INEXCUSABLE" in val:
                tipo = "F"
            elif "SINDICALES" in val:
                tipo = "G"
            elif "ASAMBLEA" in val:
                tipo = "H"
            elif "CURSO" in val:
                tipo = "I"
            elif "DEBER" in val:
                tipo = "J"

            horatmp = partes[-1].split(':')
            hora = datetime.timedelta(hours=int(horatmp[0]), minutes=int(horatmp[1]), seconds=int(horatmp[2]))
            limpias.append((tipo, hora))
        return limpias

    def bienFormada(self, me):
        cuantos = len(me)
        index = 1
        maximo = cuantos -1

        if me[0][0] == "A" and me[-1][0] == "A":
            pass
        elif (me[0][0] == "C" and me[1][0] =="C"):
            index += 2
        elif (me[-1][0] == "C" and me[-2][0] =="C"):
            maximo -= 2
        else:
            return False

        if cuantos == 2:
            if me[0][0] == "A" and me[1][0] == "A":
                return True
            else:
                return False


        while (index < maximo):
            if me[index][0] == me[index+1][0]:
                index += 2
            else:
                return False
        return True

    def procesaBienFormada(self, me):
        comida = ZERO_TIME
        art83 = ZERO_TIME
        otros = ZERO_TIME
        error = ZERO_TIME

        jornada = me[-1][1] - me[0][1]

        index = 1
        maximo = len(me)-1

        if me[0][0] == "C":
            art83 += me[1][1] - me[0][1]
            index = 3
        if me[-1][0] == "C":
            art83 += me[-1][1] - me[-2][1]
            maximo -= 2

        while (index < maximo):
            elapsedtime = me[index+1][1] - me[index][1]
            tipo = me[index][0]
            tiempo = me[index][1]
            if tipo == "A":
                pass
            elif tipo == "C":
                art83 += elapsedtime
            elif tipo == "B":
                comida += elapsedtime
            else:
                otros += elapsedtime
            index += 2

        jornada -= (art83+comida+otros)

        return {'jornada': jornada, 'comida': comida, 'art83': art83, 'otros': otros, 'error': error}

    def procesaMalFormada(self, me):
        comida = ZERO_TIME
        art83 = ZERO_TIME
        otros = ZERO_TIME
        error = ZERO_TIME
        jornada = ZERO_TIME

        total = me[-1][1] - me[0][1]
        cuantos = len(me)

        if (cuantos == 2):
            if me[0][0] == "A" or me[1][0] == "A":
                jornada = total

            else:
                error = total
            return {'jornada': jornada, 'comida': comida, 'art83': art83, 'otros': otros, 'error': error}

        if me[0][0] == "A" and me[-1][0] == "A" and cuantos & 1:
            index = 1
            if cuantos == 3:
                jornada = total
                return {'jornada': jornada, 'comida': comida, 'art83': art83, 'otros': otros, 'error': error}

            meSalgo = False
            while (index < len(me) - 1):
                suplemento = 1

                tipo = me[index][0]
                while (me[index+suplemento][0] != me[index][0]):
                    suplemento+=1
                    if index+suplemento>cuantos-1:
                        meSalgo = True
                        break
                if meSalgo:
                    break
                elapsedtime = me[index + suplemento][1] - me[index][1]
                if tipo == "A":
                    pass
                elif tipo == "C":
                    art83 += elapsedtime
                elif tipo == "B":
                    comida += elapsedtime
                else:
                    otros += elapsedtime
                index += 1+suplemento

            jornada = total - (art83 + comida + otros)
            return {'jornada': jornada, 'comida': comida, 'art83': art83, 'otros': otros, 'error': error}
        else:
            error = total
            return {'jornada': jornada, 'comida': comida, 'art83': art83, 'otros': otros, 'error': error}

    def procesaEntradas(self, me):
        #return bienFormada(me)
        if self.bienFormada(me):
            datos = self.procesaBienFormada(me)
        else:
            datos = self.procesaMalFormada(me)
        return datos

    def print_total(self, td):
        hours, remainder = divmod(td.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        hours += td.days * 24
        return str(hours) + ":" + str(minutes) +":" + str(seconds)

    def procesarFichajes(self):
        index = 1

        self.outSheet.write(0, 0, "Nombre")
        self.outSheet.write(0, 1, "Jornada")
        self.outSheet.write(0, 2, "Comidas")
        self.outSheet.write(0, 3, "Art83")
        self.outSheet.write(0, 4, "Otros")
        self.outSheet.write(0, 5, "Errores")
        self.outSheet.write(0, 6, "JornadaDia")
        self.outSheet.write(0, 7, "Exceso")

        for fichero in self.workbooks:
            NombreTecnico, resultados = self.procesaUnFichero(fichero)

            self.outSheet.write(index, 0, NombreTecnico)
            self.outSheet.write(index, 1, self.print_total(resultados['jornada']))
            self.outSheet.write(index, 2, self.print_total(resultados['comida']))
            self.outSheet.write(index, 3, self.print_total(resultados['art83']))
            self.outSheet.write(index, 4, self.print_total(resultados['otros']))
            self.outSheet.write(index, 5, self.print_total(resultados['error']))
            self.outSheet.write(index, 6, self.print_total(resultados['jornadaDia']))
            self.outSheet.write(index, 7, self.print_total(resultados['exceso']))

            index += 1

        self.outWorkbook.save(self.outFile)

    def procesaUnFichero(self, fichero, printOut = False):
        resultados = {'jornada': ZERO_TIME, 'comida': ZERO_TIME, 'art83': ZERO_TIME, 'otros': ZERO_TIME,
                  'error': ZERO_TIME, 'jornadaDia': ZERO_TIME, "exceso": ZERO_TIME}
        NombreTecnico = ""
        wb = xlrd.open_workbook(fichero)
        ws = wb.sheet_by_index(0)
        for ridx in range(ws.nrows):
            valor = ws.cell(ridx, 0).value
            try:
                horatmp = ws.cell(ridx, 4).value.split(':')
                fecha = datetime.datetime.strptime(valor, "%d/%m/%Y").date()
                jornadaDia = datetime.timedelta(hours=int(horatmp[0]), minutes=int(horatmp[1]))
            except:
                if "datos de:" in valor:
                    indx1 = valor.find("datos de:")+len("datos de:")
                    indx2 = valor.find("Fecha inicio")
                    NombreTecnico = valor[indx1:indx2]
                continue

            entradas = ws.cell(ridx,1).value.splitlines()
            respuesta = self.limpiaEntradas(entradas)
            if respuesta:
                salida = self.procesaEntradas(respuesta)
            else:
                continue

            if salida:
                if printOut:
                    print fecha.strftime('%m/%d/%Y')+": ", salida['jornada'], salida['comida'], salida['art83'], salida['otros'], salida['error']
                resultados['jornada'] += salida['jornada']
                resultados['comida'] += salida['comida']
                resultados['art83'] += salida['art83']
                resultados['otros'] += salida['otros']
                resultados['error'] += salida['error']
                resultados['jornadaDia'] += jornadaDia

        resultados["exceso"] = resultados['jornada']-resultados['jornadaDia']
        return NombreTecnico, resultados

def main():
    tareas = fichajesPost("salida\salida.xls", "entrada")
    tareas.printWorkbooks()
    tareas.procesarFichajes()
    #tareas.procesaUnFichero('JesusArroyo.xls', True)

if __name__ == "__main__":
    main()