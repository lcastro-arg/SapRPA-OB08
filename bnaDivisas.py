from bs4 import BeautifulSoup
import requests
from typing import Tuple, Optional, List
from pydantic import BaseModel

url = 'https://www.bna.com.ar/Personas'


class Divisa(BaseModel):
    moneda : str
    comprador : Optional[float]
    vendedor : Optional[float]
    promedio : Optional[float]

    def __str__(self) -> str:
        return '\n' + self.moneda.strip('*') + '\n'  "\tCompra: {:.3f}\n\tVenta: {:.3f}\n\tPromedio: {:.3f}".format(self.comprador, self.vendedor, self.promedio)

def getRequest(url : str) -> Tuple[int, bytes]:
    '''Obtener contenido de la url\n
        Devuelve http response status y contenido'''
    result = requests.get(url)
    return result.status_code, result.content


def getDivisas() -> List[Divisa] | None:
    '''Devuelve lista con precio de compra-venta por divisa'''
    status, content = getRequest(url)
    if status == 200:
        # Crear objeto html para parsear el contenido
        htmlDocument = BeautifulSoup(content, 'html.parser')

        # Buscar por css selector : class name
        tables = htmlDocument.select(".table.table.cotizacion")
        if len(tables) >= 2:
            # Buscar todas las filas de la tabla
            tdata = tables[1].find_all('td')
            
            divisas = list()
            titleCounter = 0   # <- Para rastrear titulo
            cycleCounter = 0   # <- Para rastrear compra/venta
            for row in tdata:
                try:    
                    # Si la clase es tit, corresponde a un título
                    className = row['class']
                except KeyError:
                    # El ciclo es moneda -> compra -> venta
                    cycleCounter += 1
                    if cycleCounter == 1:
                        divisas[titleCounter - 1].comprador = float(row.string)   # Si cycleCounter es 1, es compra
                    else:
                        divisas[titleCounter - 1].vendedor = float(row.string)    # Si cycleCounter es 2, es venta
                        divisas[titleCounter - 1].promedio = (divisas[titleCounter - 1].vendedor + divisas[titleCounter - 1].comprador) / 2
                        cycleCounter = 0
                else:
                    if className[0] == 'tit':
                        name = row.string
                        divisas.append(Divisa(moneda= name))
                        titleCounter += 1 

            return divisas
    
    return
                        
if __name__ == '__main__':
    dv = getDivisas()
    for d in dv:
        print(d)