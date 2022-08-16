from pydantic import BaseModel
from typing import List
import win32com.client
import sys
import subprocess
import time
from datetime import datetime
import os

from bnaDivisas import Divisa, getDivisas
from sendMail import send_email

fechaHoy = datetime.now().strftime('%d.%m.%Y')

class Credentials(BaseModel):
    username : str
    password : str
    server : str

class Tcot:
    '''Tipos de Cotización'''
    comprador = 'G'
    vendedor = 'B'
    estandar = 'M'
    costos = 'P'


class SapGui():
        
    '''Instancear interfaz gráfica de SAP y crear una sesión con el servidor\n
     El parámetro Server responde al nombre del servidor'''

    def __init__(self, server : str, user : str, password : str) -> None:

        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        self.user = user
        self.password = password

        # Crear instancia SAP GUI
        subprocess.Popen(self.path)
        time.sleep(2)

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return

        # Activar Scripting engine
        application = self.SapGuiAuto.GetScriptingEngine

        # Abrir conexión con servidor
        self.connection = application.OpenConnection(server, True)
        time.sleep(3)

        # Crear sesión
        self.session = self.connection.Children(0)
        self.sapLogin()

    def sapLogin(self) -> None:
        '''Log in en el servidor'''
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.user
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
            self.session.findById("wnd[0]").sendVkey(0)
        except:
            print(sys.exc_info()[0])

    def tipoDeCambio(self, divisas : List[Divisa]) -> None:
        '''Ingresar a OB08 y añadir entradas con el tipo de cambio'''
        
        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "ob08"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
        self.cargarTabla(divisas)

    def cargarTabla(self, divisas : List[Divisa]) -> None:
            '''Cargar tabla con los tipos de cambio'''
            div = [['ARS', 'USD'], ['ARS', 'EUR'], ['USD', 'EUR']]
            dolar =  divisas[0]
            euro = divisas[2]
            tipoCambio = [dolar.comprador, dolar.vendedor, dolar.promedio, dolar.comprador, euro.comprador, euro.vendedor, euro.promedio,
                            euro.comprador / dolar.comprador, euro.vendedor / dolar.vendedor, euro.promedio / dolar.promedio]
            tcotsList = [Tcot.vendedor, Tcot.comprador, Tcot.estandar, Tcot.costos]
            tcotIndex = 0
            tcotCycle = 1
            divCycle = 0
            tcCycle = 0
            

            for i in range(20):
                # T cot
                if i in [6, 7]:
                    tcotIndex = 3
                    
                self.session.findById(f"wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/ctxtV_TCURR-KURST[0,{i}]").text = tcotsList[tcotIndex]  
                tcotCycle += 1
                if tcotCycle == 3:
                    tcotIndex += 1
                    tcotCycle = 1

                if i == 7 or i == 13:
                    tcotIndex = 0     

                # Fecha
                self.session.findById(f"wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/ctxtV_TCURR-GDATU[1,{i}]").text = fechaHoy  

                # De moneda A moneda
                if i == 8 or i == 14:
                    divCycle += 1

                monCycle = 0 if i % 2 == 0 else 1   # Ciclo impar | par
                monCycle_col10 = 1 if i % 2 == 0 else 0 # Ciclo par | impar
                self.session.findById(f"wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/ctxtV_TCURR-FCURR[5,{i}]").text = div[divCycle][monCycle]
                self.session.findById(f"wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/ctxtV_TCURR-TCURR[10,{i}]").text = div[divCycle][monCycle_col10]

                # Divisas Col 2 y Col 7
                par = i % 2 != 0
                if par:
                    tc = ""
                    tc_col7 = str(tipoCambio[tcCycle])
                else:
                    tc = str(tipoCambio[tcCycle])
                    tc_col7 = ""

                   
                self.session.findById(f"wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSM[2,{i}]").text = ("%.3f" % float(tc) if tc != "" else tc).replace('.', ',')
                self.session.findById(f"wnd[0]/usr/tblSAPL0SAPTCTRL_V_TCURR/txtRFCU9-KURSP[7,{i}]").text = ("%.3f" % float(tc_col7) if tc_col7 != "" else tc_col7).replace('.', ',')

                if par and i != 0:
                    tcCycle += 1

            # Guardar
            #self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
        
    def closeSap(self):
        self.connection.CloseSession('ses[0]')    


if __name__ == '__main__':
    userData = Credentials(username='Username', password='Password', server='Server Name')
    
    divisas = getDivisas()

    # Chequear que haya datos para cargar
    if divisas != None:
        sap = SapGui(server=userData.server, user= userData.username, password=userData.password)
        sap.tipoDeCambio(divisas)
        sap.closeSap()
        os.system("taskkill /im saplogon.exe /F")
        
        style = '''
                <style>
                        table, td, th {
                            border : 1px solid black;
                            border-collapse : collapse;
                            }
                </style>
                '''

        send_email(usr='User@company.com', 
                    pss='Password', 
                    to_who= 'to',
                    subject=f'Tipos de cambio | {fechaHoy}',
                    body=f'''
                        <h3>Tipos de cambio del día cargados en SAP con éxito</h3>
                        <p></p>
                        <table>
                            <th>Moneda</th>
                            <th>Compra</th>
                            <th>Venda</th>
                            <th>Promedio</th>
                            <tr>
                                <td>{divisas[0].moneda}</td>
                                <td>{divisas[0].comprador:.3f}</td>
                                <td>{divisas[0].vendedor:.3f}</td>
                                <td>{divisas[0].promedio:.3f}</td>
                            </tr>
                            <tr>
                                <td>{divisas[2].moneda}</td>
                                <td>{divisas[2].comprador:.3f}</td>
                                <td>{divisas[2].vendedor:.3f}</td>
                                <td>{divisas[2].promedio:.3f}</td>
                            </tr>
                        </table>
                        {style}
                                ''')