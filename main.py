import sys
import datetime
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout

class CalculadoraHoras(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        # Widgets
        self.labelEntrada = QLabel("Hora de entrada (HH:MM):")
        self.inputEntrada = QLineEdit()
        
        self.labelSalida = QLabel("Hora de salida (HH:MM):")
        self.inputSalida = QLineEdit()
        
        self.botonCalcular = QPushButton("Calcular Horas")
        self.botonCalcular.clicked.connect(self.calcular_horas)

        self.botonImportarHojaExcel = QPushButton("Importar Hoja de Excel")
        self.botonImportarHojaExcel.clicked.connect(self.calcular_horas)
        
        self.resultado = QTextEdit()
        self.resultado.setReadOnly(True)
        
        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.labelEntrada)
        layout.addWidget(self.inputEntrada)
        layout.addWidget(self.labelSalida)
        layout.addWidget(self.inputSalida)
        layout.addWidget(self.botonCalcular)
        layout.addWidget(self.resultado)
        layout.addWidget(self.botonImportarHojaExcel)
        
        self.setLayout(layout)
        self.setWindowTitle("Calculadora de Horas Trabajadas")

        self.jornada_ordinaria = datetime.timedelta(hours=8)
        self.rangos = {
                        "diurna": (datetime.datetime.strptime("06:00", "%H:%M"), datetime.datetime.strptime("21:00", "%H:%M")),
                        "nocturna": (datetime.datetime.strptime("21:00", "%H:%M"), datetime.datetime.strptime("06:00", "%H:%M"))
        }
    
    def formatear_horas(self, hora):
        horas= str(hora).split(":")[0]
        minutos = str(hora).split(":")[1]
        return f"{horas} hora{'s' if int(horas) != 1 else ''} {f'y {minutos} minuto{"s" if int(minutos) != 1 else ""}' if int(minutos) else ''}"
    

    def imprimir_salida(self, horas_diurnas, horas_extra_diurnas, horas_ordinarias_nocturnas, horas_extra_nocturnas, horas_nocturnas):
        self.resultado.setText(f"""
    Horas ordinarias diurnas: {self.formatear_horas(horas_diurnas) if horas_diurnas else 'No hay horas ordinarias diurnas'}
    Horas extra diurnas: {self.formatear_horas(horas_extra_diurnas) if horas_extra_diurnas else 'No hay horas extra diurnas'}
    Horas ordinarias nocturnas: {self.formatear_horas(horas_ordinarias_nocturnas) if horas_ordinarias_nocturnas else 'No hay horas ordinarias nocturnas'}
    Horas extra nocturnas: {self.formatear_horas(horas_extra_nocturnas) if horas_extra_nocturnas else 'No hay horas extra nocturnas'}
    Horas con recargo nocturno: {self.formatear_horas(horas_nocturnas) if horas_nocturnas else 'No hay horas con recargo nocturno'}
        """)

    def ordenar_horas(self, entrada, salida, limite, jornada_ordinaria):
        horas_trabajadas = min(salida, limite) - entrada
        horas_ordinarias = min(horas_trabajadas, jornada_ordinaria)
        horas_extra = max(horas_trabajadas - jornada_ordinaria, datetime.timedelta())
        return horas_ordinarias, horas_extra

    def calcular_horas(self):
        try:
            hora_entrada = datetime.datetime.strptime(self.inputEntrada.text(), "%H:%M")
            hora_salida = datetime.datetime.strptime(self.inputSalida.text(), "%H:%M") 
            if hora_salida <= hora_entrada:
                hora_salida +=  datetime.timedelta(days=1)
            horas_laboradas = hora_salida - hora_entrada
            if hora_entrada >= self.rangos["diurna"][0] and hora_salida <= self.rangos["diurna"][1]:
                horas_ordinarias_diurnas, horas_extra_diurnas = self.ordenar_horas(hora_entrada, hora_salida, self.rangos["diurna"][1], self.jornada_ordinaria)
                horas_ordinarias_nocturnas = 0
                horas_extra_nocturnas = 0 
                horas_nocturnas = 0
            elif hora_entrada >= self.rangos["nocturna"][0] and hora_salida <= self.rangos["nocturna"][1]:
                horas_ordinarias_nocturnas, horas_extra_nocturnas = self.ordenar_horas(hora_entrada, hora_salida, self.rangos["nocturna"][1], self.jornada_ordinaria)
                horas_ordinarias_diurnas = 0
                horas_extra_diurnas = 0
                horas_nocturnas = horas_laboradas
            else:
                # Caso 1: Entra en el dia y sale en la noche
                if hora_entrada >= self.rangos["diurna"][0] and hora_entrada < self.rangos["diurna"][1]:
                    horas_ordinarias_diurnas, horas_extra_diurnas = self.ordenar_horas(hora_entrada, self.rangos["diurna"][1], self.rangos["diurna"][1], self.jornada_ordinaria)
                    horas_ordinarias_nocturnas, horas_extra_nocturnas = self.ordenar_horas(self.rangos["nocturna"][0], hora_salida, self.rangos["nocturna"][1] + datetime.timedelta(days=1), self.jornada_ordinaria - horas_ordinarias_diurnas)
                    horas_nocturnas = hora_salida - self.rangos["nocturna"][0]
                # Caso 2: Entra en la noche y sale en el dia
                else:
                    if hora_entrada.day == hora_salida.day:
                        horas_ordinarias_nocturnas, horas_extra_nocturnas = self.ordenar_horas(hora_entrada, self.rangos["nocturna"][1], self.rangos["nocturna"][1], self.jornada_ordinaria)
                        horas_ordinarias_diurnas, horas_extra_diurnas = self.ordenar_horas(self.rangos["diurna"][0], hora_salida, self.rangos["diurna"][1], self.jornada_ordinaria - horas_ordinarias_nocturnas)
                        horas_nocturnas = self.rangos["nocturna"][1] - hora_entrada
                    else:
                        horas_ordinarias_nocturnas, horas_extra_nocturnas = self.ordenar_horas(hora_entrada, self.rangos["nocturna"][1] + datetime.timedelta(days=1), self.rangos["nocturna"][1] + datetime.timedelta(days=1), self.jornada_ordinaria)
                        horas_ordinarias_diurnas, horas_extra_diurnas = self.ordenar_horas(self.rangos["diurna"][0] + datetime.timedelta(days=1), hora_salida, self.rangos["diurna"][1] + datetime.timedelta(days=1), self.jornada_ordinaria - horas_ordinarias_nocturnas)
                        horas_nocturnas = self.rangos["nocturna"][1] + datetime.timedelta(days=1)  - hora_entrada
            self.imprimir_salida(horas_ordinarias_diurnas, horas_extra_diurnas, horas_ordinarias_nocturnas, horas_extra_nocturnas, horas_nocturnas)
        except ValueError:
            self.resultado.setText("Error: Formato incorrecto. Use HH:MM")
            return


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ventana = CalculadoraHoras()
    ventana.show()
    sys.exit(app.exec())
