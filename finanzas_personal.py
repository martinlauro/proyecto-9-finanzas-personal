import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint

# ================================================
# 1. DATOS FINANCIEROS ANUALES
# ================================================

meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
         'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

# INGRESOS
sueldo_neto    = [350000]*12
horas_extra    = [20000, 0, 15000, 0, 25000, 0, 30000, 0, 15000, 0, 20000, 0]
aguinaldo      = [0, 0, 0, 0, 0, 175000, 0, 0, 0, 0, 0, 175000]

# GASTOS FIJOS
alquiler       = [120000]*12
luz            = [8000, 9000, 7500, 8500, 9000, 11000, 12000, 11500, 9000, 8000, 8500, 10000]
gas            = [5000, 6000, 5500, 7000, 9000, 12000, 14000, 13000, 10000, 7000, 5500, 5000]
internet       = [4500]*12
seguro         = [8000]*12

# GASTOS VARIABLES
supermercado   = [45000, 48000, 46000, 50000, 47000, 52000, 55000, 53000, 48000, 46000, 50000, 65000]
transporte     = [15000, 15000, 15000, 16000, 15000, 15000, 16000, 15000, 15000, 15000, 15000, 16000]
salud          = [5000, 5000, 8000, 5000, 5000, 5000, 5000, 8000, 5000, 5000, 5000, 5000]
ropa           = [0, 15000, 0, 20000, 0, 25000, 0, 20000, 0, 15000, 30000, 40000]
entretenimiento= [10000, 10000, 12000, 10000, 15000, 20000, 25000, 15000, 10000, 10000, 15000, 30000]

# CREAR DATAFRAME
df = pd.DataFrame({
    'mes':            meses,
    'sueldo_neto':    sueldo_neto,
    'horas_extra':    horas_extra,
    'aguinaldo':      aguinaldo,
    'alquiler':       alquiler,
    'luz':            luz,
    'gas':            gas,
    'internet':       internet,
    'seguro':         seguro,
    'supermercado':   supermercado,
    'transporte':     transporte,
    'salud':          salud,
    'ropa':           ropa,
    'entretenimiento':entretenimiento
})

# CALCULOS
df['total_ingresos']      = df['sueldo_neto'] + df['horas_extra'] + df['aguinaldo']
df['gastos_fijos']        = df['alquiler'] + df['luz'] + df['gas'] + df['internet'] + df['seguro']
df['gastos_variables']    = df['supermercado'] + df['transporte'] + df['salud'] + df['ropa'] + df['entretenimiento']
df['total_gastos']        = df['gastos_fijos'] + df['gastos_variables']
df['ahorro']              = df['total_ingresos'] - df['total_gastos']
df['pct_ahorro']          = (df['ahorro'] / df['total_ingresos'] * 100).round(1)
df['pct_gastos_fijos']    = (df['gastos_fijos'] / df['total_gastos'] * 100).round(1)
df['pct_gastos_variables']= (df['gastos_variables'] / df['total_gastos'] * 100).round(1)

# ================================================
# 2. RESUMEN EN CONSOLA
# ================================================

print("=== RESUMEN FINANCIERO ANUAL ===")
print(f"Total ingresos:  ${df['total_ingresos'].sum():>12,.0f}")
print(f"Total gastos:    ${df['total_gastos'].sum():>12,.0f}")
print(f"Ahorro total:    ${df['ahorro'].sum():>12,.0f}")
print(f"% Ahorro anual:  {df['ahorro'].sum()/df['total_ingresos'].sum()*100:.1f}%")

print("\n=== MESES CON DÉFICIT ===")
deficit = df[df['ahorro'] < 0]
if len(deficit) > 0:
    print(deficit[['mes','total_ingresos','total_gastos','ahorro']].to_string(index=False))
else:
    print("Ninguno — excelente gestión financiera!")

print("\n=== MEJOR Y PEOR MES ===")
mejor = df.loc[df['ahorro'].idxmax()]
peor  = df.loc[df['ahorro'].idxmin()]
print(f"Mejor mes:  {mejor['mes']} — Ahorro: ${mejor['ahorro']:,.0f}")
print(f"Peor mes:   {peor['mes']}  — Ahorro: ${peor['ahorro']:,.0f}")

# ================================================
# 3. GRÁFICOS
# ================================================

fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Dashboard Financiero Personal 2026', fontsize=16, fontweight='bold')

# Gráfico 1: Ingresos vs Gastos vs Ahorro
x = range(len(meses))
axes[0,0].bar(x, df['total_ingresos'], label='Ingresos', color='steelblue', alpha=0.8)
axes[0,0].bar(x, df['total_gastos'],   label='Gastos',   color='coral',     alpha=0.8)
axes[0,0].plot(x, df['ahorro'], color='green', marker='o', linewidth=2, label='Ahorro')
axes[0,0].set_title('Ingresos vs Gastos vs Ahorro')
axes[0,0].set_xticks(x)
axes[0,0].set_xticklabels([m[:3] for m in meses], rotation=45)
axes[0,0].legend()
axes[0,0].yaxis.set_major_formatter(plt.FuncFormatter(lambda v,p: f'${v/1000:.0f}k'))

# Gráfico 2: % Ahorro por mes
colores = ['green' if v >= 0 else 'red' for v in df['pct_ahorro']]
axes[0,1].bar(meses, df['pct_ahorro'], color=colores, alpha=0.8)
axes[0,1].axhline(y=0, color='black', linewidth=0.8)
axes[0,1].set_title('% Ahorro por Mes')
axes[0,1].set_xticklabels([m[:3] for m in meses], rotation=45)
axes[0,1].set_ylabel('%')

# Gráfico 3: Composición de gastos anuales
categorias_gasto = ['Alquiler','Luz','Gas','Internet','Seguro',
                    'Supermercado','Transporte','Salud','Ropa','Entretenimiento']
totales_gasto = [df[c.lower()].sum() for c in categorias_gasto]
colores_pie = ['#FF6B6B','#FF8E53','#FFA07A','#FFD700','#98FB98',
               '#87CEEB','#DDA0DD','#F0E68C','#20B2AA','#778899']
axes[1,0].pie(totales_gasto, labels=categorias_gasto, autopct='%1.1f%%',
              colors=colores_pie, startangle=90)
axes[1,0].set_title('Composición de Gastos Anuales')

# Gráfico 4: Evolución del ahorro acumulado
ahorro_acumulado = df['ahorro'].cumsum()
axes[1,1].fill_between(range(len(meses)), ahorro_acumulado,
                        alpha=0.4, color='green')
axes[1,1].plot(range(len(meses)), ahorro_acumulado,
               color='green', marker='o', linewidth=2)
axes[1,1].set_title('Ahorro Acumulado 2026')
axes[1,1].set_xticks(range(len(meses)))
axes[1,1].set_xticklabels([m[:3] for m in meses], rotation=45)
axes[1,1].yaxis.set_major_formatter(plt.FuncFormatter(lambda v,p: f'${v/1000:.0f}k'))

plt.tight_layout()
plt.savefig('dashboard_finanzas.png', dpi=150, bbox_inches='tight')
plt.show()

# ================================================
# 4. EXPORTAR A EXCEL
# ================================================

fecha_hoy = datetime.now().strftime("%Y%m%d")
archivo   = f"finanzas_personal_{fecha_hoy}.xlsx"

with pd.ExcelWriter(archivo, engine='openpyxl') as writer:
    # Hoja 1: Detalle mensual
    df.to_excel(writer, sheet_name='Detalle Mensual', index=False)

    # Hoja 2: Resumen anual
    resumen = pd.DataFrame({
        'Concepto': ['Total Ingresos','Total Gastos Fijos','Total Gastos Variables',
                     'Total Gastos','Ahorro Total','% Ahorro'],
        'Monto': [
            df['total_ingresos'].sum(),
            df['gastos_fijos'].sum(),
            df['gastos_variables'].sum(),
            df['total_gastos'].sum(),
            df['ahorro'].sum(),
            round(df['ahorro'].sum()/df['total_ingresos'].sum()*100, 1)
        ]
    })
    resumen.to_excel(writer, sheet_name='Resumen Anual', index=False)

    # Hoja 3: Ahorro por mes
    ahorro_mes = df[['mes','total_ingresos','total_gastos','ahorro','pct_ahorro']]
    ahorro_mes.to_excel(writer, sheet_name='Ahorro por Mes', index=False)

print(f"\nArchivo Excel generado: {archivo}")
print("Imagen guardada: dashboard_finanzas.png")