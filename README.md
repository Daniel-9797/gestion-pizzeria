# Gestión de caja de una pizzería

App de escritorio que uso para el día a día de una pizzería: cerrar la caja,
cuadrar el arqueo, llevar el histórico de ventas y hacerme una idea de lo que se
va a vender cada día. La hice a medida porque las hojas de Excel que usaba se me
quedaban cortas.

Está hecha en Python con Tkinter (interfaz) y PostgreSQL (datos).

## Qué hace

- **Cierre del día**: registro de la venta, desglose por formas de pago y método
  de cobro y dinero al banco.
- **Arqueo**: conteo de billetes y monedas. La apertura de cada día se hereda del
  cierre del anterior, y compara el efectivo contado con lo que debería haber.
- **Previsión**: estima la venta de cada día a partir del histórico, dándole más
  peso al año pasado y ajustando según el tipo de día (laborable, findes, etc.).
- **Seguimiento**: compara la venta prevista con la real a lo largo del mes.
- **Histórico y ranking**: consulta de ventas por día, mes y comparativa entre meses.
- **Gráficos y KPIs**: evolución de ventas y algunos indicadores del negocio.
- **Calendario semanal**: vista de la semana con la previsión de cada día.
- **Eventos**: descarga los partidos de fútbol importantes (LaLiga, Champions,
  Mundial, Eurocopa) desde football-data.org, porque un partido gordo entre semana
  cambia bastante la venta de una noche.

Opcionalmente sincroniza los cierres con una base en la nube (para consultarlos
desde otro sitio) y manda un correo con el resumen del día.

## Capturas

**Previsión** — estimación de venta por día según el histórico.
![Previsión](Capturas/prevision.png)

**Seguimiento** — previsto contra real a lo largo del mes.
![Seguimiento](Capturas/seguimiento.png)

**Gráficos**
![Gráficos](Capturas/graficos.png)

**KPIs**
![KPIs](Capturas/kpi.png)

**Calendario semanal**
![Calendario semanal](Capturas/calendario-semanal.png)

**Eventos** — partidos que pueden mover la venta de la noche.
![Eventos](Capturas/eventos.png)

## Cómo ponerla en marcha

Necesita Python 3.10+ y un PostgreSQL. Las dependencias están en `requirements.txt`.

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
python app.py
```

La configuración (base de datos, contraseñas, token de fútbol, correo) va en un
archivo `.env` que no se sube al repo. Tienes una plantilla en `.env.example` para
saber qué hay que rellenar.

## Aviso

Es una herramienta interna que fui montando para mi propio uso, así que da por
hecho mi manera de trabajar y mi base de datos. La subo para enseñar el código.
Sin una base PostgreSQL propia arranca pero sale sin datos. Ni las credenciales ni
los datos del negocio están en el repositorio.
