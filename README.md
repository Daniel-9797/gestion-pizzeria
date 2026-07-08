# Gestión de pizzería — TPV, previsión y control de food cost

Aplicación de escritorio (Python + Tkinter) para la gestión diaria de una
pizzería: cierre de caja, arqueo, previsión de ventas, control de food cost e
inventario, y avisos de eventos (partidos de fútbol) que afectan a la demanda.

## Características

- **Cierre de caja diario** con arqueo de efectivo (conteo de billetes y monedas,
  apertura heredada del cierre anterior).
- **Previsión de ventas** por media ponderada del histórico y tipo de día.
- **Control de food cost** teórico (sobre venta prevista) y real (sobre venta real),
  presupuesto de compra y seguimiento semanal.
- **Análisis de costes por producto**: evolución de precios, inflación y gasto.
- **Eventos de fútbol**: descarga partidos relevantes (LaLiga, Champions, Mundial,
  Eurocopa) que pueden alterar la demanda, vía la API de football-data.org.
- **Sincronización opcional** con una base en la nube (Neon/PostgreSQL).
- **Correo diario** con el resumen del cierre (opcional, vía Gmail SMTP).

## Requisitos

- Python 3.10+
- PostgreSQL (base local)
- Dependencias: ver `requirements.txt`

## Instalación

```bash
python -m venv venv
venv\Scripts\activate        # Windows
pip install -r requirements.txt
cp .env.example .env         # y rellena tus valores
python app.py
```

## Configuración

Toda la configuración sensible (contraseñas, tokens, datos de conexión) se lee de
variables de entorno definidas en un archivo `.env` local. Copia `.env.example`
como `.env` y rellena tus valores. **El `.env` real nunca se sube al repositorio.**

## Nota

Proyecto de uso interno. Los datos, credenciales y configuración de conexión no se
incluyen en el repositorio. Este código se comparte con fines demostrativos.
