# App Revisión de Detenciones - v6

Versión actualizada con lógica de rango referencial del reporte AMT.

## Cambio aplicado

Cuando un evento de AMT inicia exactamente en el inicio del rango descargado del reporte, ese horario se considera referencial y no necesariamente el inicio real de la detención.

Cuando un evento de AMT termina exactamente en el término del rango descargado del reporte, ese horario también se considera referencial y no necesariamente el término real de la detención.

Por lo anterior, la app no marca como error diferencias de inicio, término o duración que se expliquen por esos bordes del rango del reporte. El resto de las validaciones se mantiene: respaldo en AMT, continuidad de cortes, solapamientos y diferencias fuera de los bordes referenciales.

## Ejecución

```powershell
streamlit run app.py
```

## Dependencias

```powershell
python -m pip install -r requirements.txt
```
