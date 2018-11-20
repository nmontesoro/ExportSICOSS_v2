# ExportSICOSS_v2
Código en VFP9 para exportación de datos a formato txt

## Modo de uso
Crear un archivo `fieldinfo.txt` con la siguiente estructura:
```
tipo_datos|nombre_campo|[padding]longitud[.decimales]|fórmula
```
Los tipos de datos posibles son:
- s: _string_,
- m: dinero, y
- n: número

El _padding_ se indica anteponiendo un '0' a la longitud (completa el valor del campo con ceros hasta llegar a la longitud requerida).

La fórmula puede involucrar campos y/o una función de VFP 9.

Ejemplos:

```
s|conyuge|1|ICASE(Cónyuge, 'T', 'F')
n|cant_hijos|02.0|Cantidad_de_Hijos
s|seguro_vida|1|ICASE(MarcaSeguroVida, 'T', 'F')
m|detraccion|10.2|ICASE(Código_de_Modalidad_de_Contratacion == 99, 0, 2400 * jornada)
```

### Inicio del programa
Se invoca al programa de la siguiente manera:
```
export.exe mes [T-F]
```
Donde T o F indican si tomar la jornada laboral desde un archivo _Recibo.xls_

**Se incluye un .bat como ejemplo de cómo copiar las bases de datos de SICOSS y ejecutar el programa**
