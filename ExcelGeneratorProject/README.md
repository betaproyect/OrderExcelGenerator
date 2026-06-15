# Script de Procesamiento de Pedidos

Este script procesa un archivo Excel llamado `PEDIDOS.xlsx` y crea un nuevo archivo `PedidosProcesados.xlsx` con los siguientes pasos:

## Requisitos

- .NET 6.0 o superior instalado
- Librería NPOI (se instala automáticamente)

## Instrucciones de ejecución

1. Asegúrate de que el archivo `PEDIDOS.xlsx` está fuera de la carpeta `ExcelGeneratorProject`, al mismo nivel que ella.

2. Abre una terminal PowerShell dentro de la carpeta `ExcelGeneratorProject`.

3. Restaura las dependencias del proyecto:
   ```powershell
   dotnet restore
   ```

4. Compila y ejecuta el script:
   ```powershell
   dotnet run
   ```

5. Compila y construye:
   ```powershell
   dotnet build
   ```

6. Publicar:
   ```powershell
   dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true -p:SelfContained=true -o "..\ExcelGenerator"
   ```

## Funcionalidad del Script

### Página 1: "Coincidencias"
El script compara la columna A de la página DATOS con las columnas A de las páginas ECIGLOGISTICA, VAPERALIA y OIL4VAP. Para cada coincidencia encontrada:

- **Columna A**: Fecha actual sin hora (dd/MM/yyyy)
- **Columna B**: Códigos numéricos (3=ECIGLOGISTICA, 2=VAPERALIA, 7=OIL4VAP)
  - Si hay múltiples coincidencias: separadas por " - "
- **Columna E**: Valor de columna B de la distribuidora (formato numérico sin decimales)
- **Columna F**: Valor de columna B de la página DATOS
- **Columna G**: Abreviatura de distribuidora (ECIG, VAP, OIL)
  - Si hay múltiples coincidencias: separadas por " - "
- **Columna H**: "Leganés"

### Página 2: "NoCoincidencias"
Si hay registros en la página DATOS que no coincidan con ninguna distribuidora:

- **Columna A**: Datos de columna A de DATOS
- **Columna B**: Datos de columna B de DATOS

## Archivo de salida

El script genera `PedidosProcesados.xlsx` con la estructura de la página PLANTILLA y los datos procesados.

## Notas

- La comparación de datos es case-insensitive (mayúsculas/minúsculas no importan)
- Se eliminan espacios en blanco al inicio y final de los valores comparados
- No se generan duplicados de coincidencias
- El formato de columna E se establece como numérico sin decimales
