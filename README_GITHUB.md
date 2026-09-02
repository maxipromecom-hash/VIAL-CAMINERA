# Verificador Vial Caminera

Aplicación local para Windows que compara las filas de una planilla Excel con la consulta de antecedentes del sistema Vial Caminera.

## Generar el .EXE en GitHub

1. Crear un repositorio nuevo en GitHub (por ejemplo `verificador-vial-caminera`).
2. Subir **todo el contenido de esta carpeta**, incluida la carpeta `.github`.
3. Abrir la pestaña **Actions** del repositorio.
4. Elegir **Compilar Verificador Vial Caminera**.
5. Pulsar **Run workflow** y luego **Run workflow** nuevamente.
6. Esperar a que la ejecución termine con tilde verde.
7. Abrir esa ejecución y, al final de la página, descargar el artifact **Verificador_Vial_Caminera-Windows**.
8. Descomprimirlo. Dentro estará `Verificador_Vial_Caminera.exe`.

## Uso

- El programa NO almacena usuario ni contraseña.
- Primero seleccione el Excel.
- Pulse `ABRIR SISTEMA / INICIAR SESIÓN`.
- Inicie sesión manualmente en Chrome y abra `CONSULTA DE ANTECEDENTES`.
- Vuelva al programa y pulse `INICIAR VERIFICACIÓN`.
- Para la primera prueba, deje `Máx. filas = 1`.

El programa genera un respaldo del Excel y una copia terminada en `_VERIFICADO.xlsx`.
