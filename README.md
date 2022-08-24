# dynamicIndex
Indice con Google Apps Script (Macro)

## Este es el paso a paso de lo que va a hacer el código

Nos conectamos con Sheets y con la hoja donde va a ir mi Indice (En mi caso se llama «Inicio»)
-- Traemos todas las pestañas a una variable (En mi caso «misHojas»)
-- Borramos el índice actual (Puede que hayamos eliminado pestañas, entonces es mejor volverlo a hacer de 0)
-- Vamos a recorrer todas las hojas actuales, y para cada una vamos a:
-- Revisar que no sea la hoja de Inicio (Si es la hoja de inicio no hace nada)
-- Creamos la fórmula HIPERVINCULO (No hay otra forma en Google Apps Script para vincular que yo conozca) con el id y el nombre de cada pestaña
-- Insertamos esta formulas en la columna A y en la fila 2 (después en la 3, después en la 4, y así sucesivamente
-- (Opcional) Creamos un vínculo en cada pestaña para volver al inicio
