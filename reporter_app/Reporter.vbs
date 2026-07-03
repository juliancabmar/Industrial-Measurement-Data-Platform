' Crear el objeto del sistema

Set objShell = CreateObject("WScript.Shell")



' Ejecutar el comando (Ejemplo: abrir el bloc de notas)

objShell.Run "EXCEL.EXE /x /e Reporter.xlsm", 1, False