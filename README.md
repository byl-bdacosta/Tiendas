Generador automático de textos de horario (Excel VBA)
Este proyecto contiene una macro VBA para Excel que procesa horarios de apertura y cierre por días de la semana y genera automáticamente textos de horario en varios idiomas.

La macro no depende de la posición exacta de las columnas, y funciona siempre que existan columnas de Apertura y Cierre dentro de cada bloque de día (Lunes–Viernes, Sábado, Domingo, etc.), sin importar dónde estén ubicadas en la hoja.

Características principales:

Detección automática de:

  -Hoja de origen (Horarios habituales o HORARIO ESPAÑA)

  -Columnas de idioma (Inglés, Español, Gallego, Catalán)

    Bloques de días:

    -Lunes a Viernes

    -Sábado

    -Domingo

    -Domingo 30 (opcional)

  -Columnas de Apertura y Cierre (hasta 2 turnos)

  -Construcción automática de textos de horario:

  -Horarios continuos o con doble turno


Formatos inteligentes según 3 casos:

    Caso 1: L = S = D

    Caso 2: L = S ≠ D

    Caso 0: Formato genérico (Mon–Fri | Sat | Sun)


Textos generados en 4 idiomas:

    EN, ES, GL, CA

    Soporte para Domingo 30 de noviembre(o similar) como horario especial

  -Corrección automática de acentos mal codificados (Ã¡ → á, etc.)



📄 Archivos comprobados

La macro ha sido probada con:

Horario_Tiendas_Iberia_Actualizado_2.xlsx

Horarios y Aperturas Especiales - FW25.xlsx

Pero es totalmente compatible con cualquier formato de fichero, siempre que se respeten los nombres de las cabeceras.

📌 Requisitos de la hoja

    La macro funcionará si la hoja contiene:

    1. Una columna con cabecera:
    COD
    
    2. Bloques de días con cualquiera de estos textos:
    Lunes a Viernes
    Sábado
    Domingo
    Domingo 30 (opcional)
    
    3. Dentro de cada bloque:
    
    Columnas con título (en cualquier fila de cabecera o subcabecera):
    
    Apertura
    Cierre
    
    4. Columnas de idioma:
    Inglés  / Ingles
    Español / Espanol
    Gallego
    Catalán / Catalan


Importante:
La posición de estas columnas NO importa.
La macro las detecta automáticamente por texto, independientemente del orden o estructura de la hoja.

🧠 Lógica de horarios

Cada día se interpreta con la estructura:

    Apertura 1
    Cierre 1
    Apertura 2 (opcional)
    Cierre 2 (opcional)


La macro:

Usa dos turnos si están completos

Si hay huecos, fusiona los valores y genera un horario continuo

Casos reconocidos

    Caso	Condición	Formato generado
    1	L = S = D	"Lun - Dom: 10:00 - 21:00"
    2	L = S ≠ D	"Lun - Sáb: 10:00 - 21:00 | Dom: 11:00 - 20:00"
    0	Otros	"Lun - Vie: ... | Sáb: ... | Dom: ..."

    
🌍 Idiomas soportados

Se generan textos en:

    EN – Inglés
    
    ES – Español
    
    GL – Gallego
    
    CA – Catalán

Y se añade automáticamente el texto para:

    Domingo 30 de noviembre si el bloque existe.

🛠️ Cómo usar la macro

Abre tu archivo Excel.

Pulsa ALT + F11 para abrir el editor de VBA.

Inserta un nuevo módulo.

Copia y pega el contenido completo del archivo .bas proporcionado.

Asegúrate de que la hoja se llama:

Horarios habituales

o HORARIO ESPAÑA

Ejecuta la macro:

Horarios


Los textos generados se escribirán automáticamente en las columnas de idioma.

🔧 Corrección de caracteres mal codificados

Al final del proceso, la hoja completa es revisada para corregir caracteres como:

Ã¡ → á

Ã© → é

Ã± → ñ

Ãœ → Ü

etc.

Esto asegura que los textos finales siempre estén correctamente acentuados.

📬 Soporte

Si deseas mejorar el README, añadir imágenes, o generar una versión en inglés, solo pídelo.
