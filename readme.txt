1)Los emails que fueron detectados como 'potencialmente expensa' se les agrega un cartel que dice 'ExpensaProcesada', si llega a haber algun email
que no tiene cartel y es una expensa, avisar y reenviar de manera manual, para ver por que no detecto ese email como expensa.
2) Tenemos 2 alternativas, podemos centralizar todas las expensas a el email artusoexpensas2@gmail.com. Lo cual probablemente genere que haya muchos emails en esa casilla (tal vez sea manejable), y luego el control fino sobre emails que no fueron detectados correctamente como pago de expensas, lo pueden realizar sobre ese email, para incluirlo en el .xlsx que se genera.
La otra alternativa es trabajar con otro mail mas auxiliar a donde se reenvían todos los emails que se clasifican como expensa y si no se detectan, un humano los reenvía manualmente a ese email. Posteriormente el excel se genera automáticamente.



Estrategia:
Antes de comenzar la implementacion que estas haciendo, te queria sugerir un workflow, luego genera nuevamente el plan:entonces, lo primero seria identificar si hay o no un archivo adjunto o imagen pegada en el
  cuerpo del mail. Luego, esta informacion que es extraida por el servicio OCR, primero se debe analizar si hay uno o varios archivos adjuntos o imagenes pegadas en el body del email, si la cantidad es >=2, se
  etiqueta para revision, sino, se debe analizar lo extraido en busca de detectar si es o no un comprobante de pago, utilizando un criterio de que, si es un comprobante de pago, va a tener ciertos datos seguro
  (cuenta origen, cuenta destino, monto). Si es un comprobante y en la cuenta destino hacemos fuzzy matching, ya tenemos por seguro que es un pago de expensa y hay que extraer y volcar la informacion a la
  planilla.