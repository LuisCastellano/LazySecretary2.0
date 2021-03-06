// Envia un mensaje al destinatario que se le pasa por parametros
function sendEmail(destinatario, usuario, pass, nombre, dominio){
  // Valores de las cadenas de texto
  var valores= [
    "Credenciales de acceso",
    "Hola",
    "Se ha creado una cuenta corporativa para ti en el entorno de Google Apps. En este entorno podrás disfrutar nuestro correo corporativo y muchos más beneficios.. Tus credenciales de acceso son:",
    "Nombre de usuario: ",
    "Clave de acceso: ",
    "Para acceder tienes que hacer login en Google con tu usuario y la contraseña. La primera vez que entres se te solicitará que cambies la contraseña y que establezcas una de tu elección.",
    "Para acceder debes dirigirte a: ",
    "Con tu nueva cuenta corporativa podrás realizar muchísimas acciones y además tendrás fantásticas ventajas: ",
    "Espacio ilimitado en la nube: podrás sincronizar con tus dispositivos archivos sin límite de tamaño ni problemas de espacio (como Dropbox, pero de forma ilimitada).",
    "Sin publicidad: no se muestra publicidad de Google ni en las búsquedas ni en correo.",
    "Sin Spam ni virus: Google filtra todo lo que entra en tu bandeja de entrada.",
    "En el entorno de Google Apps encontrarás las siguientes aplicaciones integradas y que podrás disfrutar desde tu cuenta corporativa:",
    "Comunicación",
    "Correo electrónico corporativo y contactos integrados en Gmail.",
    "Conéctate con las personas que quieras mediante llamadas de voz, chat de texto o vídeo de alta definición. Puedes ahorrar tiempo y dinero en viajes, sin renunciar a ninguna de las ventajas del contacto cara a cara.",
    "Dedica menos tiempo a la planificación y más al trabajo con los calendarios, que se pueden compartir y se integran perfectamente con Gmail, Drive, Contactos, Sites y Hangouts, para que puedas saber en todo momento cuál es el próximo evento.",
    "Red social en el entorno corporativo. Podrás compartir enlaces, videos, comentarios y darte de alta en grupos afines.",
    "Almacenamiento",
    "Mantén todo tu trabajo en un lugar seguro con el almacenamiento de archivos online. Accede a tu trabajo cuando lo necesites, desde el portátil, el tablet o el teléfono móvil.",
    "Colaboración",
    "Crea y edita documentos de texto directamente en tu navegador sin necesidad de software específico. Pueden trabajar varias personas al mismo tiempo en un archivo: todos los cambios se guardan automáticamente.",
    "Crea hojas de cálculo directamente en tu navegador sin necesidad de software específico. Puedes utilizarlas para todo tipo de contenido, desde sencillas listas de tareas hasta análisis de datos con gráficos, filtros y tablas dinámicas.",
    "Crea formularios personalizados para encuestas y cuestionarios sin ningún cargo adicional. Recopila toda la información en una hoja de cálculo y analiza los datos directamente en Hojas de cálculo de Google.",
    "Crea y edita elegantes presentaciones en tu navegador sin necesidad de software específico. Pueden trabajar varias personas al mismo tiempo; de esta forma, todos tienen siempre la versión más reciente.",
    "Crea un sitio de proyectos para tu equipo con nuestra aplicación para el diseño de sitios web. Y todo sin escribir ni una sola línea de código.",
    "Y muchas herramientas más..."
  ];
  
  
    for (var i=0; i<valores.length; i++){
      valores[i]= translate(valores[i]);
    }
  
  
  MailApp.sendEmail({
     to: destinatario,
     subject: valores[0],
    htmlBody: 
    '<div id="cuerpo" style="text-align: justify; font-family: Arial, Helvetica, sans-serif; color:#2E2E2E">'+valores[1]+', '+(nombre.charAt(0).toUpperCase() + nombre.slice(1))+'!<br><br>'
      +valores[2]+'<br><br>'
      +valores[3]+"<b> "+usuario+"</b><br>"
      +valores[4]+"<b> "+pass+"</b> <br><br>"
      +valores[5]+'<br><br>'
      +valores[6]+' <a href="https://accounts.google.com">https://accounts.google.com</a><br><br>'

      +valores[7]
      +'<ul><li>'+valores[8]+'</li>'
      +'<li>'+valores[9]+'</li>'
      +'<li>'+valores[10]+'</li></ul>'
      +valores[11]+'<br><br><br>'

      +'<div style="font-weight:bold; font-size:20pt; color: #878787;">'+valores[12]+'</div>'

      +'<!-- Gmail --><img src="http://i.imgur.com/LpS8rQ9.png" /><br /><a style="font-size:14pt;" href="http://mail.google.com">Gmail</a>'
      +'<div style="color:#656565;">'+valores[13]+'.</div><br><br>'

      +'<!-- Hangouts --><img src="http://i.imgur.com/qnaPCvu.png" /><br /><a style="font-size:14pt;" href="http://www.google.es/hangouts/">Hangouts</a>'
      +'<div style="color:#656565;">'+valores[14]+'</div><br><br>'

      +'<!-- Calendar --><img src="http://i.imgur.com/PGRnI53.png" /><br /><a style="font-size:14pt;" href="http://calendar.google.com">Calendar</a>'
      +'<div style="color:#656565;">'+valores[15]+'</div><br><br>'

      +'<!-- Google+ --><img src="http://i.imgur.com/wo8joYb.png" /><br /><a style="font-size:14pt;" href="http://plus.google.com">Google+</a>'
      +'<div style="color:#656565;">'+valores[16]+'</div><br><br>'

      +'<div style="font-weight:bold; font-size:20pt; color: #878787;">'+valores[17]+'</div>'

      +'<!-- Drive --><img src="http://i.imgur.com/o9cPX2V.png" /><br /><a style="font-size:14pt;" href="http://drive.google.com">Drive</a>'
      +'<div style="color:#656565;">'+valores[18]+'</div><br><br>'

      +'<div style="font-weight:bold; font-size:20pt; color: #878787;">'+valores[19]+'</div>'

      +'<!-- Docs --><img src="http://i.imgur.com/lM48l47.png" /><br /><a style="font-size:14pt;" href="http://docs.google.com">Docs</a>'
      +'<div style="color:#656565;">'+valores[20]+'</div><br><br>'

      +'<!-- Sheets --><img src="http://i.imgur.com/uGM3hpI.png" /><br /><a style="font-size:14pt;" href="http://sheets.google.com">Sheets</a>'
      +'<div style="color:#656565;">'+valores[21]+'</div><br><br>'

      +'<!-- Forms --><img src="http://i.imgur.com/D2y8TY1.png" /><br /><a style="font-size:14pt;" href="http://forms.google.com">Forms</a>'
      +'<div style="color:#656565;">'+valores[22]+'</div><br><br>'

      +'<!-- Slides --><img src="http://i.imgur.com/OM6qd8F.png" /><br /><a style="font-size:14pt;" href="http://slides.google.com">Slides</a>'
      +'<div style="color:#656565;">'+valores[23]+'</div><br><br>'

      +'<!-- Sites --><img src="http://i.imgur.com/8j9Nkl4.png" /><br /><a style="font-size:14pt;" href="http://sites.google.com">Sites</a>'
      +'<div style="color:#656565;">'+valores[24]+'</div><br><br>'

    +'<a href="http://www.google.es/about/products/">'+valores[25]+'</a>'
   });
}

