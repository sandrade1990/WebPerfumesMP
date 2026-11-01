//Variable que mantiene el estado visible del carrito
var carritoVisible = false;
var numeroWhatsAppPedidos = '5491150119136';

function crearFichaProducto(producto){
    var item = document.createElement('div');
    item.className = 'item';
    item.setAttribute('data-sku', producto.sku);

    var titulo = document.createElement('span');
    titulo.className = 'titulo-item';
    titulo.textContent = producto.nombre;

    var imagen = document.createElement('img');
    imagen.className = 'img-item';
    imagen.src = '../img/perfumes/' + encodeURIComponent(producto.imagen);
    imagen.alt = producto.nombre;

    var precio = document.createElement('span');
    precio.className = 'precio-item';
    precio.textContent = '$' + new Intl.NumberFormat('es-AR').format(producto.precio);

    var boton = document.createElement('button');
    boton.className = 'boton-item';
    boton.textContent = 'Agregar al Carrito';

    item.appendChild(titulo);
    item.appendChild(imagen);
    item.appendChild(precio);
    item.appendChild(boton);
    return item;
}

function cargarCatalogo(){
    var contenedor = document.querySelector('[data-excel-catalog]');
    if(!contenedor){
        return Promise.resolve();
    }
    var categoriaPagina = String(contenedor.getAttribute('data-excel-catalog')).trim().toLowerCase();
    if(typeof XLSX === 'undefined'){
        return Promise.reject(new Error('No se pudo cargar el lector de Excel'));
    }

    return fetch('../data/productos.xlsx', { cache: 'no-store' })
        .then(function(respuesta){
            if(!respuesta.ok){
                throw new Error('No se pudo cargar data/productos.xlsx');
            }
            return respuesta.arrayBuffer();
        })
        .then(function(datos){
            var libro = XLSX.read(datos, { type: 'array' });
            var hoja = libro.Sheets[libro.SheetNames[0]];
            var filas = XLSX.utils.sheet_to_json(hoja, { defval: '' });
            var productos = [];

            filas.forEach(function(fila){
                var campos = {};
                Object.keys(fila).forEach(function(columna){
                    campos[columna.trim().toLowerCase()] = fila[columna];
                });

                var sku = String(campos.sku || '').trim();
                var nombre = String(campos.nombre || '').trim();
                var imagen = String(campos.imagen || '').trim();
                var categoria = String(campos.categoria || '').trim().toLowerCase();
                var precio = typeof campos.precio === 'number'
                    ? campos.precio
                    : parseInt(String(campos.precio).replace(/[^0-9]/g, ''), 10);

                if(sku && nombre && imagen && categoria === categoriaPagina && isFinite(precio) && precio > 0){
                    productos.push({ sku: sku, nombre: nombre, imagen: imagen, precio: precio });
                }
            });

            if(productos.length === 0){
                throw new Error('El Excel no contiene productos válidos para ' + categoriaPagina);
            }

            contenedor.replaceChildren();
            productos.forEach(function(producto){
                contenedor.appendChild(crearFichaProducto(producto));
            });
        });
}

//Espermos que todos los elementos de la pàgina cargen para ejecutar el script
if(document.readyState == 'loading'){
    document.addEventListener('DOMContentLoaded', ready)
}else{
    ready();
}

function ready(){
    if(!document.querySelector('.carrito')){
        return;
    }

    cargarCatalogo()
        .catch(function(error){
            console.error(error);
            var contenedor = document.querySelector('[data-excel-catalog]');
            if(contenedor){
                contenedor.textContent = 'No se pudo cargar el catálogo de productos.';
            }
        })
        .then(inicializarCarrito);
}

function inicializarCarrito(){
    // Cargamos el contenido del carrito desde LocalStorage (si existe)
  var carritoGuardado = localStorage.getItem('carrito');
  if (carritoGuardado) {
    var itemsCarrito = document.getElementsByClassName('carrito-items')[0];
    itemsCarrito.innerHTML = carritoGuardado;
    actualizarTotalCarrito();
     // Si hay elementos en el carrito, mostramos el carrito
     hacerVisibleCarrito();
  }

    //Agregremos funcionalidad a los botones eliminar del carrito
    var botonesEliminarItem = document.getElementsByClassName('btn-eliminar');
    for(var i=0;i<botonesEliminarItem.length; i++){
        var button = botonesEliminarItem[i];
        button.addEventListener('click',eliminarItemCarrito);
    }

    //Agrego funcionalidad al boton sumar cantidad
    var botonesSumarCantidad = document.getElementsByClassName('sumar-cantidad');
    for(var i=0;i<botonesSumarCantidad.length; i++){
        var button = botonesSumarCantidad[i];
        button.addEventListener('click',sumarCantidad);
    }

     //Agrego funcionalidad al buton restar cantidad
    var botonesRestarCantidad = document.getElementsByClassName('restar-cantidad');
    for(var i=0;i<botonesRestarCantidad.length; i++){
        var button = botonesRestarCantidad[i];
        button.addEventListener('click',restarCantidad);
    }

    //Agregamos funcionalidad al boton Agregar al carrito
    var botonesAgregarAlCarrito = document.getElementsByClassName('boton-item');
    for(var i=0; i<botonesAgregarAlCarrito.length;i++){
        var button = botonesAgregarAlCarrito[i];
        button.addEventListener('click', agregarAlCarritoClicked);
    }

    //Agregamos funcionalidad al botón comprar
    var botonPedido = document.getElementsByClassName('btn-pagar')[0];
    if(botonPedido){
        botonPedido.addEventListener('click',pagarClicked);
    }
}

function pagarClicked(){
    var carritoItems = document.getElementsByClassName('carrito-items')[0];
    var items = carritoItems.getElementsByClassName('carrito-item');
    var campoNombre = document.getElementById('nombre');
    var campoEmail = document.getElementById('email');
    var nombre = campoNombre ? campoNombre.value.trim() : '';
    var email = campoEmail ? campoEmail.value.trim() : '';
    var numero = numeroWhatsAppPedidos.replace(/\D/g, '');

    if(items.length === 0){
        alert('Agrega al menos un producto al carrito.');
        return;
    }
    if(!nombre){
        alert('Ingresa tu nombre para preparar el pedido.');
        if(campoNombre){
            campoNombre.focus();
        }
        return;
    }
    if(!campoEmail || !campoEmail.checkValidity()){
        alert('Ingresa un correo electrónico válido.');
        if(campoEmail){
            campoEmail.focus();
        }
        return;
    }
    if(numero.length < 10 || numero.length > 15){
        alert('Configura el número real de WhatsApp en css/app.js para recibir pedidos.');
        return;
    }

    var lineasPedido = [];
    for(var i=0; i<items.length; i++){
        var producto = items[i];
        var titulo = producto.getElementsByClassName('carrito-item-titulo')[0].innerText;
        var cantidad = producto.getElementsByClassName('carrito-item-cantidad')[0].value;
        var precio = producto.getElementsByClassName('carrito-item-precio')[0].innerText;
        var subtotal = producto.getElementsByClassName('carrito-item-subtotal')[0].innerText;
        lineasPedido.push(cantidad + ' x ' + titulo + ' | Unitario: ' + precio + ' | Subtotal: ' + subtotal);
    }

    var total = document.getElementsByClassName('carrito-precio-total')[0].innerText;
    var mensaje = [
        'Hola, quiero realizar este pedido:',
        '',
        'Nombre: ' + nombre,
        'Correo: ' + email,
        '',
        'Productos:',
        lineasPedido.join('\n'),
        '',
        'Total: ' + total
    ].join('\n');
    var urlWhatsApp = 'https://wa.me/' + numero + '?text=' + encodeURIComponent(mensaje);
    window.open(urlWhatsApp, '_blank', 'noopener,noreferrer');
}
//Funciòn que controla el boton clickeado de agregar al carrito
function agregarAlCarritoClicked(event){
    var button = event.target;
    var item = button.parentElement;
    var titulo = item.getElementsByClassName('titulo-item')[0].innerText;
    var precio = item.getElementsByClassName('precio-item')[0].innerText;
    var imagenSrc = item.getElementsByClassName('img-item')[0].src;
    console.log(imagenSrc);

    agregarItemAlCarrito(titulo, precio, imagenSrc);
    hacerVisibleCarrito();
}

//Funcion que hace visible el carrito
function hacerVisibleCarrito() {
    var itemsCarrito = document.getElementsByClassName('carrito-items')[0];
    if (itemsCarrito.childElementCount > 0) {
      carritoVisible = true;
      var carrito = document.getElementsByClassName('carrito')[0];
      carrito.style.marginRight = '0';
      carrito.style.opacity = '1';

      var items = document.getElementsByClassName('contenedor-items')[0];
      items.style.width = '60%';
    }
  }

//Funciòn que agrega un item al carrito
function agregarItemAlCarrito(titulo, precio, imagenSrc){
    var item = document.createElement('div');
    item.classList.add = ('item');
    var itemsCarrito = document.getElementsByClassName('carrito-items')[0];

    //controlamos que el item que intenta ingresar no se encuentre en el carrito
    var nombresItemsCarrito = itemsCarrito.getElementsByClassName('carrito-item-titulo');
    for(var i=0;i < nombresItemsCarrito.length;i++){
        if(nombresItemsCarrito[i].innerText==titulo){
            alert("El item ya se encuentra en el carrito");
            return;
        }
    }

    var itemCarritoContenido = `
        <div class="carrito-item">
            <img src="${imagenSrc}" width="80px" alt="">
            <div class="carrito-item-detalles">
                <span class="carrito-item-titulo">${titulo}</span>
                
                <div class="selector-cantidad">
                    <i class="fa-solid fa-minus restar-cantidad"></i>
                    <input type="text" value="1" class="carrito-item-cantidad" disabled>
                    <i class="fa-solid fa-plus sumar-cantidad"></i>
                </div>
                <span class="carrito-item-precio">${precio}</span>
                <div>
                <span class="titulo-item">Subtotal: </span>
                <span class="carrito-item-subtotal">$0,00</span> 
                </div>
            </div>
            <button class="btn-eliminar">
                <i class="fa-solid fa-trash"></i>
            </button>
        </div>
    `
    item.innerHTML = itemCarritoContenido;
    itemsCarrito.append(item);

    //Agregamos la funcionalidad eliminar al nuevo item
     item.getElementsByClassName('btn-eliminar')[0].addEventListener('click', eliminarItemCarrito);

    //Agregmos al funcionalidad restar cantidad del nuevo item
    var botonRestarCantidad = item.getElementsByClassName('restar-cantidad')[0];
    botonRestarCantidad.addEventListener('click',restarCantidad);

    //Agregamos la funcionalidad sumar cantidad del nuevo item
    var botonSumarCantidad = item.getElementsByClassName('sumar-cantidad')[0];
    botonSumarCantidad.addEventListener('click',sumarCantidad);

    calcularSubtotal(item);

    //Actualizamos total
    actualizarTotalCarrito();
}
//Aumento en uno la cantidad del elemento seleccionado
function sumarCantidad(event){
    var buttonClicked = event.target;
    var selector = buttonClicked.parentElement;
    console.log(selector.getElementsByClassName('carrito-item-cantidad')[0].value);
    var cantidadActual = selector.getElementsByClassName('carrito-item-cantidad')[0].value;
    cantidadActual++;
    selector.getElementsByClassName('carrito-item-cantidad')[0].value = cantidadActual;
    var item = selector.closest('.carrito-item');
    calcularSubtotal(item);
    actualizarTotalCarrito();
    guardarCarritoEnLocalStorage();
    hacerVisibleCarrito();
}

function calcularSubtotal(item) {
    var cantidadItem = item.getElementsByClassName('carrito-item-cantidad')[0];
    var cantidad = cantidadItem.value;
    var precioElemento = item.getElementsByClassName('carrito-item-precio')[0];
    var precio = parseFloat(precioElemento.innerText.replace('$', '').replace('.', ''));
    var subtotal = precio * cantidad;
    item.getElementsByClassName('carrito-item-subtotal')[0].innerText = '$' + subtotal.toLocaleString("es") + ",00";
}

//Resto en uno la cantidad del elemento seleccionado
function restarCantidad(event){
    var buttonClicked = event.target;
    var selector = buttonClicked.parentElement;
    console.log(selector.getElementsByClassName('carrito-item-cantidad')[0].value);
    var cantidadActual = selector.getElementsByClassName('carrito-item-cantidad')[0].value;
    cantidadActual--;
    if(cantidadActual>=1){
        selector.getElementsByClassName('carrito-item-cantidad')[0].value = cantidadActual;
        // Actualizar el subtotal del ítem
        var item = selector.closest('.carrito-item');
        calcularSubtotal(item);
        actualizarTotalCarrito();
        hacerVisibleCarrito();
    }
}

//Elimino el item seleccionado del carrito
function eliminarItemCarrito(event){
    var buttonClicked = event.target;
    buttonClicked.parentElement.parentElement.remove();
    //Actualizamos el total del carrito
    actualizarTotalCarrito();
    guardarCarritoEnLocalStorage();

    //la siguiente funciòn controla si hay elementos en el carrito
    //Si no hay elimino el carrito
    ocultarCarrito();
}
//Funciòn que controla si hay elementos en el carrito. Si no hay oculto el carrito.
function ocultarCarrito(){
    var carritoItems = document.getElementsByClassName('carrito-items')[0];
    if(carritoItems.childElementCount==0){
        var carrito = document.getElementsByClassName('carrito')[0];
        carrito.style.marginRight = '-100%';
        carrito.style.opacity = '0';
        carritoVisible = false;
    
        var items =document.getElementsByClassName('contenedor-items')[0];
        items.style.width = '100%';
    }
}
//Actualizamos el total de Carrito
function actualizarTotalCarrito(){
    //seleccionamos el contenedor carrito
    var carritoContenedor = document.getElementsByClassName('carrito')[0];
    var carritoItems = carritoContenedor.getElementsByClassName('carrito-item');
    var total = 0;
    //recorremos cada elemento del carrito para actualizar el total
    for(var i=0; i< carritoItems.length;i++){
        var item = carritoItems[i];
        var precioElemento = item.getElementsByClassName('carrito-item-precio')[0];
        //quitamos el simobolo peso y el punto de milesimos.
        var precio = parseFloat(precioElemento.innerText.replace('$','').replace('.',''));
        var cantidadItem = item.getElementsByClassName('carrito-item-cantidad')[0];
        console.log(precio);
        var cantidad = cantidadItem.value;
        total = total + (precio * cantidad);
    }
    total = Math.round(total * 100)/100;

    document.getElementsByClassName('carrito-precio-total')[0].innerText = '$'+total.toLocaleString("es") + ",00";

    guardarCarritoEnLocalStorage();

}

 // Función que guarda el contenido del carrito en LocalStorage
 function guardarCarritoEnLocalStorage() {
    var itemsCarrito = document.getElementsByClassName('carrito-items')[0];
    var contenidoCarrito = itemsCarrito.innerHTML;
    localStorage.setItem('carrito', contenidoCarrito);
  }

 