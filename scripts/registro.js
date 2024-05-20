//1.VERIFICAR CONEXION CON HTML
console.log("HOLA SOY REGISTRO");


//2. CREAR MI FUNCION PARA REGISTRO

const registrarUsuario = async (event) => {
    event.preventDefault();

    //2.1 OBTENER LOS DATOS DEL FORMULARIO creando una variable para cada uno de estos datos

    const nombreCompleto = document.getElementById("nombreCompleto").value;

    const correoElectronico = document.getElementById("correoElectronico").value;

    const contrasena = document.getElementById("contrasena").value;

    console.log(nombreCompleto, correoElectronico, contrasena);

    //2.2 CREAR UN OBJETO CON LOS DATOS DEL USUARIO

    const datosUsuario = {
        nombreCompleto,
        correoElectronico,
        contrasena
    }
    console.log("Datos de usuario:", datosUsuario);

    //2.3 HACER LA PETICION POST PARA CREAR USUARO EN LA BASE DE DATOS

    try {
        //CREACION USUARIO EN LA BASE DE DATOS
        console.log("Datos de usuario:", datosUsuario);
        const respuesta = await fetch(" http://localhost:9000/api/crearUsuario",
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(datosUsuario)
            }

        );
        const nuevoUsuario = await respuesta.json();
        console.log("Usuario creado exitosamente", nuevoUsuario);

        //CONDICIONAL PARA REDIRECCIONAR A ingreso.html

        if (nuevoUsuario) {
            alert("Registro exitoso!");
            window.location.href = "./ingreso.html"
        } else {
            alert("Ups! Error de registro, intentelo nuevamente");
        }

    } catch (error) {
        console.error("Error al registrar usuario", error);
    }

};
