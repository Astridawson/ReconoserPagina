//1.VERIFICAR CONEXION CON HTML
console.log("HOLA SOY INGRESO");


//2. CREAR MI FUNCION PARA REGISTRO

const iniciarSesion = async (event) => {
    event.preventDefault();

    //2.1 OBTENER LOS VALORES INGRESADOS POR EL USUARIO


    const correoElectronico = document.getElementById("correoElectronico").value;

    const contrasena = document.getElementById("contrasena").value;

    console.log(correoElectronico, contrasena);

    try {

        //2.2 REALIZAR LA PETICION GET A NUESTRA BASE DE DATOS

        const respuesta = await fetch("http://localhost:9000/api/obtenerUsuario");

        const usuarios = await respuesta.json();

        console.log(usuarios);

        //3.  VERIFICAR SI SE ENCONTRO UN USUARIO CON EL CORREO Y LA CONTRASEÑA DADOS

        const esUsuarioRegistrado = usuarios.find(usuario => usuario.correoElectronico == correoElectronico && usuario.contrasena == contrasena);

        if (esUsuarioRegistrado) {

            //VERIFICAMOS SI ES ADMIN

            const correoAdmin = "admin@gmail.com";

            if (esUsuarioRegistrado.correoElectronico == correoAdmin) {
                alert("Hola administrador!");
                window.location.href = "./admin.html"
            }else{

                alert("Ingreso exitoso");
                window.location.href = "./index.html"
            }

        } else {
            alert("Correo o contraseña incorrectos. Usuario no encontrado! Vuelve a intentar o registrate");
        }


    } catch (error) {
        console.error("Error al verificar inicio de sesión:", error);
    }


}