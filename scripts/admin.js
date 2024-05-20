

//1. VERIFICAR CONEXION CON HTML
console.log("HOLA SOY ADMIN");

//FUNCION PARA OBTENER LOS DATOS DE MI BASE DE DATOS

const mostrarUsuarios = async () => {
    try {
        const respuesta = await fetch("http://localhost:9000/api/obtenerUsuario");
        const usuarios = await respuesta.json();

        console.log(usuarios);
        mostrarTabla(usuarios);

    } catch (error) {
        console.error("Error al obtener los usuarios", error);
    }
}

function mostrarTabla(usuarios) {
    const tabla = document.getElementById("miTabla");
    tabla.innerHTML = "";

    usuarios.forEach(usuario => {
        // POR CADA USUARIO (forEach) ME VA A CREAR UNA FILA EN LA TABLA
        tabla.innerHTML += `
    
                <tr>
                    <td>${usuario.nombreCompleto}</td>
                    <td>${usuario.correoElectronico}</td>
                    <td>
                        <button type="button" class="btn btn-danger" id="${usuario._id}" onclick="eliminarUsuario(event)"><i class="bi bi-trash"></i> Eliminar</button>
                    </td>
                </tr>
    
    `
    });
}

mostrarUsuarios();

//FUNCION PARA ELIMINAR USUARIOS POR ID

function eliminarUsuario(event) {

    console.log("Eliminar");

    const idUsuarioEliminar = event.target.id

    console.log(idUsuarioEliminar);

    fetch(`http://localhost:9000/api/eliminarUsuario/${idUsuarioEliminar}` , {method: "DELETE"} ) .then(

    response => {

        if(!response.ok){

            console.error("Error! no se pudo eliminar usuario");
        } else{
        alert("Usuario eliminado correctamente")
       mostrarUsuarios();
        }
    }).catch(error => {

        console.log("Error al eliminar usuario",  error);
    })
}