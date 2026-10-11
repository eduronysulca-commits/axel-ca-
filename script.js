let datos = [];
let archivosGuardados = [];


// =====================================================
// CONFIGURACIÓN
// =====================================================

const DB_NAME = "DashboardVentasDB";
const DB_VERSION = 1;
const STORE_NAME = "archivos";


// =====================================================
// PRODUCTOS EXCLUIDOS
// =====================================================

const PRODUCTOS_EXCLUIDOS = [

    "alta",

    "portabilidad",

    "prepago",

    "adaptador hora 20w usb/c-c",

    "recuperacion de numero"

];


const PRODUCTOS_QUE_EMPIEZAN_EXCLUIDOS = [

    "cargador"

];


// =====================================================
// NORMALIZAR TEXTO
// =====================================================

function normalizarTexto(texto) {

    return String(texto)

        .trim()

        .toLowerCase()

        .normalize("NFD")

        .replace(
            /[\u0300-\u036f]/g,
            ""
        );

}


// =====================================================
// EXCLUIR PRODUCTOS
// =====================================================

function productoDebeExcluirse(producto) {

    if (!producto) {

        return true;

    }


    const texto =
        normalizarTexto(producto);


    if (
        PRODUCTOS_EXCLUIDOS
            .includes(texto)
    ) {

        return true;

    }


    for (
        const palabra
        of PRODUCTOS_EXCLUIDOS
    ) {

        if (
            texto.startsWith(
                palabra + " "
            )
        ) {

            return true;

        }

    }


    for (
        const palabra
        of PRODUCTOS_QUE_EMPIEZAN_EXCLUIDOS
    ) {

        if (
            texto.startsWith(palabra)
        ) {

            return true;

        }

    }


    return false;

}



 // =====================================================
// BASE DE DATOS EN GOOGLE SHEETS
// MEDIANTE CLOUDFLARE WORKER
// =====================================================

const API_URL =
    "https://dashboard-ventas-api.movistarzamacola.workers.dev";

// =====================================================
// CONSULTAR REGISTROS
// =====================================================

async function consultarRegistros() {
    const respuesta = await fetch(API_URL, {
        method: "GET",
        cache: "no-store"
    });

    if (!respuesta.ok) {
        throw new Error("No se pudo consultar Google Sheets.");
    }

    const resultado = await respuesta.json();

    if (!resultado.ok) {
        throw new Error(
            resultado.error || "Error al consultar registros."
        );
    }

    return resultado.registros || [];
}

// =====================================================
// GUARDAR ARCHIVO
// =====================================================

async function guardarArchivo(archivo) {
    const id =
        globalThis.crypto?.randomUUID?.() ||
        `${Date.now()}-${Math.random().toString(36).slice(2)}`;

    const registros = archivo.ventas.map(venta => ({
        ID: id,
        NombreArchivo: archivo.nombre,
        Fecha: archivo.fecha,
        Codigo: venta.codigo,
        Producto: venta.producto,
        Cantidad: venta.cantidad,
        UM: venta.um,
        Base: venta.base,
        Total: venta.total
    }));

    const respuesta = await fetch(API_URL, {
        method: "POST",
        headers: {
            "Content-Type": "text/plain;charset=UTF-8"
        },
        body: JSON.stringify({
            accion: "guardar",
            registros: registros
        })
    });

    if (!respuesta.ok) {
        throw new Error("No se pudieron guardar las ventas.");
    }

    const resultado = await respuesta.json();

    if (!resultado.ok) {
        throw new Error(
            resultado.error || "Error al guardar ventas."
        );
    }

    return id;
}

// =====================================================
// OBTENER ARCHIVOS
// =====================================================

async function obtenerArchivos() {
    const registros = await consultarRegistros();
    const grupos = new Map();

    registros.forEach(registro => {
        const id = String(registro.ID ?? "");

        if (!id) return;

        if (!grupos.has(id)) {
            grupos.set(id, {
                id: id,
                nombre: String(registro.NombreArchivo ?? ""),
                fecha: registro.Fecha ?? "",
                ventas: []
            });
        }

        grupos.get(id).ventas.push({
            codigo: String(registro.Codigo ?? ""),
            producto: String(registro.Producto ?? ""),
            cantidad: Number(registro.Cantidad) || 0,
            um: String(registro.UM ?? ""),
            base: Number(registro.Base) || 0,
            total: Number(registro.Total) || 0
        });
    });

    return Array.from(grupos.values());
}

// =====================================================
// CARGAR DATOS GUARDADOS
// =====================================================

async function cargarDatosGuardados() {
    const estado = document.getElementById("estadoTexto");

    try {
        if (estado) {
            estado.textContent = "Consultando Google Sheets...";
        }

        archivosGuardados = await obtenerArchivos();
        datos = [];

        archivosGuardados.forEach(archivo => {
            archivo.ventas.forEach(venta => {
                datos.push({
                    ...venta,
                    archivoId: archivo.id,
                    archivoNombre: archivo.nombre
                });
            });
        });

        actualizarContadorArchivos();
        mostrarHistorial();
        actualizarFiltros();
        actualizarDashboard();

        if (estado) {
            estado.textContent = "Datos sincronizados con Google Sheets.";
        }
    } catch (error) {
        console.error("Error al cargar datos:", error);

        if (estado) {
            estado.textContent = "Error al consultar Google Sheets.";
        }
    }
}

// =====================================================
// SUBIR EXCEL
// =====================================================

document
    .getElementById("archivoExcel")
    .addEventListener(
        "change",
        function (event) {

            const archivos =
                Array.from(
                    event.target.files
                );


            if (
                archivos.length === 0
            ) {

                return;

            }


            procesarArchivos(
                archivos
            );


            event.target.value = "";

        }
    );


// =====================================================
// PROCESAR VARIOS ARCHIVOS
// =====================================================

async function procesarArchivos(
    archivos
) {

    const estado =
        document.getElementById(
            "estadoTexto"
        );


    for (
        const archivo
        of archivos
    ) {

        try {

            estado.textContent =
                "Procesando " +
                archivo.name +
                "...";


            const ventas =
                await leerExcel(
                    archivo
                );


            if (
                ventas.length === 0
            ) {

                continue;

            }


            await guardarArchivo({

                nombre:
                    archivo.name,

                fecha:
                    new Date()
                        .toISOString(),

                ventas:
                    ventas

            });

        } catch (error) {

            console.error(
                error
            );

        }

    }


    await cargarDatosGuardados();


    estado.textContent =
        "Archivos cargados correctamente.";

}


// =====================================================
// LEER EXCEL
// =====================================================

function leerExcel(
    archivo
) {

    return new Promise(
        (resolve, reject) => {

            const lector =
                new FileReader();


            lector.onload =
                function (event) {

                    try {

                        const workbook =
                            XLSX.read(
                                event.target.result,
                                {
                                    type:
                                        "array"
                                }
                            );


                        const primeraHoja =
                            workbook
                                .SheetNames[0];


                        const hoja =
                            workbook.Sheets[
                                primeraHoja
                            ];


                        const filas =
                            XLSX.utils
                                .sheet_to_json(
                                    hoja,
                                    {
                                        defval: ""
                                    }
                                );


                        resolve(
                            procesarFilas(
                                filas
                            )
                        );


                    } catch (error) {

                        reject(
                            error
                        );

                    }

                };


            lector.onerror =
                function () {

                    reject(
                        lector.error
                    );

                };


            lector.readAsArrayBuffer(
                archivo
            );

        }
    );

}


// =====================================================
// PROCESAR FILAS
// =====================================================

function procesarFilas(
    filas
) {

    const resultado = [];


    filas.forEach(
        fila => {

            const producto =
                fila["Producto"] ||
                fila["PRODUCTO"] ||
                "";


            if (
                productoDebeExcluirse(
                    producto
                )
            ) {

                return;

            }


            const cantidad =
                Number(
                    fila["Cantidad"] ||
                    fila["CANTIDAD"] ||
                    0
                );


            const base =
                Number(
                    fila["Base imponible"] ||
                    fila["Base Imponible"] ||
                    fila["BASE IMPONIBLE"] ||
                    0
                );


            const total =
                Number(
                    fila["Total"] ||
                    fila["TOTAL"] ||
                    0
                );


            if (
                cantidad <= 0
            ) {

                return;

            }


            resultado.push({

                codigo:
                    fila["Código"] ||
                    fila["Codigo"] ||
                    fila["CÓDIGO"] ||
                    "",

                producto:
                    String(
                        producto
                    ).trim(),

                cantidad:
                    cantidad,

                um:
                    fila["UM"] ||
                    "",

                base:
                    base,

                total:
                    total

            });

        }
    );


    return resultado;

}


// =====================================================
// ACTUALIZAR CONTADOR DE ARCHIVOS
// =====================================================

function actualizarContadorArchivos() {

    const elemento =
        document.getElementById(
            "cantidadArchivos"
        );


    if (elemento) {

        elemento.textContent =
            archivosGuardados.length;

    }

}


// =====================================================
// MOSTRAR HISTORIAL
// =====================================================

function mostrarHistorial() {

    const contenedor =
        document.getElementById(
            "historialArchivos"
        );


    if (!contenedor) {

        return;

    }


    contenedor.innerHTML = "";


    if (
        archivosGuardados.length === 0
    ) {

        contenedor.innerHTML = `

            <p class="vacio">
                No hay archivos guardados.
            </p>

        `;

        return;

    }


    archivosGuardados
        .slice()
        .reverse()
        .forEach(
            archivo => {

                const fecha =
                    new Date(
                        archivo.fecha
                    );


                const div =
                    document.createElement(
                        "div"
                    );


                div.className =
                    "historial-item";


                div.innerHTML = `

                    <div>

                        <strong>
                            ${archivo.nombre}
                        </strong>

                        <br>

                        <small>
                            ${fecha.toLocaleString()}
                        </small>

                    </div>


                    <button
                        class="btn-eliminar"
                        onclick="
                            eliminarArchivoGuardado(
                                ${archivo.id}
                            )
                        "
                    >
                        Eliminar
                    </button>

                `;


                contenedor.appendChild(
                    div
                );

            }
        );

}


// =====================================================
// ELIMINAR ARCHIVO
// =====================================================

async function eliminarArchivoGuardado(
    id
) {

    if (
        !confirm(
            "¿Quieres eliminar este archivo?"
        )
    ) {

        return;

    }


    await eliminarArchivo(id);

    await cargarDatosGuardados();

}


// =====================================================
// BOTÓN ELIMINAR TODO
// =====================================================

document
    .getElementById("btnBorrarTodo")
    .addEventListener(
        "click",
        async function () {

            if (
                !confirm(
                    "¿Seguro que quieres eliminar todos los archivos?"
                )
            ) {

                return;

            }


            await eliminarTodos();

            await cargarDatosGuardados();

        }
    );


// =====================================================
// FILTROS
// =====================================================

document
    .getElementById("filtroProducto")
    .addEventListener(
        "change",
        actualizarDashboard
    );


document
    .getElementById("filtroArchivo")
    .addEventListener(
        "change",
        actualizarDashboard
    );


document
    .getElementById("minimoVendidos")
    .addEventListener(
        "change",
        actualizarDashboard
    );


// =====================================================
// ACTUALIZAR FILTROS
// =====================================================

function actualizarFiltros() {

    const filtroProducto =
        document.getElementById(
            "filtroProducto"
        );


    const filtroArchivo =
        document.getElementById(
            "filtroArchivo"
        );


    const productoActual =
        filtroProducto.value;


    const archivoActual =
        filtroArchivo.value;


    filtroProducto.innerHTML = `

        <option value="">
            Todos los productos
        </option>

    `;


    const productos =
        [
            ...new Set(
                datos.map(
                    venta =>
                        venta.producto
                )
            )
        ].sort();


    productos.forEach(
        producto => {

            filtroProducto.innerHTML += `

                <option value="${producto}">
                    ${producto}
                </option>

            `;

        }
    );


    filtroProducto.value =
        productoActual;


    filtroArchivo.innerHTML = `

        <option value="">
            Todos los archivos
        </option>

    `;


    archivosGuardados.forEach(
        archivo => {

            filtroArchivo.innerHTML += `

                <option
                    value="${archivo.id}"
                >
                    ${archivo.nombre}
                </option>

            `;

        }
    );


    filtroArchivo.value =
        archivoActual;

}


// =====================================================
// DATOS FILTRADOS
// =====================================================

function obtenerDatosFiltrados() {

    const producto =
        document.getElementById(
            "filtroProducto"
        ).value;


    const archivo =
        document.getElementById(
            "filtroArchivo"
        ).value;


    return datos.filter(
        venta => {

            const coincideProducto =
                !producto ||
                venta.producto ===
                producto;


            const coincideArchivo =
                !archivo ||
                String(
                    venta.archivoId
                ) ===
                String(
                    archivo
                );


            return (
                coincideProducto &&
                coincideArchivo
            );

        }
    );

}


// =====================================================
// ACTUALIZAR DASHBOARD
// =====================================================

function actualizarDashboard() {

    const ventas =
        obtenerDatosFiltrados();


    let totalUnidades = 0;

    let totalDinero = 0;


    const productos = {};


    ventas.forEach(
        venta => {

            totalUnidades +=
                Number(
                    venta.cantidad
                );


            totalDinero +=
                Number(
                    venta.total
                );


            if (
                !productos[
                    venta.producto
                ]
            ) {

                productos[
                    venta.producto
                ] = {

                    nombre:
                        venta.producto,

                    unidades:
                        0,

                    dinero:
                        0

                };

            }


            productos[
                venta.producto
            ].unidades +=
                Number(
                    venta.cantidad
                );


            productos[
                venta.producto
            ].dinero +=
                Number(
                    venta.total
                );

        }
    );


    document
        .getElementById(
            "totalProductos"
        )
        .textContent =
        totalUnidades.toLocaleString(
            "es-PE"
        );


    document
        .getElementById(
            "productosDiferentes"
        )
        .textContent =
        Object.keys(
            productos
        ).length;


    document
        .getElementById(
            "totalVendido"
        )
        .textContent =
        "S/ " +
        totalDinero.toFixed(2);


    const promedio =
        totalUnidades > 0
            ? totalDinero /
              totalUnidades
            : 0;


    document
        .getElementById(
            "promedioProducto"
        )
        .textContent =
        "S/ " +
        promedio.toFixed(2);


    mostrarProductosMasVendidos(
        productos
    );


    mostrarProductosMasDinero(
        productos
    );


    mostrarTablaProductos(
        productos
    );


    mostrarTablaDetalle(
        ventas
    );

}


// =====================================================
// PRODUCTOS MÁS VENDIDOS
// =====================================================

function mostrarProductosMasVendidos(
    productos
) {

    const contenedor =
        document.getElementById(
            "rankingProductos"
        );


    const minimo =
        Number(
            document.getElementById(
                "minimoVendidos"
            ).value
        );


    const ranking =
        Object.values(
            productos
        )

        .filter(
            producto =>
                producto.unidades >=
                minimo
        )

        .sort(
            (a, b) =>
                b.unidades -
                a.unidades
        );


    contenedor.innerHTML = "";


    if (
        ranking.length === 0
    ) {

        contenedor.innerHTML = `

            <p class="vacio">
                No hay productos que cumplan este filtro.
            </p>

        `;

        return;

    }


    ranking.forEach(
        (producto, index) => {

            contenedor.innerHTML += `

                <div class="ranking-item">

                    <span>
                        ${index + 1}.
                        ${producto.nombre}
                    </span>

                    <strong>
                        ${producto.unidades}
                        unidades
                    </strong>

                </div>

            `;

        }
    );

}


// =====================================================
// PRODUCTOS QUE MÁS DINERO GENERAN
// =====================================================

function mostrarProductosMasDinero(
    productos
) {

    const contenedor =
        document.getElementById(
            "rankingDinero"
        );


    const ranking =
        Object.values(
            productos
        )

        .sort(
            (a, b) =>
                b.dinero -
                a.dinero
        )

        .slice(
            0,
            10
        );


    contenedor.innerHTML = "";


    if (
        ranking.length === 0
    ) {

        contenedor.innerHTML = `

            <p class="vacio">
                No hay datos disponibles.
            </p>

        `;

        return;

    }


    ranking.forEach(
        (producto, index) => {

            contenedor.innerHTML += `

                <div class="ranking-item">

                    <span>
                        ${index + 1}.
                        ${producto.nombre}
                    </span>

                    <strong>
                        S/
                        ${producto.dinero.toFixed(2)}
                    </strong>

                </div>

            `;

        }
    );

}


// =====================================================
// TABLA RESUMEN
// =====================================================

function mostrarTablaProductos(
    productos
) {

    const tabla =
        document.getElementById(
            "tablaProductos"
        );


    tabla.innerHTML = "";


    Object.values(
        productos
    )

    .sort(
        (a, b) =>
            b.unidades -
            a.unidades
    )

    .forEach(
        producto => {

            tabla.innerHTML += `

                <tr>

                    <td>
                        ${producto.nombre}
                    </td>

                    <td>
                        ${producto.unidades}
                    </td>

                    <td>
                        S/
                        ${producto.dinero.toFixed(2)}
                    </td>

                </tr>

            `;

        }
    );

}


// =====================================================
// TABLA DETALLE
// =====================================================

function mostrarTablaDetalle(
    ventas
) {

    const tabla =
        document.getElementById(
            "tablaDetalle"
        );


    tabla.innerHTML = "";


    ventas.forEach(
        venta => {

            tabla.innerHTML += `

                <tr>

                    <td>
                        ${venta.codigo}
                    </td>

                    <td>
                        ${venta.producto}
                    </td>

                    <td>
                        ${venta.cantidad}
                    </td>

                    <td>
                        ${venta.um}
                    </td>

                    <td>
                        S/
                        ${Number(
                            venta.base
                        ).toFixed(2)}
                    </td>

                    <td>
                        S/
                        ${Number(
                            venta.total
                        ).toFixed(2)}
                    </td>

                    <td>
                        ${venta.archivoNombre || ""}
                    </td>

                </tr>

            `;

        }
    );

}


// =====================================================
// INICIAR DASHBOARD
// =====================================================

document.addEventListener(
    "DOMContentLoaded",
    function () {

        cargarDatosGuardados();

    }
);
