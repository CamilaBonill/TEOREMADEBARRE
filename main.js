const canvas = document.querySelector('canvas');
const ctx = canvas.getContext('2d');


//variable
var color = "#82C2E5"

/*------------------------inicio---------------------------*/

const  dist1 = parseFloat(document.getElementById('dist1').value);
const  dist2 = parseFloat(document.getElementById('dist2').value);
dibujarcarga(dist1, dist2)

/*------------------------Calculos de la posicion de la resultante y Eje En el Camion---------------------------*/

function posicionEjeCamion(){
    //variable
    const dist1 = parseFloat(document.getElementById('dist1').value);
    const dist2 = parseFloat(document.getElementById('dist2').value);
    const Carga1 = parseFloat(document.getElementById('carga1').value);
    const Carga2 = parseFloat(document.getElementById('carga2').value);
    const Carga3 = parseFloat(document.getElementById('carga3').value);

    // resultante
    var resultante = (Carga1 + Carga2 + Carga3);
    document.getElementById("Resultante").value = resultante;

    // calcular la posicion de la resultante,

    //let PR = (Carga2*(dist1 + (dist1 + dist2)))/resultante
    let PR = ((Carga2*dist2) + (Carga1*(dist1 + dist2)))/resultante
    document.getElementById("Posicion_Rsultante").value = PR;

    // calcular la posicion del eje
    let PE = (dist2 - PR)/2
    console.log(PR, PE)
    
    document.getElementById("Posicion_EjeLuz").value = PE;
    dibujarcarga(dist1, dist2)
    dibujarclaroluz(PE, PR, dist1, dist2)
    
}

/*------------------------Dibujo---------------------------*/
// eje horizontal
eje(100, 300, 400, 0, 4, "#000000")
cuadricula()

function dibujarcarga(dist1, dist2) {
    //rueda
    carga(160, 250, 0, "red")
    // distancia ente rueda 4.3 m, 25 es igual a metro
    carga((160+(25*dist1)), 250, 0, "red")
    carga((160+(25*(dist1 + dist2))), 250, 0, "red")

    // distancia entre rueda
    eje(160, 350, (25*(dist1 + dist2)), 0, 2, "#ae9c8f")
    eje(160, 345, 0, 10, 2, "#ae9c8f")
    eje((160 + (25*dist1)), 345, 0, 10, 2, "#ae9c8f")
    eje((160 + (25*(dist1 + dist2))), 345, 0, 10, 2, "#ae9c8f")

}


function dibujarclaroluz(PE, PR, dist1){
    //resultante
    // Carga de medio fija en 567.5
    let PC2 = 267.5
    carga((PC2 +(25*(dist2 - PR))), 250, 30, "green")

    // distancia x medida desde el eje delantero
    eje((PC2), 320, (25*PR), 0, 2, "#ae9c8f")
    eje(PC2, 315, 0, 10, 2, "#ae9c8f")
    eje((PC2 + (25*PR)), 315, 0, 10, 2, "#ae9c8f")

    // distancia el eje y la carga mas cercana a la resultantes
    eje((PC2), 330, (PE*25), 0, 2, "#ae9c8f")
    eje((PC2), 325, 0, 10, 2, "#ae9c8f")


    //eje claro de luz
    eje((PC2 + (PE*25)), 200, 0, 200, 3, "#ff9800")
    return  CL = (PC2 + (PE*25))
}

/*------------------------Dibujo para colocar el puente---------------------------*/
function ubicarcamion(){
    const LP = parseFloat(document.getElementById('longitud_puente').value);
    let IP = (CL - (LP*25)/2)
    let IF = (CL + (LP*25)/2)
    eje(IP, 300, (LP*25), 0, 6, "#6f3460")
    triangulo2(IP, 290, "#52fd00")
    triangulo2(IF, 290, "#52fd00")
    eje(CL, 225, 0, 150, 5, "#55ffe2")
    console.log(LP)
}


//actualiza la pantalla cuando el usuario hace click en el botón "nuevo juego"

function upgrade(){
    location.reload();
}

