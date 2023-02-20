// eje 
function eje(x, y, z, J, i, color){
    ctx.strokeStyle = color
    ctx.lineWidth= i
    ctx.beginPath();
    ctx.moveTo(x, y); //100, 300
    ctx.lineTo((x + z), (y + J));
    ctx.stroke();
}


// carga

function carga(x, y, z, color){
    ctx.strokeStyle = color
    ctx.lineWidth= 3
    ctx.beginPath();
    ctx.moveTo(x, y - z); //150, 250
    ctx.lineTo(x, (y + 50));
    ctx.stroke();
    triangulo(x, y, color)
}
 // triangulo
  function triangulo(x, y, color){
    
    ctx.beginPath();
    ctx.moveTo(x, (y + 50));
    ctx.lineTo((x - 10), (y + 30));
    ctx.lineTo((x + 10), (y + 30));
    ctx.lineTo(x, (y + 50));
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();
        
  }

  function triangulo2(x, y, color){
    
    ctx.beginPath();
    ctx.moveTo(x, (y + 10));
    ctx.lineTo((x - 10), (y + 30));
    ctx.lineTo((x + 10), (y + 30));
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();

        
  }