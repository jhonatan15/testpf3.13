document.getElementById('buttn-ini').addEventListener('click', verificar, false)

function verificar(){
  var suma = 0;
  var los_cboxes = document.getElementsByClassName('btn-check');
  for (var i = 0, j = los_cboxes.length; i < j; i++) {

      if(los_cboxes[i].checked == true){
      suma++;
      }
  }
  if(suma < 10){
    alert('debe responder todas las preguntas');
    return false;
  } else {
    document.getElementById("buttn-ini").type = "submit";
  }
}
