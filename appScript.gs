function onOpen() {  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TERMINALES')
      .addItem('calcular', 'terminales_Function')
      .addToUi();
}

function getDistanceFromLatLonInKm(lat1,lon1,lat2,lon2) {
  var R = 6371; // Radius of the earth in km
  var dLat = deg2rad(lat2-lat1);  // deg2rad below
  var dLon = deg2rad(lon2-lon1); 
  var a = 
    Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(deg2rad(lat1)) * Math.cos(deg2rad(lat2)) * 
    Math.sin(dLon/2) * Math.sin(dLon/2)
    ; 
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)); 
  var d = R * c; // Distance in km
  return d;
}

function deg2rad(deg) {
  return deg * (Math.PI/180)
}

function terminales_Function() {
  let x = [];
  let y = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TERMINALES');
  var ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('datos');
  var dataRangeAll = ss.getDataRange();
  var lr = dataRangeAll.getLastRow();
  //Logger.log('Total de registros: '+(lr-1))   
  for(let i = 2; i <= lr; i++){
    var lat2 = Number(ss.getRange(i,3).getValue()); //LAT terminal 
    var lon2 = Number(ss.getRange(i,4).getValue()); //LONG terminal
    var lat1 = Number(ss1.getRange(2,1).getValue()); //LAT user
    var lon1 = Number(ss1.getRange(2,2).getValue()); //LONG user
    let d = getDistanceFromLatLonInKm(lat1,lon1,lat2,lon2);
    //Logger.log('valor='+d);
    //ss.getRange(i,5).setValue(d);
    //let terminal = 'En '+ss.getRange(i,1).getValue() + ': '+ss.getRange(i,2).getValue();
    let terminal = ss.getRange(i,1).getValue();
    //let distancia = ss.getRange(i,5).getValue();
    //Logger.log(terminal,distancia)
    x.push({terminal,d});
    //Logger.log(x)
  }
  x.sort(function (a, b) {
  if (a.d > b.d) {
    return 1;
  }
  if (a.d < b.d) {
    return -1;
  }
  // a must be equal to b
  return 0;
});
    Logger.log(x);
    //array para solo Ciudades
    for(let j = 0; j< x.length; j++){
      y.push(x[j].terminal);
    }
    //quitar repetidas
    let result = y.filter((item,index)=>{
      return y.indexOf(item) === index;
    })
    //Logger.log('Las 3 terminales mas cercanas son: '+x[0].terminal+', '+x[1].terminal+', '+x[2].terminal+';'); 
    Logger.log('Las 3 ciudades mas cercanas son: '+result[0]+', '+result[1]+', '+result[2]+';');
    ss1.getRange(2,3).setValue(result[0]);
    ss1.getRange(3,3).setValue(result[1]);
    ss1.getRange(4,3).setValue(result[2]);
}
