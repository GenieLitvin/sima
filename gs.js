var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var login="";
var passw="";
var errs=[];

function onFormSubmit(event) {
  var lr=sheet.getLastRow(); 
  var values = event.namedValues;
  var sid = values["Артикул "]; /// sid, артикул товара
  var qty = values["Количество"]; ///кол-во товара
  
  var items=[{"sid":+sid,"qty": +qty}]; 
  var delivery = getDeliveryCost(items); ///стоимость доставки товара
  
  var data = loadData(sid);
  data.cost = delivery[sid].cost.toString().replace(".", ",");

  pastDataToSheet(lr, data);

}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sima-Land')
      .addItem('Загрузить данные','start')
      .addItem('Добавить в корзину','add')
      .addToUi();
}


function auth(nlogin, npassw){
  login = nlogin;
  passw=npassw;
  
  add();
}

function add(){
  if (login=="") return getLogin();
  if (passw=="") return getLogin();
 
  var ranges = sheet.getSelection().getActiveRangeList().getRanges(); ///выделеные диапазоны
  for (var i = 0; i < ranges.length; i++) {
     addRange(ranges[i]);  
  }
}


function getLogin(){
  
  var interface = HtmlService.createHtmlOutputFromFile("auth.html").setWidth(300)
         .setHeight(200)
   ;
  
  var UiApp=SpreadsheetApp.getUi().showModalDialog(interface, 'Авторизация');
  
  };


function addRange(range){
 
     var sidValues = range.getValues();    ///массив значений выделеного диапазона
     var firstRow = range.getRow();            /// номер первой строки  
     var curRow=firstRow;                     ///текущая строка
     var cnt = sidValues.length;
  
     for (var i = 0; i < cnt; i++) {
       var sid = sidValues[i].toString(); /// sid, артикул товара
       var qty = sheet.getRange(curRow,CTY_N).getValue(); ///кол-во товара
       var id = sheet.getRange(curRow,ID_N).getValue();
       addToCart(id,sid,qty, curRow);
       curRow = curRow +1;
     }
   if (errs.length>0) throw (errs.join(";"));
  

}

function addToCart(id,sid,qty, curRow){
 
    var headers = {
      "Authorization" : "Basic " + Utilities.base64Encode(login + ':' + passw)
    };
    
    var data = {"item_id":id,"item_sid":sid,"cart_id":"main","qty":qty,"has_been_added_by_the_piece":true};
    var params = {
      'method' : 'post',
      'contentType': 'application/json',
      "headers":headers,
      'payload' : JSON.stringify(data)
    }; 
    var url = 'https://www.sima-land.ru/api/v3/cart-item/';

  try{  
   
    var response = UrlFetchApp.fetch(url, params);
    sheet.getRange(curRow,STATE_N).setValue("В корзине");
    /*Logger.log(response.getContentText());*/
  }catch(e){
    errs.push(e); //e[1].message
    Logger.log('ERROR')
    sheet.getRange(curRow,STATE_N).setValue(e);
  }
  
}

function start() {
  
  var ranges = sheet.getSelection().getActiveRangeList().getRanges(); ///выделеные диапазоны
  var items = []; //!!
  for (var i = 0; i < ranges.length; i++) {
     loadRange(ranges[i]);  
  }

}

/// загрузка данных одного диапазона
function loadRange(range){
     
     var items = [];
     var sidValues = range.getValues();    ///массив значений выделеного диапазона
     var firstRow = range.getRow();            /// номер первой строки  
     var curRow=firstRow;                     ///текущая строка
     var  cnt = sidValues.length;
  
     for (var i = 0; i < cnt; i++) {
       var sid = sidValues[i].toString(); /// sid, артикул товара
       var qty = sheet.getRange(curRow,CTY_N).getValue(); ///кол-во товара
       items.push({"sid":sid ,"qty": qty})       
       curRow = curRow +1;
     } 
  
     var delivery = getDeliveryCost(items); ///стоимость доставки по всем товарам  
  
     curRow = firstRow;  
     for (var j = 0; j < cnt; j++) {       
       var sid = sidValues[j];  
       
       var data = loadData(sid);    /// TODO     
       data.cost = delivery[sid].cost.toString().replace(".", ",");
       pastDataToSheet(curRow,data);       
       curRow = curRow+1;     
     }

}

///вставка полученных данных в таблицу
function pastDataToSheet(row,data){
       sheet.getRange(row,DELIVERY_N).setValue(data.cost);        
       sheet.getRange(row,NAME_N).setValue(data.name);     
       sheet.getRange(row,PRICE_N).setValue(data.price);     
       sheet.getRange(row,WHOLESALE_N).setValue(data.wholesale_price);
       sheet.getRange(row,CATEGORY_N).setValue(data.category);
       sheet.getRange(row,ID_N).setValue(data.id); 
}

///загрузка данных о товаре с симы
function loadData(sid){
  
   var response = UrlFetchApp.fetch("https://www.sima-land.ru/api/v3/item/?sid="+sid);  
   var json = response.getContentText();
   var data = JSON.parse(json);
   var item = data.items[0];
  
   var price = item.price.toString().replace(".", ",");
   var wholesale_price = item.wholesale_price.toString().replace(".", ",");
   var isOpt = false;
   if ((item.action_urls)&&(item.action_urls['special-offer'])&&(item.action_urls['special-offer'].includes('opt'))) isOpt=true;
   return {'id':item.id,'name':item.name, 'price': price,'wholesale_price':wholesale_price, 'category': getCategory(item.special_offer_id, isOpt)} ;
}

function getDeliveryCost(items){
  var data =  {"items":items, "settlement_id":MINSK}; 
  
  var options = {
  'method' : 'post',
  'contentType': 'application/json',
  'payload' : JSON.stringify(data)
  };
  
  var response = UrlFetchApp.fetch('https://www.sima-land.ru/api/v3/delivery-calc/', options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}

function getCategory(special_offer_id, isOpt){
    if (categories[special_offer_id])  return categories[special_offer_id]; 
    if (isOpt) return "Опт";
    return "";
  
}
var categories = {
    '82':'Авто и мото',
    '89':'Семена',
    '122':'Светильники',
    '80':'Баня и сауна',
    '46':'Электротовары',
    '121':'Опт',
    '12':'Спорт и отдых',
    '120':'Праздничное освещение',
    '2':'Текстиль, одежда и обувь',
    '84':'Сад и огород',
    '79':'Мебель',
    '86':'Средства для сада',
    '5':'Канцтовары',
    '83':'Бижутерия и оборудование',
    '81':'Галантерея и швейная галантерея',
     '3':'Детские товары',
    '8':'Инструменты и сантехника',
     '85':'Кожгалантерея',
     '11':'Подарочные продукты питания',
     '9':'Посуда и хозтовары',
    '13':'Рыбалка'
  }

var MINSK = 26162465;
var CTY_N = 5; ///кол-во товара
var NAME_N = 4; ///наименование
var PRICE_N = 8; ///цена
var WHOLESALE_N = 9; /// оптовая цена
var DELIVERY_N =10 ; ///стоимость доставки
var CATEGORY_N = 11; ///категория товара
var ID_N=17; ///id товара
var STATE_N=18; ///статус в корзине
