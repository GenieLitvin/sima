var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sima-Land')
      .addItem('Загрузить данные','start')
      .addToUi();
}

function start() {
  
  var ranges = sheet.getSelection().getActiveRangeList().getRanges(); ///выделеные диапазоны
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

       var cost = delivery[sid].cost.toString().replace(".", ",");
       sheet.getRange(curRow,DELIVERY_N).setValue(cost); 
       
       var data = loadData(sid);    /// TODO     
       sheet.getRange(curRow,NAME_N).setValue(data.name);     
       sheet.getRange(curRow,PRICE_N).setValue(data.price);     
       sheet.getRange(curRow,WHOLESALE_N).setValue(data.wholesale_price);
       sheet.getRange(curRow,CATEGORY_N).setValue(data.category);    
       
       curRow = curRow+1;     
     }

}


///загрузка данных о товаре с симы
function loadData(sid){
  
   var response = UrlFetchApp.fetch("https://www.sima-land.ru/api/v3/item/?sid="+sid);  
   var json = response.getContentText();
   var data = JSON.parse(json);
   var item = data.items[0];
  
   var price = item.price.toString().replace(".", ",");
   var wholesale_price = item.wholesale_price.toString().replace(".", ","); 
   return {'name':item.name, 'price': price,'wholesale_price':wholesale_price, 'category': getCategory(item.special_offer_id)} ;
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

function getCategory(id){
  var data = {
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
  
    return data[id];
  
}

var MINSK = 26162465;
var CTY_N = 5; ///кол-во товара
var NAME_N = 4; ///наименование
var PRICE_N = 8; ///цена
var WHOLESALE_N = 9; /// оптовая цена
var DELIVERY_N =10 ; ///стоимость доставки
var CATEGORY_N = 11; ///категория товара
