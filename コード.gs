function viewstop(){
  price = getStockPriceGoogle(8001)

  Logger.log(price)
}
function getStockPriceGoogle(code){
  let url = "https://www.google.com/finance/quote/" + code + ":TYO";
  let html = UrlFetchApp.fetch(url).getContentText();
  let stockPrice = Parser.data(html)
    .from("<div class=\"YMlKec fxKbKc\">")
    .to("</div>")
    .build();
  return stockPrice;
}

