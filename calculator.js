var INPUT_RANGE = "A2:D123"
var INPUT_SHEET_NAME = "Merged"

function main() {
  var transactions = readTransactions()
  var acquisitions = []
  var sales = []
  for (var i in transactions) {
    var tx = transactions[i]
    var copy = Object.assign({}, tx)
    if (tx.quantity > 0) {
      acquisitions.push(copy)
    }
    if (tx.quantity < 0) {
      copy.quantity = Math.abs(copy.quantity)
      sales.push(copy)
    }
  }
  dumpTransactions(transactions)
  dumpTransactions(acquisitions)
  dumpTransactions(sales)
  var gain = 0
  var sameDayResults = []
  var sameMonthResults = []
  var poolResults = []
  for (var i in sales) {
    var tx = sales[i]
    var sameDay = sellByDateDiff(tx, tx.quantity, acquisitions, 0)
    var sameMonth = sellByDateDiff(sales[i], sameDay.toSell, acquisitions, 30)
    var fromPool = sellByPool(sales[i], sameMonth.toSell, acquisitions)
    sameDayResults.push(sameDay)
    sameMonthResults.push(sameMonth)
    poolResults.push(fromPool)
    gain += sameDay.gain + sameMonth.gain + fromPool.gain
    Logger.log(Utilities.formatString(
      "%s\tSold %.2f gain=%.2f sameDay=%.2f sameMonth=%.2f pool=%.2f prices=%s, numPerPrice=%s",
      dateString(tx.date),
      tx.quantity,
      sameDay.gain + sameMonth.gain + fromPool.gain, 
      sameDay.numSold, 
      sameMonth.numSold, 
      fromPool.numSold,
      sameDay.salePrices + sameMonth.salePrices + fromPool.salePrices, 
      sameDay.saleAmounts + sameMonth.saleAmounts + fromPool.saleAmounts))
  }
  renderResults(sales, sameDayResults, sameMonthResults, poolResults)
  Logger.log("Total gain " + gain)
}

function sellByDateDiff(tx, toSell, acquisitions, numDays) {
  var gain = 0
  var salePrices = []
  var saleAmounts = []
  var cost = 0
  for (var i in acquisitions) {
    var lot = acquisitions[i]
    if (Math.abs(datediff(tx.date, lot.date)) <= numDays && tx.date < lot.date) {
      var soldFromLot = Math.min(lot.quantity, toSell)
      gain += (tx.price * tx.fxRate - lot.price * lot.fxRate) * soldFromLot
      toSell -= soldFromLot
      lot.quantity -= soldFromLot
      salePrices.push(lot.price)
      saleAmounts.push(soldFromLot)
      cost += lot.price * soldFromLot * lot.fxRate
    }
    if (toSell == 0) {
      break
    }
    if (toSell < 0) {
      throw new Error("tosell < 0 " + toSell)
    }
  }
  return {
    gain: gain,
    numSold: tx.quantity - toSell,
    toSell: toSell,
    cost: cost,
    salePrices: salePrices,
    saleAmounts: saleAmounts,
  }
}

function sellByPool(tx, toSell, acquisitions) {
  var numShares = 0
  var avgPrice = 0
  for (var i in acquisitions) {
    var aq = acquisitions[i]
    if (aq.date < tx.date && aq.quantity > 1e-6) {
      avgPrice = (avgPrice * numShares + aq.quantity * aq.price * aq.fxRate) / (numShares + aq.quantity)
      numShares += aq.quantity
    }
  }
  if (numShares - toSell < 0 && Math.abs(numShares - toSell) > 1e-6) {
    throw new Error("not enough in pool to sell " + toSell + ' ' + numShares)
  }
  for (var i in acquisitions) {
    var aq = acquisitions[i]
    if (aq.date < tx.date) {
      aq.quantity = Math.max(0, aq.quantity - toSell * (aq.quantity / numShares))
    }
  }
  var gain = toSell * (tx.price * tx.fxRate - avgPrice) 
  return {
    gain: gain,
    numSold: toSell,
    toSell: 0,
    cost: avgPrice * toSell,
    salePrices: [avgPrice],
    saleAmounts: [toSell],
    numSharesInPool: numShares,
  }
}

function datediff(first, second) {        
  return Math.round((second - first) / (1000 * 60 * 60 * 24));
}

function readTransactions() {
  var range = SpreadsheetApp.getActive().getSheetByName(INPUT_SHEET_NAME).getRange(INPUT_RANGE)
  var transactions = []
  for (var i = 1; i <= range.getNumRows(); ++i) {
    transactions.push({
      date: range.getCell(i, 1).getValue(),
      price: range.getCell(i, 2).getValue(),
      quantity: range.getCell(i, 3).getValue(),
      fxRate: range.getCell(i, 4).getValue(),
    })
  }
  return transactions
}

function dumpTransactions(txn) {
  for (var i in txn) {
    var tx = txn[i]
    Logger.log(Utilities.formatString("%s\t%s\t%.4f\t%.2f\t%f", dateString(tx.date), tx.quantity < 0 ? "SELL" : "BUY", tx.quantity, tx.price, tx.fxRate))
  }
}

function dateString(date) {
  return Utilities.formatDate(date, "Europe/London", "yyyy-MM-dd")
}

function renderResults(sales, dayResults, monthResults, poolResults) {
  var sheetData = [
    [
      "Date", 
      "Quantity", 
      "Price", 
      "FX rate", 
      "Gain", 
      "Proceeds",
      "Cost",
      "Num sold same day", 
      "Same day cost GBP", 
      "Num sold same Month", 
      "Same month cost gbp", 
      "Num sold pool", 
      "Pool cost GBP",
      "Num shares in pool",
    ]
  ]
  for (var i in sales) {
    var tx = sales[i]
    var day = dayResults[i]
    var month = monthResults[i]
    var pool = poolResults[i]
    sheetData.push(
      [
        tx.date, 
        tx.quantity, 
        tx.price, 
        tx.fxRate,
        day.gain + month.gain + pool.gain, 
        tx.quantity * tx.price * tx.fxRate,
        day.cost + month.cost + pool.cost,
        day.numSold.toString(), 
        day.salePrices.toString(), 
        month.numSold.toString(), 
        month.salePrices.toString(), 
        pool.numSold, 
        pool.salePrices,
        pool.numSharesInPool,
      ]
    )
  }
  var sheet = SpreadsheetApp.getActive().getSheetByName("Results")
  sheet.clear()
  var neededCols = Math.max(0, 1 + sheetData[0].length - sheet.getMaxColumns())
  var neededRows = Math.max(0, 1 + sheetData.length - sheet.getMaxRows())
  if (neededCols > 0) {
    sheet.insertColumns(1, neededCols)
  }
  if (neededCols > 0) {
    sheet.insertRows(1, neededRows)
  }
  var range = sheet.getRange(1,1,sheetData.length, sheetData[0].length)
  range.setValues(sheetData)
}
