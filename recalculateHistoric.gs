function recalculateHistoric () {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var calsheet = ss.getSheetByName("Operaciones");
    var source = calsheet.getRange("A:N");
    var values = source.getValues();

    var historics = {}
    var marketsByCountry = getMarketsByCountry();
    var currencyTickers = getCurrencyTickers();

    var my_historic = {};
    for (var i = 1; i < values.length; i++) {
        if (values[i][1] == "Dividendo" || values[i][1] == "") continue;

        var row = values[i];
        var operation_date = getDayFromDate(row[2]);
        var operation_actions = (row[1] == "Venta") ? row[8] * -1 : row[8] * 1;
        var operation_cost = row[13];

        // Get symbol historic if not gathered yet
        if (!historics.hasOwnProperty(row[5])) {
            historics[row[5]] = getHistoric(row[5] + marketsByCountry[row[6]], row[2] / 1000);
        }

        // Get currency exchange
        if (row[7] != "€" && !historics.hasOwnProperty(row[7])) {
            historics[row[7]] = getHistoric(currencyTickers[row[7]] + "%3DX", row[2] / 1000);
        }

        for (var j = 0; j < historics[row[5]].timestamp.length; j++) {
            var day = getDayFromDate(historics[row[5]].timestamp[j] * 1000);
            var next_day;
            if (j + 1 >= historics[row[5]].timestamp.length) {
                next_day = new Date(day);
                next_day.setDate(next_day.getDate() + 1)
            }
            else
                next_day = getDayFromDate(historics[row[5]].timestamp[j + 1] * 1000);

            while (+day < +next_day) {

                if (!my_historic.hasOwnProperty(day)) {
                    my_historic[day] = { 'input': 0, 'value': 0, 'ibex': 0, 'sp500': 0, 'eur600': 0 }
                }


                var currency_price = row[7] == "€" ? 1 : historics[row[7]].close[j];
                var tmp = j-1;
                while( currency_price == null){
                    currency_price = historics[row[7]].close[tmp];
                    tmp = tmp -1;
                }

                if (+day === +operation_date) {
                    my_historic[day] = {
                        'input': my_historic[day].input + operation_cost,
                        'value': my_historic[day].value + (operation_actions * historics[row[5]].close[j]) / currency_price,
                        'ibex': 0,
                        'sp500': 0,
                        'eur600': 0
                    }
                } else if (+day >= +operation_date) {
                    my_historic[day] = {
                        'input': my_historic[day].input,
                        'value': my_historic[day].value + (operation_actions * historics[row[5]].close[j]) / currency_price,
                        'ibex': 0,
                        'sp500': 0,
                        'eur600': 0
                    }
                }
              
                day.setDate(day.getDate() + 1)
                day = getDayFromDate(day)
            }

        }

    }

    addMarketData(getHistoric("%5EIBEX", values[1][2] / 1000), "ibex", my_historic);
    addMarketData(getHistoric("%5EGSPC", values[1][2] / 1000), "sp500", my_historic);
    addMarketData(getHistoric("%5ESTOXX", values[1][2] / 1000), "eur600", my_historic);


    var historicSheet = ss.getSheetByName("Histórico");

    // Add empty row to avoid loss styles
    historicSheet.insertRowsAfter(historicSheet.getLastRow(), 1)

    historicSheet.deleteRows(2, historicSheet.getLastRow() - 1);

    var i = 1;
    for (var key in my_historic) {
        while (true) { // PFFFFFFF to improve
            i = i + 1;
            day = getDayFromDate(key)
            appendArray = []
            appendArray.push(day.getDate() + "/" + (day.getMonth() + 1) + "/" + day.getFullYear())
            appendArray.push(my_historic[key].input == 0 ? "" : my_historic[key].input);
            appendArray.push(my_historic[key].value);
            appendArray.push((i == 2) ? '=C2/E2' : '=if(B' + i + '="";D' + (i - 1) + ';C' + i + '/E' + i + ')');
            appendArray.push((i == 2) ? 100 : '=if(B' + i + '="";C' + i + '/D' + i + ';E' + (i - 1) + ')');
            appendArray.push((i == 2) ? '' : '=(E' + i + '-E' + (i - 1) + ')/E' + (i - 1));
            appendArray.push('')
            appendArray.push(my_historic[key].ibex)
            appendArray.push((i == 2) ? 100 : '=I' + (i - 1) + '*(1+J' + i + ')');
            appendArray.push((i == 2) ? '' : '=(H' + i + '-H' + (i - 1) + ')/H' + (i - 1));
            appendArray.push('');
            appendArray.push(my_historic[key].sp500)
            appendArray.push((i == 2) ? 100 : '=M' + (i - 1) + '*(1+N' + i + ')');
            appendArray.push((i == 2) ? '' : '=(L' + i + '-L' + (i - 1) + ')/L' + (i - 1));
            appendArray.push('');
            appendArray.push(my_historic[key].eur600)
            appendArray.push((i == 2) ? 100 : '=Q' + (i - 1) + '*(1+R' + i + ')');
            appendArray.push((i == 2) ? '' : '=(P' + i + '-P' + (i - 1) + ')/P' + (i - 1));


            historicSheet.appendRow(appendArray);

            if (my_historic[key].input != 0) {
                my_historic[key].input = 0;
            } else {
                break;
            }
        }
    }
}

function addMarketData (data, market, my_historic) {
    for (var i = 0; i < data.timestamp.length; i++) {
        var day = getDayFromDate(new Date(data.timestamp[i] * 1000));
        var next_day;
        if (i + 1 >= data.timestamp.length) {
            next_day = new Date(day);
            next_day.setDate(next_day.getDate() + 1)
        }
        else
            next_day = getDayFromDate(data.timestamp[i + 1] * 1000);

        while (+day < +next_day) {
            if (my_historic.hasOwnProperty(day)) {
                my_historic[day][market] = data.close[i];
            }
            day.setDate(day.getDate() + 1)
            day = getDayFromDate(day)
        }
    }

}

function getMarketsByCountry () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var calsheet = ss.getSheetByName("Cálculo");
    var source = calsheet.getRange("8:11");
    var values = source.getValues();

    var markets = {};
    for (var i = 1; i < values[0].length; i++) {
        markets[values[0][i]] = values[3][i];
    }

    return markets;
}

function getCurrencyTickers () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var calsheet = ss.getSheetByName("Cálculo");
    var source = calsheet.getRange("3:5");
    var values = source.getValues();

    var tickers = {};
    for (var i = 1; i < values[0].length; i++) {
        tickers[values[0][i]] = values[2][i];
    }

    return tickers;

}


function formatDate (date) {
    // https://stackoverflow.com/a/23593099/710162
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;

    return [year, month, day].join('-');
}

function getDayFromDate (date) {
    return new Date(new Date(date).setHours(0, 0, 0, 0));
}

function getHistoric (symbol, since) {
    //var since = Math.floor((Date.now()) / 1000)-(10*24*60*60);
    var today = Math.floor(Date.now() / 1000) - (24 * 60);
    var url = "https://query1.finance.yahoo.com/v8/finance/chart/" + symbol + "?symbol=" + symbol + "&period1=" + since + "&period2=" + today + "&interval=1d&includePrePost=true&events=div%7Csplit%7Cearn&lang=es-ES&region=ES&crumb=xBNPEtxqGjk&corsDomain=es.finance.yahoo.com";

    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());
    return { "timestamp": json.chart.result["0"].timestamp, "close": json.chart.result["0"].indicators.quote["0"].close };
}
