/* 
 *   Copyright 2020 Daniel Argüelles
 *
 *   Licensed under the Apache License, Version 2.0 (the "License");
 *   you may not use this file except in compliance with the License.
 *   You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 *   Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *   See the License for the specific language governing permissions and
 *   limitations under the License.
 */


function recalculateHistoric () {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var calsheet = ss.getSheetByName("Operaciones");
    var source = calsheet.getRange("A:N");
    var values = source.getValues();

    var historics = {}
    var marketsByCountry = getMarketsByCountry();

    var historicSheet = ss.getSheetByName("Histórico");

    // Clear historic sheet
    historicSheet.insertRowsAfter(historicSheet.getMaxRows(), 1)  // Add empty row to avoid loss styles
    historicSheet.deleteRows(2, historicSheet.getMaxRows() - 2);

    var my_historic = {};
    for (var i = 1; i < values.length; i++) {

        // Empty row?
        if (values[i][1] == "") continue;

        var row = values[i];
        var operation_date = getDayFromDate(row[2]);
        var operation_actions, operation_cost;
        switch (row[1]) {
            case "Venta":
                operation_actions = -1 * row[8];
                operation_cost = row[13];
                break;
            case "Dividendo":
                operation_actions = 0;
                operation_cost = -1 * row[13];
                break;
            case "Compra":
            case "Script":
                operation_actions = row[8];
                operation_cost = row[13];
                break;
        }

        // Get symbol historic if not gathered yet
        if (!historics.hasOwnProperty(row[5])) {
            historics[row[5]] = getHistoric(row[5] + marketsByCountry[row[6]], row[2] / 1000);
        }

        var symbolCurrency = historics[row[5]].currency;

        // Get currency exchange
        if (symbolCurrency != "EUR" && !historics.hasOwnProperty(symbolCurrency)) {
            historics[symbolCurrency] = getHistoric(symbolCurrency + "%3DX", row[2] / 1000);
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
                    my_historic[day] = { 'input': 0, 'value': 0 }
                }


                var currency_price = symbolCurrency == "EUR" ? 1 : historics[symbolCurrency].close[j];
                var tmp = j - 1;
                while (currency_price == null) {
                    currency_price = historics[symbolCurrency].close[tmp];
                    tmp = tmp - 1;
                }

                if (symbolCurrency == "GBp" && row[7] == "£") {
                    currency_price = currency_price * 100;
                }


                if (+day === +operation_date) {
                    my_historic[day] = {
                        'input': my_historic[day].input + operation_cost,
                        'value': my_historic[day].value + (operation_actions * historics[row[5]].close[j]) / currency_price
                    }
                } else if (+day >= +operation_date) {
                    my_historic[day] = {
                        'input': my_historic[day].input,
                        'value': my_historic[day].value + (operation_actions * historics[row[5]].close[j]) / currency_price
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

    var table = []
    var i = 1;
    for (var key in my_historic) {
        while (true) { // PFFFFFFF to improve
            i = i + 1;
            day = getDayFromDate(key)
            appendArray = []
            appendArray.push(day.getDate() + "/" + (day.getMonth() + 1) + "/" + day.getFullYear())
            appendArray.push(my_historic[key].input == 0 ? "" : my_historic[key].input);
            appendArray.push(my_historic[key].value != null ? my_historic[key].value : "=C" + (i - 1));
            appendArray.push((i == 2) ? '=C2/E2' : '=if(B' + i + '="";D' + (i - 1) + ';C' + i + '/E' + i + ')');
            appendArray.push((i == 2) ? 100 : '=if(B' + i + '="";C' + i + '/D' + i + ';E' + (i - 1) + ')');
            appendArray.push((i == 2) ? '' : '=(E' + i + '-E' + (i - 1) + ')/E' + (i - 1));
            appendArray.push('')
            appendArray.push(my_historic[key].ibex != null ? my_historic[key].ibex : "=H" + (i - 1))
            appendArray.push((i == 2) ? 100 : '=I' + (i - 1) + '*(1+J' + i + ')');
            appendArray.push((i == 2) ? '' : '=(H' + i + '-H' + (i - 1) + ')/H' + (i - 1));
            appendArray.push('');
            appendArray.push(my_historic[key].ibex != null ? my_historic[key].sp500 : "=L" + (i - 1))
            appendArray.push((i == 2) ? 100 : '=M' + (i - 1) + '*(1+N' + i + ')');
            appendArray.push((i == 2) ? '' : '=(L' + i + '-L' + (i - 1) + ')/L' + (i - 1));
            appendArray.push('');
            appendArray.push(my_historic[key].ibex != null ? my_historic[key].eur600 : "=P" + (i - 1))
            appendArray.push((i == 2) ? 100 : '=Q' + (i - 1) + '*(1+R' + i + ')');
            appendArray.push((i == 2) ? '' : '=(P' + i + '-P' + (i - 1) + ')/P' + (i - 1));

            table.push(appendArray);

            if (my_historic[key].input != 0) {
                my_historic[key].input = 0;
            } else {
                break;
            }
        }
    }

    // Write result in the sheet
    historicSheet.insertRowsAfter(historicSheet.getMaxRows(), table.length);
    var range = historicSheet.getRange(2, 1, table.length, table[0].length);
    range.setValues(table);
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

    var responseBody = response.getContentText();
    var responseCode = response.getResponseCode();
    if (responseCode === 200) {
        var json = JSON.parse(responseBody);
        return {
            "timestamp": json.chart.result["0"].timestamp,
            "close": json.chart.result["0"].indicators.quote["0"].close,
            "currency": json.chart.result["0"].meta.currency
        };
    } else {
        throw (Utilities.formatString("Request failed. Expected 200, got %d: %s. URL: %s", responseCode, responseBody, url));
    }
}
