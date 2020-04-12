function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Настройки')
      .addItem('Генерация игры', 'generateGame')
      .addToUi();
}


function generateGame() {
    var sidebarHTML = '<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">';
    sidebarHTML += '<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>';

    sidebarHTML += '<form style="padding: 20px;text-align:center;">\
        <div class="form-group">\
            <label for="playersNames">Имена игроков\n(по одному на строке)</label>\
            <textarea class="form-control" id="playersNames" name="playersNames" rows="5"></textarea> \
        </div>\
        <div class="form-group">\
            <label for="nDarkRounds">Количество тёмных раундов</label>\
            <input type="number" class="form-control" id="nDarkRounds" name="nDarkRounds" value=4></textarea> \
        </div>\
        <div class="form-group">\
            <label for="nGoldRounds">Количество золотых раундов</label>\
            <input type="number" class="form-control" id="nGoldRounds" name="nGoldRounds" value=4></textarea> \
        </div>\
        <div class="form-group">\
            <label for="nDiamondRounds">Количество алмазных раундов</label>\
            <input type="number" class="form-control" id="nDiamondRounds" name="nDiamondRounds" value=1></textarea> \
        </div>\
        <button type="submit" class="btn btn-primary">Сгенерировать игру</button>\
    </form>';

    sidebarHTML += "<script>\
     $(document).on('submit', 'form', function () { \
        google.script.run \
     .writeStrInTable(\
{ playersNames: $('#playersNames').val(), nDarkRounds: $('#nDarkRounds').val(), nGoldRounds: $('#nGoldRounds').val(), nDiamondRounds: $('#nDiamondRounds').val() }\
     );\
        return false;\
     });\
\
     $('#sidebarClose').on('click', function() {\
        google.script.host.close();\
     });\
     </script>";


    var htmlOutput = HtmlService
        .createHtmlOutput(sidebarHTML)
        .setTitle('Настройки игры');

    SpreadsheetApp.getUi().showSidebar(htmlOutput);

}


String.prototype.replaceAll = function(search, replace){
  return this.split(search).join(replace);
}

String.prototype.format = function() {
  a = this;
  for (k in arguments) {
    a = a.replaceAll("{" + k + "}", arguments[k])
  }
  return a
}


function setBorders(sheet, nPlayers, totalRounds) {
    sheet.getRange(3, 2, totalRounds, 4*nPlayers).setBorder(true, true, true, true, true, true, 'lightgray', SpreadsheetApp.BorderStyle.SOLID);
    for (var i = 0; i < nPlayers; i++) {
        sheet.getRange(1, 4*i + 2, totalRounds + 2, 4).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(1, 4*i + 2, 2, 4).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
}


function getAddress(x, y) {
    return "INDIRECT(ADDRESS({0};{1}))".format(x, y)
}


function getFormula(x, y, round, addPrevious=true) {
    var roundScores = {
        "main": 100,
        "dark": 100,
        "gold": 250,
        "diamond": 500
    }

    var first = getAddress(x, y - 2);
    var second = getAddress(x, y - 1);
    var mainScore = roundScores[round];
    var overflowScore = roundScores[round] / 10;
    var passScore = roundScores[round] / 2;
    var formula = "=IF(OR({0}=\"\";{1}=\"\");\"\";IF({0}={1};IF({0}=0;{4};{0}*{2});IF({0}<{1};{1}*{3};-{2}*({0}-{1}))))".format(first, second, mainScore, overflowScore, passScore);
    if (addPrevious) {
        formula = formula.slice(0, formula.length - 1) + "+" + getAddress(x - 1, y) + ")";
    }
    return formula;
}


function fillFormulas(sheet, nPlayers, nMainRounds, nDarkRounds, nGoldRounds, nDiamondRounds) {
    for (var i = 0; i < nMainRounds; i++) {
        for (var j = 0; j < nPlayers; j++) {
            var x = i + 3;
            var y = 4*j + 5;
            var addPrevious;
            if (i == 0) {
                addPrevious = false;
            } else {
                addPrevious = true;
            }
            sheet.getRange(x, y).setValue(getFormula(x, y, "main", addPrevious=addPrevious));
        }
    }
    for (var i = 0; i < nDarkRounds; i++) {
        for (var j = 0; j < nPlayers; j++) {
            var x = i + 3 + nMainRounds;
            var y = 4*j + 5;
            sheet.getRange(x, y).setValue(getFormula(x, y, "dark"));
        }
    }
    for (var i = 0; i < nGoldRounds; i++) {
        for (var j = 0; j < nPlayers; j++) {
            var x = i + 3 + nMainRounds + nDarkRounds;
            var y = 5*j + 6;
            sheet.getRange(x, y).setValue(getFormula(x, y, "gold"));
        }
    }
    for (var i = 0; i < nDiamondRounds; i++) {
        for (var j = 0; j < nPlayers; j++) {
            var x = i + 3 + nMainRounds + nDarkRounds + nGoldRounds;
            var y = 4*j + 5;
            sheet.getRange(x, y).setValue(getFormula(x, y, "diamond"));
        }
    }
}


function fillRoundNumbers(sheet, nMainRounds, nDarkRounds, nGoldRounds, nDiamondRounds) {
    for (var i = 0; i < nMainRounds; i++) {
        sheet.getRange(i + 3, 1).setValue(i + 1);
    }
    for (var i = 0; i < nDarkRounds; i++) {
        sheet.getRange(i + nMainRounds + 3, 1).setValue('Т');
    }
    for (var i = 0; i < nGoldRounds; i++) {
        sheet.getRange(i + nMainRounds + nDarkRounds + 3, 1).setValue('З');
    }
    for (var i = 0; i < nDiamondRounds; i++) {
        sheet.getRange(i + nMainRounds + nDarkRounds + nGoldRounds + 3, 1).setValue('А');
    }
}


function setColumnsWidths(sheet, nPlayers) {
    sheet.setColumnWidth(1, 25);
    for (var i = 0; i < nPlayers; i++) {
        sheet.setColumnWidth(4*i + 2, 25);
        sheet.setColumnWidth(4*i + 3, 25);
        sheet.setColumnWidth(4*i + 4, 25);
        sheet.setColumnWidth(4*i + 5, 80);
    }
}


function setRowsHights(sheet, totalRounds) {
    sheet.setRowHeight(1, 30);
    sheet.setRowHeights(2, totalRounds + 1, 23);
}


function fillPlayersNames(sheet, nPlayers, playersNames) {
    for (var i = 0; i < nPlayers; i++) {
        sheet.getRange(1, 4*i + 3).setValue(playersNames[i]);
    }
    sheet.getRange(1, 2, 2, 4*nPlayers).setBackground("lightcyan");
}


function fillPlayersScores(sheet, nPlayers, totalRounds) {
    for (var i = 0; i < nPlayers; i++) {
        var first = getAddress(3, 4*i + 5);
        var last = getAddress(totalRounds + 2, 4*i + 5);
        var curCell = getAddress(2, 4*i + 5);
        sheet.getRange(2, 4*i + 5).setValue("=OFFSET({0};COUNT({1}:{2});0)".format(curCell, first, last));
    }
}


function setCrowns(sheet, nPlayers) {
    for (var i = 0; i < nPlayers; i++) {
        var conditions = []
        for (var j = 0; j < nPlayers; j++) {
            if (i != j) {
                conditions.push("{0}>={1}".format(getAddress(2, 4*i + 5), getAddress(2, 4*j + 5)));
                conditions.push("{0}<>0".format(getAddress(2, 4*i + 5)));
            }
        }
        var stringCondition = conditions.join(";");
        var formula = "=IF(AND({0});\"👑\";\"\")".format(stringCondition);
        sheet.getRange(1, 4*i + 2).setValue(formula).setFontSize(14);
    }
}


function setAlignments(sheet, totalRounds, nPlayers) {
    sheet.getRange(1, 1, totalRounds + 2, 4*nPlayers + 1).setHorizontalAlignment("center");
    for (var i = 0; i < nPlayers; i++) {
        sheet.getRange(2, 4*i + 2).setHorizontalAlignment("left");
        sheet.getRange(1, 4*i + 5, totalRounds + 2, 1).setHorizontalAlignment("right");
        sheet.getRange(1, 4*i + 3).setHorizontalAlignment("left");
    }
}


function getDealer(sheet, totalRounds, nPlayers) {
    for (var i = 0; i < totalRounds; i++) {
        var conditions = [];
        for (var j = 0; j < nPlayers; j++) {
            conditions.push(getAddress(i + 3, 4*j + 5) + "=\"\"");
        }
        var stringCondition = conditions.join(";");
        sheet.getRange(i + 3, nPlayers*4 + 2).setValue("=IF(OR({0});0;1)".format(stringCondition));
    }
    sheet.hideColumn(sheet.getRange(i + 3, nPlayers*4 + 2));
    for (var i = 0; i < nPlayers; i++) {
        sheet.getRange(2, 4*i + 2).setValue("=IF(MOD(SUM({0}:{1});{2})={3};\"раздаёт\";\"\")".format(getAddress(3, 4*nPlayers + 2), getAddress(3 + totalRounds, 4*nPlayers + 2), nPlayers, i));
        sheet.getRange(2, 4*i + 2).setFontColor("red");
        sheet.getRange(2, 4*i + 2).setFontSize(7);
    }
}


function generateMessages(sheet, totalRounds, nPlayers, nMainRounds) {
    for (var i = 0; i < totalRounds; i++) {
        var orders = [];
        for (var j = 0; j < nPlayers; j++) {
            orders.push(getAddress(i + 3, 4*j + 3));
        }
        var stringOrders = orders.join(";");
        var facts = [];
        for (var j = 0; j < nPlayers; j++) {
            facts.push(getAddress(i + 3, 4*j + 4));
        }
        var stringFacts = facts.join(";");
        sheet.getRange(i + 3, 4*nPlayers + 3).setValue("=IF(AND({3}>={4};COUNT({0})={1});\"кроме \"&{2}-SUM({0});IF(AND(COUNT({5})={4};SUM({5})<>{2});\"количество взятых ставок не сходится\";\"\"))"
        .format(stringOrders, nPlayers - 1, Math.min(i + 1, nMainRounds), i, nPlayers, stringFacts));
    }
}


function roundResults(sheet, totalRounds, nPlayers) {
    for (var i = 0; i < totalRounds; i++) {
        for (var j = 0; j < nPlayers; j++) {
            var order = getAddress(i + 3, 4*j + 3);
            var fact = getAddress(i + 3, 4*j + 4);
            sheet.getRange(i + 3, 4*j + 2).setValue("=IF(OR({0}=\"\";{1}=\"\");\"\";IF({0}={1};\"+\";IF({0}>{1};\"-\";\"✔\")))"
            .format(order, fact));
        }
    }
}


function writeStrInTable(input) {
    var playersNames = input.playersNames.split('\n');
    var nPlayers = playersNames.length;
    var nMainRounds = Math.min(13, Math.floor(54 / nPlayers));
    var nDarkRounds = parseInt(input.nDarkRounds);
    var nGoldRounds = parseInt(input.nGoldRounds);
    var nDiamondRounds = parseInt(input.nDiamondRounds);
    var totalRounds = nMainRounds + nDarkRounds + nGoldRounds + nDiamondRounds;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    setColumnsWidths(sheet, nPlayers);
    setRowsHights(sheet, totalRounds);
    setBorders(sheet, nPlayers, totalRounds);
    sheet.setHiddenGridlines(true);
    setAlignments(sheet, totalRounds, nPlayers);
    fillPlayersNames(sheet, nPlayers, playersNames);
    fillRoundNumbers(sheet, nMainRounds, nDarkRounds, nGoldRounds, nDiamondRounds);
    fillFormulas(sheet, nPlayers, nMainRounds, nDarkRounds, nGoldRounds, nDiamondRounds);
    fillPlayersScores(sheet, nPlayers, totalRounds);
    setCrowns(sheet, nPlayers);
    getDealer(sheet, totalRounds, nPlayers);
    roundResults(sheet, totalRounds, nPlayers);
    generateMessages(sheet, totalRounds, nPlayers, nMainRounds);
}
