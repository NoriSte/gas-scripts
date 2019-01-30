
// generic sheet data
var startRow = 4;
var endRow = 260;

// colums data
var cashColumn = 'C';
var bankAccountColumn = 'D';
var column3 = 'E';
var column4 = 'F';
var column5 = 'G';
var column6 = 'H';
var column7 = 'I';
var column8 = 'J';
var column9 = 'K';
var column10 = 'L';
var lastColumn = column10;

// savings data
var savingsRow = '22';

////////////////////////////////////////////////////////////////////////

var _getSheet = function getSheet(){
  return SpreadsheetApp.getActive().getActiveSheet();
}

/**
 * Sposta i soldi di tutte le colonne usando come pivot il conto comune e azzerando ogni budget presente precedentemente
 */

function balanceMoney() {
  if(PropertiesService.getScriptProperties().getProperty('running') === 'true') {
    throw new Error('The script is already running');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('running', 'true');
  
  var oldBalances = JSON.stringify(_getBalances());
  Logger.log('Old balances: ' + oldBalances)
  _balanceMoneyWithBanckAccount(cashColumn);
  _balanceMoneyWithBanckAccount(column3);
  _balanceMoneyWithBanckAccount(column4);
  _balanceMoneyWithBanckAccount(column5);
  _balanceMoneyWithBanckAccount(column6);
  _balanceMoneyWithBanckAccount(column7);
  _balanceMoneyWithBanckAccount(column8);
  _balanceMoneyWithBanckAccount(column9);
  _balanceMoneyWithBanckAccount(column10);
  var newBalances = JSON.stringify(_getBalances());
  Logger.log('New balances: ' + newBalances)
  
  PropertiesService.getScriptProperties().setProperty('running', 'false');
  
  if(oldBalances != newBalances) {
    throw new Error('The balances don\'t match, check them manually... (from '+oldBalances+' to '+newBalances+')');
  }
}

function _getBalances() {
  var sheet = _getSheet();
  var range = sheet.getRange(cashColumn+'1'+':'+lastColumn+'1');
  return range.getValues();
}

/**
 * Prendendo una colonna di riferimento, sposta i soldi usando come pivot il conto comune e azzerando ogni budget presente precedentemente nella colonna stessa
 */
function _balanceMoneyWithBanckAccount(column) {
  var sheet = _getSheet();
  var rangeToBalance = sheet.getRange(column+startRow+':'+column+endRow);
  var cellsToBalance = rangeToBalance.getValues();
  var backAccountRange = sheet.getRange(bankAccountColumn+startRow+':'+bankAccountColumn+endRow);
  var backAccountCells = backAccountRange.getValues();
  var i, n = cellsToBalance.length;
  var amount = 0;
  for(i = 0; i < n; i++) {
    var valueToBalance = cellsToBalance[i][0];
    var bankAccountValue = backAccountCells[i][0];
    if(valueToBalance) {
      amount += valueToBalance;
      backAccountCells[i][0] = backAccountCells[i][0] + valueToBalance;
      cellsToBalance[i][0] = '';
      Logger.log('moved ' + valueToBalance + ' € from bank account ');
    }
  }
  
  backAccountRange.setValues(backAccountCells);
  rangeToBalance.setValues(cellsToBalance);
  
  Logger.log('totally moved ' + amount + ' € from bank account ');
  
  const savingsBanckAccountRange = sheet.getRange('C'+startRow+':C'+endRow)
  var savingsToBalanceCell = sheet.getRange(column+savingsRow);
  var savingsBankAccountCell = sheet.getRange(bankAccountColumn+savingsRow);
  savingsToBalanceCell.setValue(amount);
  savingsBankAccountCell.setValue(savingsBankAccountCell.getValue() - amount);
};