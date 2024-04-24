(function()
{
    
    var sheets = Api.GetSheets();
    var oWorksheet = Api.GetActiveSheet();
    var sAddress = oWorksheet.GetRange("A1").GetAddress(true, true, "xlA1", false);
    
    Api.GetRange("a7").SetValue(sAddress)
    
    for (var i = 0; i < sheets.length;i++){
        sheets[0].SetHyperlink(`a+${i+1}`,"", `${sheets[i].GetName()}!a1`, `${sheets[i].GetName()}`, "")
    }
})();