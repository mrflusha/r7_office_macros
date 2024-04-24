(function()
{
    var sheets = Api.GetSheets();

    let arr = []
   
    for(var i = 0; i < sheets.length; i++){
      
        arr.push(sheets[i].GetName().replaceAll(' ','_'));
        sheets[i].SetName(arr[i])
    }
    
})();