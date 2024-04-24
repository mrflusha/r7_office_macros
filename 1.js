var sheets = Api.GetSheets(); 

    for (var i = 0; i < sheets.length; i++){
        Api.GetActiveSheet().GetRange("C"+i.toString()).SetValue("Hello world");
        sheets[i].SetVisible(true);
        

    }