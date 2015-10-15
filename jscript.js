// JavaScript File

function search(username, password) {
    if (readdata(username, password)) {
        return true;
    }
    else {
        return false;
    }
}

//Taken from http://stackoverflow.com/questions/16630413/how-can-i-read-an-excel-file-with-javascript-without-activexobject
function readdata(x,y) {
    var xval = x;
    var yval = y;
    var yon = false;
    try {
        var excel = new ActiveXObject("Excel.Application");
        excel.Visible = false;
        var excel_file = excel.Workbooks.Open("\\Accountinfo.xlsx");// alert(excel_file.worksheets.count);
        var excel_sheet = excel_file.Worksheets("Sheet1");
        
        for(var i=0;i<100000;i++)
        {
           var temp = excel_sheet.Cells(1,i).Value;
           
            if (temp == null) {
                break
            }
            else {
                
                var temp2 = excel_sheet.Cells(2,i).Value;
                if (x == temp && y == temp2){
                    yon = true;
                    break;
                }
            }
        }
    return yon;
    }
    catch (ex) {
        alert(ex);
    }
}