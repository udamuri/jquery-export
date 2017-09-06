/*
 * jQuery Client Side Export Plugin Library
 * 
 */

(function ($) {
    var $defaults = {
        containerid: null
        , exporttype: 'excel'
        , jsondata: null
        , columns: null
        , returnUri: false
        , worksheetName: "mySheet"
        , encoding: "utf-8"
        , showLabel: true
    };

    var $settings = $defaults;

    $.fn.jqueryexport = function (options) {

        $settings = $.extend({}, $defaults, options);

        var gridData = $settings.jsondata;
        var excelData;

        return Initialize();
		
		function Initialize() {
            var type = $settings.exporttype.toLowerCase();

            switch (type) {
                case 'excel':
                    excelData = Export(ConvertDataStructureToTable());
                    break;
                case 'csv':
                    excelData = JSONToCSVConvertor(gridData, $settings.worksheetName , $settings.showLabel);
                    break;
                case 'pdf':
                    excelData = ExportPdf(ConvertDataStructureToTable());
                    break;
            }

       
            if ($settings.returnUri) {
                return excelData;
            }
            else {

                if (!isBrowserIE())
                {
                    window.open(excelData);
                }

               
            }
        }

        function ConvertDataStructureToTable() {
            var result = "<table id='tabledata'>";

            result += "<thead><tr>";
            $($settings.columns).each(function (key, value) {
                if (this.ishidden != true) {
                    result += "<th";
                    if (this.width != null) {
                        result += " style='width: " + this.width + "'";
                    }
                    result += ">";
                    result += this.headertext;
                    result += "</th>";
                }
            });
            result += "</tr></thead>";

            result += "<tbody>";
            $(gridData).each(function (key, value) {
                result += "<tr>";
                $($settings.columns).each(function (k, v) {
                    if (value.hasOwnProperty(this.datafield)) {
                        if (this.ishidden != true) {
                            result += "<td";
                            if (this.width != null) {
                                result += " style='width: " + this.width + "'";
                            }
                            result += ">";
                            result += value[this.datafield];
                            result += "</td>";
                        }
                    }
                });
                result += "</tr>";
            });
            result += "</tbody>";

            result += "</table>";

            return result;
        }

        function Export(htmltable) {

            if (isBrowserIE()) {
        
                exportToExcelIE(htmltable);
            }
            else {
                var excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
                excelFile += "<head>";
                excelFile += '<meta http-equiv="Content-type" content="text/html;charset=' + $defaults.encoding + '" />';
                excelFile += "<!--[if gte mso 9]>";
                excelFile += "<xml>";
                excelFile += "<x:ExcelWorkbook>";
                excelFile += "<x:ExcelWorksheets>";
                excelFile += "<x:ExcelWorksheet>";
                excelFile += "<x:Name>";
                excelFile += "{worksheet}";
                excelFile += "</x:Name>";
                excelFile += "<x:WorksheetOptions>";
                excelFile += "<x:DisplayGridlines/>";
                excelFile += "</x:WorksheetOptions>";
                excelFile += "</x:ExcelWorksheet>";
                excelFile += "</x:ExcelWorksheets>";
                excelFile += "</x:ExcelWorkbook>";
                excelFile += "</xml>";
                excelFile += "<![endif]-->";
                excelFile += "</head>";
                excelFile += "<body>";
                excelFile += htmltable.replace(/"/g, '\'');
                excelFile += "</body>";
                excelFile += "</html>";

                var uri = "data:application/vnd.ms-excel;base64,";
                var ctx = { worksheet: $settings.worksheetName, table: htmltable };

                return (uri + base64(format(excelFile, ctx)));
            }
        }

        function ExportPdf(htmltable) {
            var uri = "data:application/pdf;base64,";
            return (uri + base64(htmltable));
        }

        function JSONToCSVConvertor(JSONData, ReportTitle, ShowLabel) {
              //If JSONData is not an object then JSON.parse will parse the JSON string in an Object
              var arrData = typeof JSONData != 'object' ? JSON.parse(JSONData) : JSONData;
              
              var CSV = '';    
              //Set Report title in first row or line
              
              //CSV += ReportTitle + '\r\n\n';

              //This condition will generate the Label/Header
              if (ShowLabel) {
                  var row = "";
                  
                  //This loop will extract the label from 1st index of on array
                  for (var index in arrData[0]) {
                      
                      //Now convert each value to string and comma-seprated
                      row += index + ',';
                  }

                  row = row.slice(0, -1);
                  
                  //append Label row with line break
                  CSV += row + '\r\n';
              }
              
              //1st loop is to extract each row
              for (var i = 0; i < arrData.length; i++) {
                  var row = "";
                  
                  //2nd loop will extract each column and convert it in string comma-seprated
                  for (var index in arrData[i]) {
                      row += '"' + arrData[i][index] + '",';
                  }

                  row.slice(0, row.length - 1);
                  
                  //add a line break after each row
                  CSV += row + '\r\n';
              }

              if (CSV == '') {        
                  alert("Invalid data");
                  return;
              }   
              
              //Generate a file name
              var fileName = "report_";
              //this will remove the blank-spaces from the title and replace it with an underscore
              fileName += ReportTitle.replace(/ /g,"_");   
              
              //Initialize file format you want csv or xls
              var uri = 'data:text/csv;charset=utf-8,' + escape(CSV);
              
              // Now the little tricky part.
              // you can use either>> window.open(uri);
              // but this will not work in some browsers
              // or you will not get the correct file extension    
              
              //this trick will generate a temp <a /> tag
              var link = document.createElement("a");    
              link.href = uri;
              
              //set the visibility hidden so it will not effect on your web-layout
              link.style = "visibility:hidden";
              link.download = fileName + ".csv";
              
              //this part will append the anchor tag and remove it after automatic click
              document.body.appendChild(link);
              link.click();
              document.body.removeChild(link);
        }

        function isBrowserIE() {
            var msie = !!navigator.userAgent.match(/Trident/g) || !!navigator.userAgent.match(/MSIE/g);
            if (msie > 0) {  // If Internet Explorer, return true
                return true;
            }
            else {  // If another browser, return false
                return false;
            }
        }

        function base64(s) {
            return window.btoa(unescape(encodeURIComponent(s)));
        }

        function format(s, c) {
            return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; });
        }

        function exportToExcelIE(table) {


            var el = document.createElement('div');
            el.innerHTML = table;

            var tab_text = "<table border='2px'><tr bgcolor='#87AFC6'>";
            var textRange; var j = 0;
            var tab;
                  

            if ($settings.datatype.toLowerCase() == 'table') {            
                tab = document.getElementById($settings.containerid);  // get table              
            }
            else{
                tab = el.children[0]; // get table
            }

          
        
            for (j = 0 ; j < tab.rows.length ; j++) {
                tab_text = tab_text + tab.rows[j].innerHTML + "</tr>";
                //tab_text=tab_text+"</tr>";
            }

            tab_text = tab_text + "</table>";
            tab_text = tab_text.replace(/<A[^>]*>|<\/A>/g, "");//remove if u want links in your table
            tab_text = tab_text.replace(/<img[^>]*>/gi, ""); // remove if u want images in your table
            tab_text = tab_text.replace(/<input[^>]*>|<\/input>/gi, ""); // reomves input params

            var ua = window.navigator.userAgent;
            var msie = ua.indexOf("MSIE ");

            if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // If Internet Explorer
            {
                txtArea1.document.open("txt/html", "replace");
                txtArea1.document.write(tab_text);
                txtArea1.document.close();
                txtArea1.focus();
                sa = txtArea1.document.execCommand("SaveAs", true, "download");
            }
            else                
                sa = window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));

            return (sa);


        }

        
        
    };
})(jQuery);

//get columns
function getColumns(paramData){

    var header = [];
    $.each(paramData[0], function (key, value) {
        //console.log(key + '==' + value);
        var obj = {}
        obj["headertext"] = key;
        obj["datatype"] = "string";
        obj["datafield"] = key;
        header.push(obj);
    }); 
    return header;

}
