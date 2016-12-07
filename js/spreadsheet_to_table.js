var SpreadsheetToHtmlTable = SpreadsheetToHtmlTable || {};

SpreadsheetToHtmlTable = {
    init: function (options) {

      options = options || {};
      var datatables_options = options.datatables_options || {};
      var csv_options = {separator: ',', delimiter: '"'};
      url = 'data/'+options.filepath;

      function to_csv(workbook) {
        var result = [];
        workbook.SheetNames.forEach(function(sheetName) {
          var csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
          if(csv.length > 0){
            result.push({"name":sheetName,"data":csv})
          }
        });
        return result;
      }


      function createTable(output){
            
            html = '';
            nav = '';
            tableaux = '';
            $.each(output,function(i,d){
              nav +=  '<li role="presentation"'+(i==0 ? ' class="active"' : '')+'><a href="#tab_'+i+'" aria-controls="home" role="tab" data-toggle="tab">'+d.name+'</a></li>';
              tableaux += '<div role="tabpanel" class="tab-pane '+(i==0 ? ' active' : '')+'" id="tab_'+i+'"><table class="preview-container-table"></table></div>'
            })
            html += '<ul class="nav nav-tabs">'+nav+'</ul><div class="tab-content">'+tableaux+'</div>';
            $("#preview-container").html(html)

            $.each(output,function(i,d){ /* Création des tableaux */

                  var csv_data = $.csv.toArrays(d.data, csv_options);
                
                  /* On prépare les en-têtes de colonnes */
                  colonnes = [];
                  $.each(csv_data[0],function(i,d){
                    colonnes.push({ "title": d })
                  })

                  csv_data.shift() /* On enlève l'en-tête de colonne des données */

                  datatables_options["data"] = csv_data;
                  datatables_options["columns"] = colonnes;
                  $('#tab_'+i+' .preview-container-table').DataTable( datatables_options);

            })
             
             $("#download_link").html("<p><a class='btn btn-info' href='data/" + filepath + "'><i class='glyphicon glyphicon-download'></i> Télécharger le fichier original</a></p>");

      }

      function process_wb(wb) {
        var output = to_csv(wb);
        createTable(output)
      }


     var oReq;
      if(window.XMLHttpRequest) oReq = new XMLHttpRequest();
      else if(window.ActiveXObject) oReq = new ActiveXObject('MSXML2.XMLHTTP.3.0');
      else throw "XHR unavailable for your browser";
      oReq.open("GET", url, true);

      if(typeof Uint8Array !== 'undefined') {
        oReq.responseType = "arraybuffer";
        oReq.onload = function(e) {
          if(typeof console !== 'undefined') console.log("onload", new Date());
          var arraybuffer = oReq.response;
          var data = new Uint8Array(arraybuffer);
          var arr = new Array();
          for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
          var wb = XLSX.read(arr.join(""), {type:"binary"});
          process_wb(wb);
        };
      } else {
        oReq.setRequestHeader("Accept-Charset", "x-user-defined");  
        oReq.onreadystatechange = function() { if(oReq.readyState == 4 && oReq.status == 200) {
          var ff = convertResponseBodyToText(oReq.responseBody);
          if(typeof console !== 'undefined') console.log("onload", new Date());
          var wb = XLSX.read(ff, {type:"binary"});
          process_wb(wb);
        } };
      }

      oReq.send();
      



    }
}