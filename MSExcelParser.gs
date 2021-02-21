/**
* Parsing MS Excel files and returns values in JSON format.
*
* @param {BlobSource} blob the blob from MS Excel file
* @param {String[]} requiredSheets the array of required sheet names (if omitted returns all)
* @return {Object} Object of sheet names and values (2D arrays)
*/
function parseMSExcelBlob(blob, requiredSheets){
    var col_cache = {};
    var forbidden_chars = {
        "&lt;": "<",
        "&gt;": ">",
        "&amp;": "&",
        "&apos;": "'",
        "&quot;": '"'
    };
    
    blob.setContentType("application/zip");
    var parts = Utilities.unzip(blob);
    
    var relationships = {};
    for( var part of parts ){
        var part_name = part.getName();
        if( part_name === "xl/_rels/workbook.xml.rels" ){
            var txt = part.getDataAsString();
            var rels = breakUpString(txt, '<Relationship ', '/>');
            for( var i = 0; i < rels.length; i++ ){
                var rId = breakUpString(rels[i], 'Id="', '"')[0];
                var path = breakUpString(rels[i], 'Target="', '"')[0];
                relationships[rId] = "xl/" + path;
            }
        }
    }
    
    var worksheets = {};
    for( var part of parts ){
        var part_name = part.getName();
        if( part_name === "xl/workbook.xml" ){
            var txt = part.getDataAsString();
            var sheets = breakUpString(txt, '<sheet ', '/>');
            for( var i = 0; i < sheets.length; i++ ){
                var sh_name = breakUpString(sheets[i], 'name="', '"')[0];
                sh_name = decodeForbiddenChars(sh_name);
                var rId = breakUpString(sheets[i], 'r:id="', '"')[0];
                var path = relationships[rId];
                if( path.includes("worksheets") ){
                    worksheets[path] = sh_name;
                }
            }
        }
    }
    
    requiredSheets = Array.isArray(requiredSheets) && requiredSheets.length && requiredSheets || [];
    var worksheets_needed = [];
    for( var path in worksheets ){
        if( !requiredSheets.length || requiredSheets.includes(worksheets[path]) ){
            worksheets_needed.push(path);
        }
    }
    if( !worksheets_needed.length ) return {"Error": "Requested worksheets not found"};
    
    var sharedStrings = [];
    for( var part of parts ){
        var part_name = part.getName();
        if( part_name === "xl/sharedStrings.xml" ){
            var txt = part.getDataAsString();
            txt = txt.replace(/ xml:space="preserve"/g, "");
            sharedStrings = breakUpString(txt, '<si>', '</si>');
            for( var i = 0; i < sharedStrings.length; i++ ){
                var str = breakUpString(sharedStrings[i], '<t>', '</t>')[0];
                sharedStrings[i] = decodeForbiddenChars(sharedStrings[i]);
            }
        }
    }
    
    var result = {};
    for( var part of parts ){
        var part_name = part.getName();
        if( worksheets_needed.includes(part_name) ){
            var txt = part.getDataAsString();
            txt = txt.replace(/ xml:space="preserve"/g, "");
            var cells = breakUpString(txt, '<c ', '</c>');
            var tbl = [[]];
            for( var i = 0; i < cells.length; i++ ){
                var r = breakUpString(cells[i], 'r="', '"')[0];
                var t = breakUpString(cells[i], 't="', '"')[0];
                if( t === "inlineStr" ){
                    var data = breakUpString(cells[i], '<t>', '</t>')[0];
                    data = decodeForbiddenChars(data);
                }else if( t === "s" ){
                    var v = breakUpString(cells[i], '<v>', '</v>')[0];
                    var data = sharedStrings[v];
                }else{
                    var v = breakUpString(cells[i], '<v>', '</v>')[0];
                    var data = Number(v);
                }
                var row = r.replace(/[A-Z]/g, "") - 1;
                var col = colNum(r.replace(/[0-9]/g, "")) - 1;
                if( tbl[row] ){
                    tbl[row][col] = data;
                }else{
                    tbl[row] = [];
                    tbl[row][col] = data;
                }
            }
            var sh_name = worksheets[part_name];
            result[sh_name] = squareTbl(tbl);
        }
    }
    
    
    function decodeForbiddenChars(txt){
        if( !txt ) return txt;
        for( var char in forbidden_chars ){
            var regex = new RegExp(char,"g");
            txt = txt.replace(regex, forbidden_chars[char]);
        }
        return txt;
    }
    
    function breakUpString(str, start_patern, end_patern){
        var arr = [], raw = str.split(start_patern), i = 1, len = raw.length;
        while( i < len ){ arr[i - 1] = raw[i].split(end_patern, 1)[0]; i++ };
        return arr;
    }
    
    function colNum(char){
        if( col_cache[char] ) return col_cache[char];
        var alph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ", i, j, result = 0;
        for( i = 0, j = char.length - 1; i < char.length; i++, j-- ){
            result += Math.pow(alph.length, j) * (alph.indexOf(char[i]) + 1);
        }
        col_cache[char] = result;
        return result;
    }
    
    function squareTbl(arr){
        var tbl = [];
        var x_max = 0;
        var y_max = arr.length;
        for( var y = 0; y < y_max; y++ ){
            arr[y] = arr[y] || [];
            if( arr[y].length > x_max ){ x_max = arr[y].length };
        }
        for( var y = 0; y < y_max; y++ ){
            var row = [];
            for( var x = 0; x < x_max; x++ ){
                row.push(arr[y][x] || arr[y][x] === 0 ? arr[y][x] : "");
            }
            tbl.push(row);
        }
        return tbl.length ? tbl : [[]];
    }
    
    
    return result;
}


