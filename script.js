/**
**  Disclaimer: XLSX Code comes from http://oss.sheetjs.com/js-xlsx/
**  This is just a project to show Vue.js working in conjunction with HTML5 Drag and Drop and XLSX
**/

/** DATA STATE **/
var state={
    tickets:[{name:"test"}],
    headers:["Test header"]
  }
/** HELPERS **/
function get_header_row(sheet) {
    var headers = [], range = XLSX.utils.decode_range(sheet['!ref']);
    var C, R = range.s.r; /* start in the first row */
    for(C = range.s.c; C <= range.e.c; ++C) { /* walk every column in the range */
        var cell = sheet[XLSX.utils.encode_cell({c:C, r:R})] /* find the cell in the first row */
        var hdr = "UNKNOWN " + C; // <-- replace with your desired default 
        if(cell && cell.t) hdr = XLSX.utils.format_cell(cell);
        headers.push(hdr);
    }
    return headers;
}
function fixdata(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
	o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
	return o;
}
function workbook_to_json(workbook) {
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if(roa.length > 0){
			result[sheetName] = roa;
		}
	});
	return result;
}
/** PARSING and DRAGDROP **/
function handleDrop(e) {
  e.stopPropagation();
  e.preventDefault();
  console.log("DROPPED");
  var files = e.dataTransfer.files, i, f;
  for (i = 0, f = files[i]; i != files.length; ++i) {
    var reader = new FileReader(),
        name = f.name;
    reader.onload = function(e) {
      var results, 
          data = e.target.result, 
          fixedData = fixdata(data), 
          workbook=XLSX.read(btoa(fixedData), {type: 'base64'}), 
          firstSheetName = workbook.SheetNames[0], 
          worksheet = workbook.Sheets[firstSheetName];
      state.headers=get_header_row(worksheet);
      results=XLSX.utils.sheet_to_json(worksheet);
      state.tickets=results;
    };
    reader.readAsArrayBuffer(f);
  }
}
function handleDragover(e) {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}
 /*Testing only*
var drop=document.getElementById("drop");
if(drop.addEventListener) {
	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
}*/

/** VIEW **/
var myView=new Vue({
  el:"#app",
  data:state,
  methods:{
  handleDragover:handleDragover,
  handleDrop:handleDrop
}
});