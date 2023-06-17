const XLSX = require("xlsx");

(async () => {

    const workbook = XLSX.readFile('parsing.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_row_object_array(worksheet);

    rows.forEach(row => {
        for (const key in row) {
            if (Object.prototype.hasOwnProperty.call(row, key) && typeof row[key] === 'number' && !isNaN(row[key])) {
                if (key === 'date') {
                    var date = XLSX.SSF.parse_date_code(row[key]);
                    var val = new Date();
                    if (date == null) throw new Error("Bad Date Code: " + v);
                    val.setUTCDate(date.d);
                    val.setUTCMonth(date.m - 1);
                    val.setUTCFullYear(date.y);
                    val.setUTCHours(date.H);
                    val.setUTCMinutes(date.M);
                    val.setUTCSeconds(date.S);
                    row[key] = val;
                }else if(key !== 'srid'){
                    row[key] = row[key].toString();
                }
            }
        }
    });

    // console.log(rows);
    var json_object = JSON.stringify(rows);
    console.log(json_object);

})();
