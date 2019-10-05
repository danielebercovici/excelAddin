'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            $('#clean_names').click(cleanFirstName);
        });
    });

    function cleanFirstName() {
        Excel.run(function (context) {
            var sheet = context.workbook.worksheets.getItem("Sheet1");
            var range = sheet.getRange("G1:G3");
            // range.values = [["hello"], ["poop"], ["test"]];
            // range.format.autofitColumns();

            range.load("values");
            var copyValues = JSON.parse(JSON.stringify(range.values));
            var l = range.values.length;

            for (let i = 0; i < l; i++) {
                var oldString = copyValues[i][0];
                var newString = oldString.replace(/^./, oldString[0].toUpperCase());
                copyValues[i][0] = newString; 
            }
            range.values = copyValues;
            range.format.autofitColumns();

            return context.sync()
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
})();