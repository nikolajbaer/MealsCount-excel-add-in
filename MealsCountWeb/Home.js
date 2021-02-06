'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            $('#refresh').click(refreshData);
        });
    });

    async function pollEndpoint(url) {
        let response = await fetch(url);
        let data = await response.json();
        return data;
    }

    async function checkRange(rangeStr, url) {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange(rangeStr);
        if (range.values.data === undefined) {
            let data = await pollEndpoint(url);
            return data;
        } else {
            throw 'Suggested range is already populated'
        }
    }

    function refreshData() {
        Excel.run(function (context) {            
            data = checkRange('https://jsonplaceholder.typicode.com/todos/1')
                .then(
                    function (data) {
                        let sheet = context.workbook.worksheets.getActiveWorksheet();
                        let range = sheet.getRange("A1");
                        range.values = JSON.stringify(data);
                        context.sync();
                    },
                    function (err) {
                        console.log(err);
                    }
                );            
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
})();