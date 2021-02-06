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

    function refreshData() {
        Excel.run(function (context) {            
            data = pollEndpoint('https://jsonplaceholder.typicode.com/todos/1')
                .then(
                    function (data) {
                        let sheet = context.workbook.worksheets.getActiveWorksheet();
                        let range = sheet.getRange("A1");
                        let text = range.convertDataTypeToText();
                        if (text === undefined) {
                            range.values = JSON.stringify(data);
                            context.sync();
                        }else {
                            throw 'Values already present in range'
                        }
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