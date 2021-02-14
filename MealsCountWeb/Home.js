'use strict';

(function () {

    Office.onReady(function () {
        $(document).ready(function () {
            createWorksheet();
            $('#refresh').click(refreshData);
        });
    });

    // Add a worksheet to store results within the current workbook
    function createWorksheet() {
        Excel.run(function (context) {
            var sheets = context.workbook.worksheets;
            sheets.add("Meals Count");

            return context.sync()
                .then(
                    () => console.log('Added worksheet'),
                    (err) => console.error(err)
                );
        });
    }

    // React to user's button click
    function refreshData() {

        // Fetch school info for a given state / district 
        async function getSchoolInfo(url) {
            // Get data from MC API
            let data = await fetch(url)
                .then((resp) => resp.json())
                .then((data) => enrichSchoolInfo(data))
                .catch((err) => console.error(err));
            return data;
        }

        // Compute new values and add them to school info object
        function enrichSchoolInfo(schoolInfo) {
            //Validate inputs

            // Translate mcJson into array that can be used to populate worksheet

            // For testing ignore the input and parse a hardcoded sample
            //schoolInfo = { "name": "Queen Creek Unified District", "code": "70295000", "total_enrolled": 7627, "overall_isp": 0.07657007997902189, "school_count": 9, "best_strategy": null, "est_reimbursement": 0.0, "rates": { "free_lunch": 3.41, "paid_lunch": 0.32, "free_bfast": 1.84, "paid_bfast": 0.31 }, "schools": [{ "school_code": "70295102", "school_name": "Desert Mountain Elementary", "school_type": "n/a", "total_enrolled": 657, "total_eligible": 51, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0776, "active": true }, { "school_code": "70295107", "school_name": "Faith Mather Sossaman Elementary School", "school_type": "n/a", "total_enrolled": 660, "total_eligible": 50, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0758, "active": true }, { "school_code": "70295104", "school_name": "Frances Brandon-Pickett Elementary", "school_type": "n/a", "total_enrolled": 548, "total_eligible": 42, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0766, "active": true }, { "school_code": "70295105", "school_name": "Gateway Polytechnic Academy", "school_type": "n/a", "total_enrolled": 853, "total_eligible": 56, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0657, "active": true }, { "school_code": "70295103", "school_name": "Jack Barnes Elementary School", "school_type": "n/a", "total_enrolled": 377, "total_eligible": 31, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0822, "active": true }, { "school_code": "70295121", "school_name": "Newell Barney Middle School", "school_type": "n/a", "total_enrolled": 877, "total_eligible": 67, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0764, "active": true }, { "school_code": "70295101", "school_name": "Queen Creek Elementary School", "school_type": "n/a", "total_enrolled": 620, "total_eligible": 65, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.1048, "active": true }, { "school_code": "70295201", "school_name": "Queen Creek High School", "school_type": "n/a", "total_enrolled": 2158, "total_eligible": 166, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0769, "active": true }, { "school_code": "70295106", "school_name": "Queen Creek Middle School", "school_type": "n/a", "total_enrolled": 877, "total_eligible": 56, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0639, "active": true }] };
            for (let school of schoolInfo.schools){
                school['daily_breakfast_served'] = parseInt(school['total_enrolled'] * 0.15);
                school['daily_lunch_served'] = parseInt(school['total_enrolled'] * 0.55);
            }
            return schoolInfo;
        }

        // Format the solution so that it can be pasted into the Meals Count Excel Worksheet
        function formatOptimalSolution(solution) {
            return JSON.stringify(solution);
        }

        // Post and poll for the optimal solution for the given district using the enriched school info object
        async function optimize(schoolInfo) {
            let results = await new Promise((resolve, reject) => {
                // Get URL to poll for results
                let requestObj = {
                    method: 'POST',
                    body: JSON.stringify(schoolInfo)
                };
                let solutionUrl = await fetch('https://www.mealscount.com/api/districts/optimize-async/', requestObj)              
                    .then((resp) => resp.json())
                    .then((data) => data.results_url)
                    .catch((err) => reject(err));

                // Poll every second until status code 200 and results are present then kill poller and return results
                let getResults = async function () {

                    let resp = await fetch(solutionUrl)
                        .then((resp) => resp)
                        .catch((err) => { clearInterval(poller); reject(err); });

                    if (resp.status_code == 200) {
                        clearInterval(poller);
                        resp.json()
                            .then((data) => resolve(data))
                            .catch((err) => reject(err));
                    }
                };

                // Poll every second
                let poller = setInterval(getResults, 1000);
            });

            return results;    
        }

        
        // Get values from user
        let state = document.getElementById("state").value;
        let district = document.getElementById("district").value;
        let url = 'https://www.mealscount.com/static/' + state + '/' + district + '_district.json';

        //let url = 'https://jsonplaceholder.typicode.com/todos/1';
        //let url = 'https://www.mealscount.com/static/az/70295000_district.json';

        //Show working indicator
        let working = document.getElementById("working");
        working.style.display = "block";

        // Fetch data and populate newly created worksheet
        getSchoolInfo(url)
            .then((enrichedData) => {
                optimize(enrichedData)
                    .then((optimalSolution) => {
                        let formattedSolution = formatOptimalSolution(optimalSolution);
                        // When promise is resolved, update Excel
                        Excel.run((context) => {
                            // Get Excel objects
                            let sheet = context.workbook.worksheets.getItem("Meals Count");
                            let range = sheet.getRange("A1");

                            // Define what data should go in proxy objects
                            range.values = formattedSolution;

                            // Actually put that data there
                            return context.sync().then(function () {
                                console.log('Populated workbook')
                            });
                        }).catch((err) => console.error(err));
                    }).catch((err) => console.error(err));
            }).catch((err) => console.error(err));                    
    }
})();