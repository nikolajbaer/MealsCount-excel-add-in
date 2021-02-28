'use strict';

(function () {
    Office.onReady(function () {
        $(document).ready(function () {
            createWorksheet();
            //$('#refresh').click(refreshData);
        });
    });

    // Add a worksheet to store results within the current workbook
    function createWorksheet() {
        Excel.run(context => {
            var sheets = context.workbook.worksheets;
            sheets.add("Meals Count");

            return context.sync()
                .then(
                    () => console.log('Added worksheet'),
                    (err) => console.error(err)
                );
        });
    }
})()

// React to user's button click
function refreshData() {

    // Compute new values and add them to school info object
    function enrichSchoolInfo(schoolInfo, state) {
        // Validate input?

        // Example
        //schoolInfo = { "name": "Queen Creek Unified District", "code": "70295000", "total_enrolled": 7627, "overall_isp": 0.07657007997902189, "school_count": 9, "best_strategy": null, "est_reimbursement": 0.0, "rates": { "free_lunch": 3.41, "paid_lunch": 0.32, "free_bfast": 1.84, "paid_bfast": 0.31 }, "schools": [{ "school_code": "70295102", "school_name": "Desert Mountain Elementary", "school_type": "n/a", "total_enrolled": 657, "total_eligible": 51, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0776, "active": true }, { "school_code": "70295107", "school_name": "Faith Mather Sossaman Elementary School", "school_type": "n/a", "total_enrolled": 660, "total_eligible": 50, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0758, "active": true }, { "school_code": "70295104", "school_name": "Frances Brandon-Pickett Elementary", "school_type": "n/a", "total_enrolled": 548, "total_eligible": 42, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0766, "active": true }, { "school_code": "70295105", "school_name": "Gateway Polytechnic Academy", "school_type": "n/a", "total_enrolled": 853, "total_eligible": 56, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0657, "active": true }, { "school_code": "70295103", "school_name": "Jack Barnes Elementary School", "school_type": "n/a", "total_enrolled": 377, "total_eligible": 31, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0822, "active": true }, { "school_code": "70295121", "school_name": "Newell Barney Middle School", "school_type": "n/a", "total_enrolled": 877, "total_eligible": 67, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0764, "active": true }, { "school_code": "70295101", "school_name": "Queen Creek Elementary School", "school_type": "n/a", "total_enrolled": 620, "total_eligible": 65, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.1048, "active": true }, { "school_code": "70295201", "school_name": "Queen Creek High School", "school_type": "n/a", "total_enrolled": 2158, "total_eligible": 166, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0769, "active": true }, { "school_code": "70295106", "school_name": "Queen Creek Middle School", "school_type": "n/a", "total_enrolled": 877, "total_eligible": 56, "daily_breakfast_served": null, "daily_lunch_served": null, "isp": 0.0639, "active": true }] };
        schoolInfo['state_code'] = state;
        for (let school of schoolInfo.schools) {
            school['daily_breakfast_served'] = parseInt(school['total_enrolled'] * 0.15);
            school['daily_lunch_served'] = parseInt(school['total_enrolled'] * 0.55);
        }

        console.log(schoolInfo.schools[0]['daily_breakfast_served']);
        return schoolInfo;
    }

    // Format the solution so that it can be pasted into the Meals Count Excel Worksheet
    function formatOptimalSolution(solution) {
        return JSON.stringify(solution);
    }

    // Post and poll for the optimal solution for the given district using the enriched school info object
    function optimize(schoolInfo) {
        let results = new Promise((resolve, reject) => {

            // Get the URL where the optimal solution will be posted
            let requestObj = {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(schoolInfo),
            };
            let solutionUrl = fetch('https://www.mealscount.com/api/districts/optimize-async/', requestObj)
                .then(resp => resp.json())
                .then(data => { console.log(data.results_url); return data.results_url; })
                .catch(err => { console.error(err); reject(err); })

            // Start polling
            solutionUrl.then(url => {

                // Function which polls Lambda 
                function getResults() {
                    console.log('Polling: ' + url);
                    fetch(url)
                        .then(resp => {
                            // If results are ready then resolve the promise
                            if (resp.status == 200) {
                                clearInterval(poller);
                                resp.json()
                                    .then(data => resolve(formatOptimalSolution(data)))
                                    .catch(err => { console.error(err); reject(err); })
                            }
                        }).catch(err => { clearInterval(poller); console.error(err); reject(err); })
                }

                let poller = setInterval(getResults, 1000);
            }).catch(err => { clearInterval(poller); console.error(err); });
        });

        return results;
    }


    /*     MAIN     */

    // Get values from user
    let state = document.getElementById("state").value;
    let district = document.getElementById("district").value;
    let url = 'https://www.mealscount.com/static/' + state + '/' + district + '_district.json';

    // Show user we're doing work
    let working = document.getElementById("working");
    working.style.display = "block";

    // Fetch data and populate newly created worksheet

    
    // Doesn't work
    fetch(url)
        .then(resp => resp.json())
        .then(data => enrichSchoolInfo(data, state))
        .then(async function (enrichedData) {
            let optimalSolution = await optimize(enrichedData);
            Excel.run(context => {
                // Get Excel objects
                let sheet = context.workbook.worksheets.getItem("Meals Count");
                let range = sheet.getRange("A1");

                // Define what data should go in proxy objects
                range.values = optimalSolution;

                // Actually put that data there
                return context.sync().then(() => {
                    working.style.display = "none";
                    console.log('Populated workbook');
                });
            })
        }).catch(err => console.error(err));
}