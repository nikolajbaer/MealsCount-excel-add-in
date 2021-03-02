'use strict';

(function () {
    Office.onReady(function () {
        $(document).ready(function () {
            createWorksheet();
        });
    });

    // Add a worksheet to store results within the current workbook
    function createWorksheet() {
        Excel.run(context => {
            var sheets = context.workbook.worksheets;
            sheets.add("Meals Count");

            return context.sync().then(
                   () => console.log('Added worksheet'),
                   err => console.error(err)
            );
        });
    }
})()

// React to user's button click
function refreshData() {

    // Compute new values and add them to school info object
    function enrichSchoolInfo(schoolInfo, state) {

        // Validate input?

        schoolInfo['state_code'] = state;
        for (let school of schoolInfo.schools) {
            school['daily_breakfast_served'] = parseInt(school['total_enrolled'] * 0.15);
            school['daily_lunch_served'] = parseInt(school['total_enrolled'] * 0.55);
        }
        return schoolInfo;
    }

    // Format the solution so that it can be pasted into the Meals Count Excel Worksheet
    // This is super brittle but will need to change anyway
    function formatOptimalSolution(solution) {
            /*
           let results = [];
           let bestStrategy = solution.best_strategy
           for (let strategy of solution.strategies) {
               if (strategy.name === bestStrategy) {
                   for (let group of strategy.groups) {
                       results.concat([group.name].concat(group.school_reimbursements))
                   }
               }
           }
           return results;
           */
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

    // Get district info, reformat it, POST to get solution, poll until solution is ready
    fetch(url)
        .then(resp => resp.json())
        .then(data => enrichSchoolInfo(data, state))
        .then(async function (enrichedData) {
            let optimalSolution = await optimize(enrichedData);
            Excel.run(context => {
                // Get Excel objects
                let sheet = context.workbook.worksheets.getItem("Meals Count");
                let range = sheet.getRange("A1");

                // Define what data should go in Excel "proxy" objects
                range.values = formatOptimalSolution(optimalSolution);

                // Actually put that data there
                return context.sync().then(() => {
                    working.style.display = "none";
                    console.log('Populated workbook');
                });
            })
        }).catch(err => console.error(err));
}