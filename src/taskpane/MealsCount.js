const IMPORT_SHEET_NAME =  "MealsCount Import"
const MC_HEADERS = ["school_name","total_enrolled","total_eligible","daily_breakfast_served","daily_lunch_served","grouping"] 

export function createWorksheet() {
    Excel.run(context => {
        var sheet = context.workbook.worksheets.getItemOrNullObject(IMPORT_SHEET_NAME)
        return context.sync().then( () => {
            console.log(sheet)
            if(sheet.isNullObject){
                console.log("adding MealsCount sheet")
                sheet = context.workbook.worksheets.add(IMPORT_SHEET_NAME)
            }else{
                console.log("MealsCount Sheet Found")
            }
            // ensure we have headers
            var headerRange = sheet.getRange("A1:F1")
            headerRange.values = [MC_HEADERS]
            return context.sync()
        })
    });
}

export function runOptimize(e) {
    set_notice("Please wait, sending school data to mealscount.com")
    Excel.run(context => {
        var sheet = context.workbook.worksheets.getItemOrNullObject(IMPORT_SHEET_NAME)
        const schools = sheet.getRange("MC_SCHOOLDATA") // LA and NY have more than 1000 schools..
        schools.load('values')
        const school_data = []
        context.sync().then( () => {
            // Gather data from School Range
            console.log(schools.values)

            schools.values.forEach( row => {
                if(row[0]){
                    school_data.push({
                        school_code:row[0],
                        school_name:row[0],
                        total_enrolled:row[1],
                        total_eligible:row[2],
                        daily_breakfast_served:row[3],
                        daily_lunch_served:row[4],                     
                    })
                }
            })

            console.log(school_data)

            let state = document.getElementById("state").value;
            let district = document.getElementById("district").value;

            if(!state || !district){
                set_notice("Please make sure you have a the district and state codes specified")
            }else{
                const rqdata = build_optimize_request_object(school_data,state,district)
                console.log(rqdata)
                do_optimize(rqdata)
            }
        })
        return context.sync()
    })
    e.preventDefault()
    return false
}

function set_notice(text){
    const notice = document.getElementById("mc_notice")
    notice.innerText = text 
    notice.style.display = "block"
}

function build_optimize_request_object(school_data,state,district){
    return {
        code: district,
        name: district,
        state_code: state.toLowerCase(),
        schools: school_data,
        rates: {
            free_lunch: 3.41,
            paid_lunch: 0.32,
            free_bfast: 1.84,
            paid_bfast: 0.31
        }
    }
}

function update_school_groupings(data){
    const school_group_map = {}
    let g_indx = 1
    data.strategies[data.best_index].groups.forEach( g => {
        g.school_codes.forEach( code => {
            school_group_map[code] = g_indx
        }) 
        g_indx += 1 // map name to a 1-based group number for less confusion
    })

    console.log("School Grouping Data!",school_group_map)

    // fill in groupings column in our named range
    Excel.run(context => {
        var sheet = context.workbook.worksheets.getItemOrNullObject(IMPORT_SHEET_NAME)
        const range = sheet.getRange("MC_SCHOOLDATA") // LA and NY have more than 1000 schools..
        range.load('values')

        return context.sync().then( () => {
            range.values = range.values.map( row => {
                if(!row[0]){ return row } // skip empty rows
                if(school_group_map[row[0]] != undefined){
                    row[5] = school_group_map[row[0]]
                }else{
                    row[5] = "NOT-FOUND"
                }
                console.log(row)
                return row
            }) 
            set_notice("Optimization Complete!")
            return context.sync()
        }) 
    })
}

function do_optimize(schoolInfo) {
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
                                .then(data => {
                                    update_school_groupings(data) 
                                })
                                .catch(err => { console.error(err); reject(err); })
                        }
                    }).catch(err => { clearInterval(poller); console.error(err); reject(err); })
            }

            let poller = setInterval(getResults, 1000);
        }).catch(err => { clearInterval(poller); console.error(err); });
    });
}

/*
 * Uses specfieid state/district to fill in mealscount district data tab
*/
export function fillSchoolData(e){

    // Get values from user
    let state = document.getElementById("state").value;
    let district = document.getElementById("district").value;
    let url = 'https://www.mealscount.com/static/' + state.toLowerCase() + '/' + district + '_district.json';

    console.log("Querying" + url)

    set_notice("Please wait, retrieving school data")

    // Get district info, reformat it, POST to get solution, poll until solution is ready
    fetch(url)
        .then(resp => resp.json())
        .then(async function (data) {
            Excel.run(context => {
                console.log("received data",data)
                let sheet = context.workbook.worksheets.getItem(IMPORT_SHEET_NAME);
                let range = sheet.getRange("A2:F"+(data.schools.length+1));
                console.log("getting range A2:F"+(data.schools.length+1))

                const school_data = []
                data.schools.forEach( school => {
                    school_data.push( [
                        school.school_name,
                        school.total_enrolled,
                        school.total_eligible,
                        school.daily_breakfast_served,
                        school.daily_lunch_served,
                        1,
                    ])
                })

                console.log(school_data.length,school_data[0].length,school_data)
                range.values = school_data
                sheet.names.add("MC_SCHOOLDATA",range)

                const code_range = sheet.getRange("A2:A"+(data.schools.length+1))
                code_range.numberFormat = "#"

                return context.sync().then(() => {
                    working.style.display = "none";
                    console.log('Populated workbook');
                });
            })
        }).catch(err => console.error(err));

    e.preventDefault()
    return false
}