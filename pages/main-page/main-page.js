const Excel = require('excel4node')

const yearPicker = document.getElementById('year-picker')
const monthPicker = document.getElementById('month-picker')
const generateBtn = document.getElementById('generateBtn')
const travelPicker = document.getElementById('travel-picker')

const COLUMNS = ['Day in Week', 'Day of Month', 'Hours', 'Extra Hours', 'Travels', 'Notes']
const DAYS_OF_WEEK = { 0: "Sunday", 1: "Monday", 2: "Tuesday", 3: "Wednesday", 4: "Thursday", 5: "Friday", 6: "Saturday" }
const FULL_HOURS = 9


let currentDate = new Date()
let travelBase = 5.9
travelPicker.value = travelBase

travelPicker.addEventListener('input', ev => {
    if (travelPicker.value != '') {
        travelBase = travelPicker.value
    }
})

yearPicker.value = currentDate.getFullYear()
yearPicker.max = currentDate.getFullYear()
yearPicker.min = currentDate.getFullYear() - 10
monthPicker.value = currentDate.getMonth() + 1

generateBtn.addEventListener('click', event => {
    event.preventDefault()


    let workbook = new Excel.Workbook();
    let ws = workbook.addWorksheet(`${monthPicker.value}.${yearPicker.value}`)

    let headerStyle = workbook.createStyle({
        font: {
            size: 12
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#87b3c4'
        }
    });

    let weekendStyle = workbook.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#dfde9d'
        }
    });


    // write table headers
    for (let col = 0; col < COLUMNS.length; col++) {
        ws.cell(1, col + 1).string(COLUMNS[col]).style(headerStyle)
    }

    let { numOfDays, firstDay } = monthInfo(monthPicker.value, yearPicker.value)
    for (let day = 1; day <= numOfDays; day++) {
        if (firstDay % 7 > 4) {
            ws.cell(day + 1, 1).string(DAYS_OF_WEEK[firstDay % 7]).style(weekendStyle) // days of week for weekend
        } else {
            ws.cell(day + 1, 1).string(DAYS_OF_WEEK[firstDay % 7]) // days of week
        }
        ws.cell(day + 1, 2).number(day) // day in month
        firstDay++

        // write extra hours formula
        ws.cell(day + 1, 4).formula(`IF(${FULL_HOURS}-C${day + 1}<0,C${day + 1}-${FULL_HOURS},0)`)

        // write travels formula
        ws.cell(day + 1, 5).formula(`${travelBase}*2*IF(C${day + 1}>0,1,0)`)
    }


    workbook.write('Excel.xlsx');
})


function monthInfo(month, year) {
    let date = new Date(year, month, 0)
    let numOfDays = date.getDate()
    date.setDate(1)
    return {
        numOfDays: numOfDays,
        firstDay: date.getDay()
    }
}