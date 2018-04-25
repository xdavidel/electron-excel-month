const Excel = require('excel4node')

const yearPicker = document.getElementById('year-picker')
const monthPicker = document.getElementById('month-picker')
const generateBtn = document.getElementById('generateBtn')
const travelPicker = document.getElementById('travel-picker')

const COLUMNS = ['Day in Week', 'Day of Month', 'Hours', '120%', '150%', 'Travels', 'Notes']
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

    let separatorStyle = workbook.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#121212'
        }
    })

    let sumStyle = workbook.createStyle({
        font: {
            color: '#FF0800',
            size: 12,
            bold: true
        }

    })

    let { numOfDays, firstDay } = monthInfo(monthPicker.value, yearPicker.value)

    for (let col = 0; col < COLUMNS.length; col++) {

        // write table headers
        ws.cell(1, col + 1).string(COLUMNS[col]).style(headerStyle)

        // write table end line
        ws.cell(numOfDays + 2, col + 1).string('').style(separatorStyle)
    }

    for (let day = 1; day <= numOfDays; day++) {
        if (firstDay % 7 > 4) {
            ws.cell(day + 1, 1).string(DAYS_OF_WEEK[firstDay % 7]).style(weekendStyle) // days of week for weekend
        } else {
            ws.cell(day + 1, 1).string(DAYS_OF_WEEK[firstDay % 7]) // days of week
        }
        ws.cell(day + 1, 2).number(day) // day in month
        firstDay++

        // write extra hours formula
        ws.cell(day + 1, 4).formula(`IF(${FULL_HOURS}-C${day + 1}<0,IF(C${day + 1}-${FULL_HOURS}>2,2,C${day + 1}-${FULL_HOURS}),0)`)

        // write 150% formula
        ws.cell(day + 1, 5).formula(`IF(C${day + 1}-${FULL_HOURS}>2,C${day + 1}-${FULL_HOURS}-2,0)`)

        // write travels formula
        ws.cell(day + 1, 6).formula(`${travelBase}*IF(C${day + 1}>0,1,0)`)
    }

    // write table sums
    for (let col = 2; col < COLUMNS.length - 1; col++) {
        let currentColumn = getExcelColumn(col + 1)
        ws.cell(numOfDays + 3, col + 1).formula(`SUM(${currentColumn}2:${currentColumn}${numOfDays + 1})`).style(sumStyle)
    }



    workbook.write(`${yearPicker.value}_${monthPicker.value}.xlsx`);
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

function getExcelColumn(num) {
    let char = 'A'

    return String.fromCharCode(char.charCodeAt(0) + num - 1)
}