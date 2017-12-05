const Excel = require('excel4node')

const yearPicker = document.getElementById('year-picker')
const generateBtn = document.getElementById('generateBtn')

const TITLES = ['Day in Week', 'Day of Month', 'Hours', 'Extra Hours', 'Travels', 'Notes']
const DAYS_OF_WEEK = { 0: "Sunday", 1: "Monday", 2: "Tuesday", 3: "Wednesday", 4: "Thursday", 5: "Friday", 6: "Saturday" }
const FULL_HOURS = 9


let currentDate = new Date()

yearPicker.value = currentDate.getFullYear()
yearPicker.max = currentDate.getFullYear()
yearPicker.min = currentDate.getFullYear() - 10


generateBtn.addEventListener('click', event => {
    event.preventDefault()


    let workbook = new Excel.Workbook();
    let ws = workbook.addWorksheet('sheet1')

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
    for (let col = 0; col < TITLES.length; col++) {
        ws.cell(1, col + 1).string(TITLES[col]).style(headerStyle)
    }

    let { numOfDays, firstDay } = monthInfo(currentDate.getMonth() + 1, currentDate.getFullYear())
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
    }


    // // Set value of cell A1 to 100 as a number type styled with paramaters of style
    // ws.cell(1, 1).number(100).style(style);

    // // Set value of cell B1 to 300 as a number type styled with paramaters of style
    // ws.cell(1, 2).number(200).style(style);

    // // Set value of cell C1 to a formula styled with paramaters of style
    // ws.cell(1, 3).formula('A1 + B1').style(style);

    // // Set value of cell A2 to 'string' styled with paramaters of style
    // ws.cell(2, 1).string('string').style(style);

    // // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
    // ws.cell(3, 1).bool(true).style(style).style({ font: { size: 14 } });

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