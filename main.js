import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'

let DOCTOR_NAMES = []
const DAYS_OF_WEEK = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'];
const CALENDAR_SLOTS = {
  isHoliday: [
    { start: "08:00", end: "12:00", doctorsRequired: 4 },
    { start: "12:00", end: "16:00", doctorsRequired: 2 },
    { start: "16:00", end: "20:00", doctorsRequired: 1 },
    { start: "20:00", end: "24:00", doctorsRequired: 1 }
  ],
  holidayEve: [
    { start: "08:00", end: "12:00", doctorsRequired: 3 },
    { start: "12:00", end: "16:00", doctorsRequired: 2 },
    { start: "16:00", end: "20:00", doctorsRequired: 2 },
    { start: "20:00", end: "24:00", doctorsRequired: 2 }
  ],
  isDuringSchoolHolidays: {
    isWeekDay: [
      { start: "08:00", end: "13:00", doctorsRequired: 2 },
      { start: "13:00", end: "18:00", doctorsRequired: 2 },
      { start: "18:00", end: "22:00", doctorsRequired: 1 },
      { start: "18:00", end: "24:00", doctorsRequired: 1 }
    ],
    isweekEnd: [
      { start: "08:00", end: "12:00", doctorsRequired: 4 },
      { start: "12:00", end: "16:00", doctorsRequired: 2 },
      { start: "16:00", end: "20:00", doctorsRequired: 1 },
      { start: "20:00", end: "24:00", doctorsRequired: 1 }
    ]
  },
  isNotDuringSchoolHolidays: {
    isWeekDay: [
      [
        { start: "08:00", end: "13:00", doctorsRequired: 1 },
        { start: "13:00", end: "18:00", doctorsRequired: 1 },
        { start: "18:00", end: "22:00", doctorsRequired: 1 },
        { start: "18:00", end: "24:00", doctorsRequired: 1 }
      ],
      [
        { start: "08:00", end: "14:00", doctorsRequired: 1 },
        { start: "14:00", end: "20:00", doctorsRequired: 1 },
        { start: "18:00", end: "22:00", doctorsRequired: 1 },
        { start: "20:00", end: "24:00", doctorsRequired: 1 }
      ]
    ],
    saturday: [
      { start: "08:00", end: "12:00", doctorsRequired: 3 },
      { start: "12:00", end: "16:00", doctorsRequired: 2 },
      { start: "16:00", end: "20:00", doctorsRequired: 2 },
      { start: "20:00", end: "24:00", doctorsRequired: 2 }
    ],
    sunday: [
      { start: "08:00", end: "12:00", doctorsRequired: 3 },
      { start: "12:00", end: "16:00", doctorsRequired: 2 },
      { start: "16:00", end: "20:00", doctorsRequired: 2 },
      { start: "20:00", end: "24:00", doctorsRequired: 1 }
    ],
  },
}

const yearSelector = document.getElementById('year-selector')
const monthSelector = document.getElementById('month-selector')
const generatePlanningButton = document.getElementById('download-planning')
const doctorsList = document.getElementById('doctors-list-input')
const addDoctorsListButton = document.getElementById('add-doctor')
const currentDoctorsListPlaceholder = document.getElementById('current-doctors-list')

addDoctorsListButton.addEventListener('click', () => {
  const doctorName = doctorsList.value

  if (doctorName) {
    DOCTOR_NAMES = doctorName.split(',').map(name => name.trim())

    currentDoctorsListPlaceholder.innerHTML = DOCTOR_NAMES.map(name => `<li>${name}</li>`).join('')
  } else {
    currentDoctorsListPlaceholder.innerHTML = "Aucun médecin n'a encore été ajouté."
  }

}, { passive: true })

yearSelector.setAttribute('selected', true)

monthSelector.addEventListener('change', ({ target: { value: month, checked } }) => {
  document.querySelectorAll('table').forEach(table => {
    if (table.getAttribute('data-month') === month) {
      table.style.display = checked ? 'table' : 'none'
    }
  })
}, { passive: true })

generatePlanningButton.addEventListener('click', () => {
  const selectedYear = parseInt(yearSelector.value)

  createCalendar(selectedYear).then((slots) => {
    const selectedMonths = Array.from(monthSelector.querySelectorAll('input:checked')).map(input => parseInt(input.value))

    createExcelFile(slots, selectedMonths)
  })
}, { passive: true })


async function getSchoolHolidays(year) {
  const URL_API = `https://data.education.gouv.fr/api/explore/v2.1/catalog/datasets/fr-en-calendrier-scolaire/records?limit=20&refine=zones%3A%22Zone%20A%22&refine=location%3A%22Bordeaux%22&refine=population%3A%22-%22&refine=population%3A%22%C3%89l%C3%A8ves%22&refine=annee_scolaire%3A%22${year - 1}-${year}%22&refine=annee_scolaire%3A%22${year}-${year + 1}%22`
  const response = await fetch(URL_API)
  const { results } = await response.json()

  return results
}

async function getHolidays(year) {
  const URL_API = `https://calendrier.api.gouv.fr/jours-feries/metropole/${year}.json`
  const response = await fetch(URL_API)
  const results = await response.json()

  return results
}

function getDaysNameOfTheYear(year) {
  const result = {}
  const months = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

  for (let month = 0; month < 12; month++) {
    const monthNumber = months[month]

    result[monthNumber] = {}

    for (let day = 1; day <= 31; day++) {
      const date = new Date(`${year}-${month + 1}-${day}`)

      if (date.getFullYear() === year && date.getMonth() === month) {
        const dayName = date.toLocaleDateString("us-US", { weekday: "long" }).toLowerCase()

        result[monthNumber][day.toString().padStart(2, "0")] = dayName
      } else {
        break
      }
    }
  }

  return result
}

function isDateInRange(date, range) {
  const startDate = range.start_date.split('T')[0]
  const endDate = range.end_date.split('T')[0]

  return date > startDate && date <= endDate
}

function getSlotsForDay(day, calendarSlots) {
  const { isHoliday, isHolidayEve, isSchoolHoliday, dayName } = day
  const { isHoliday: holidaySlots, holidayEve, isDuringSchoolHolidays, isNotDuringSchoolHolidays } = calendarSlots

  if (isHoliday) return holidaySlots
  if (isHolidayEve) return holidayEve

  if (isSchoolHoliday) {
    const isWeekend = dayName === 'samedi' || dayName === 'dimanche'
    return isWeekend ? isDuringSchoolHolidays.isweekEnd : isDuringSchoolHolidays.isWeekDay
  }

  switch (dayName) {
    case 'samedi':
      return isNotDuringSchoolHolidays.saturday
    case 'dimanche':
      return isNotDuringSchoolHolidays.sunday
    default:
      return isNotDuringSchoolHolidays.isWeekDay[0]
  }
}

async function createCalendar(year) {
  const schoolHolidays = await getSchoolHolidays(year)
  const holidays = await getHolidays(year)
  const daysNameOfTheYear = getDaysNameOfTheYear(year)
  const schoolHolidaysLookup = schoolHolidays.reduce((acc, holiday) => {
    acc[holiday.start_date.split('T')[0]] = holiday

    return acc
  }, {})

  const holidaysLookup = Object.keys(holidays).reduce((acc, holiday) => {
    acc[holiday] = true

    return acc
  }, {})

  const calendar = Object.keys(daysNameOfTheYear).reduce((acc, month) => {
    const days = Object.keys(daysNameOfTheYear[month])

    days.forEach(day => {
      const dayName = daysNameOfTheYear[month][day]
      const date = `${year}-${month}-${day}`
      const currentDate = new Date(year, month - 1, day)

      currentDate.setDate(currentDate.getDate() + 1)
      const dateAfter = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}-${String(currentDate.getDate()).padStart(2, '0')}`

      currentDate.setDate(currentDate.getDate() - 2)
      const dateBefore = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}-${String(currentDate.getDate()).padStart(2, '0')}`

      let isHoliday = holidaysLookup[date] !== undefined
      const isSchoolHoliday = Object.values(schoolHolidaysLookup).some(holiday => isDateInRange(date, holiday))
      const isHolidayEve = holidaysLookup[dateAfter] !== undefined

      if (holidaysLookup[dateAfter] !== undefined && new Date(dateAfter).getDay() === 2) {
        isHoliday = true
      }

      if (holidaysLookup[dateBefore] !== undefined && new Date(dateBefore).getDay() === 4) {
        isHoliday = true
      }

      acc[date] = {
        dayName,
        isHoliday,
        isSchoolHoliday,
        isHolidayEve
      }
    })

    return acc
  }, {})

  const slots = Object.keys(calendar).reduce((acc, date) => {
    const day = calendar[date]

    acc[date] = getSlotsForDay(day, CALENDAR_SLOTS)

    return acc
  }, {})

  return slots
}

function createHeaders(worksheet) {
  for (let i = 0; i < 7; i++) {
    const startColumn = (i * 4) + 1;
    const endColumn = startColumn + 3;

    if (!worksheet.getCell(1, startColumn).isMerged) {
      const cell = worksheet.getCell(1, startColumn);

      worksheet.mergeCells(1, startColumn, 1, endColumn);

      cell.value = DAYS_OF_WEEK[i];
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4C68AF' }
      };
      cell.font = {
        color: { argb: 'FFFFFFFF' },
        size: 24,
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center'
      };
    }
  }
}

function createDayGroup(worksheet, rowIndex, dayColumnStart, dayColumnEnd, dateObj) {
  if (!worksheet.getCell(rowIndex, dayColumnStart).isMerged) {
    const cell = worksheet.getCell(rowIndex, dayColumnStart);

    worksheet.mergeCells(rowIndex, dayColumnStart, rowIndex, dayColumnEnd);
    worksheet.getRow(rowIndex).height = 50;

    cell.value = dateObj.toLocaleDateString('fr-FR', { day: 'numeric', month: 'long' });
    cell.font = {
      size: 16,
      bold: true,
      italic: true
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center'
    };

  }
}

function fillSlots(worksheet, rowIndex, dayColumnStart, slotArray, doctors) {
  for (let i = 0; i < 4; i++) {
    const doctorsRequired = slotArray[i].doctorsRequired;

    worksheet.getCell(rowIndex + 1, dayColumnStart + i).value = `${slotArray[i].start}-${slotArray[i].end}`;

    for (let j = 0; j < doctorsRequired; j++) {
      const cell = worksheet.getCell(rowIndex + 2 + j, dayColumnStart + i);

      cell.value = `Doctor ${j + 1}`;
      cell.dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [`"${doctors.join(',')}"`]
      };
    }
  }
}

function applyBorderToDay(worksheet, rowIndex, dayColumnStart, dayColumnEnd) {
  const borderStyle = { style: 'thin' };

  for (let i = rowIndex; i < rowIndex + 6; i++) {
    for (let j = dayColumnStart; j <= dayColumnEnd; j++) {
      const cell = worksheet.getCell(i, j);

      cell.border = {
        top: i === rowIndex ? borderStyle : null,
        left: j === dayColumnStart ? borderStyle : null,
        bottom: i === rowIndex + 5 ? borderStyle : null,
        right: j === dayColumnEnd ? borderStyle : null
      };
    }
  }
}

function mergeRemainingCells(worksheet, rowIndex, dayColumnStart, slotArray) {
  for (let i = 0; i < 4; i++) {
    const doctorsRequired = slotArray[i].doctorsRequired;

    if (doctorsRequired < 4) {
      const mergeStart = rowIndex + 2 + doctorsRequired;
      const mergeEnd = rowIndex + 5;
      const cell = worksheet.getCell(mergeStart, dayColumnStart + i);

      worksheet.mergeCells(mergeStart, dayColumnStart + i, mergeEnd, dayColumnStart + i);

      cell.alignment = {
        vertical: 'top',
        horizontal: 'left'
      };
    }
  }
}

async function createExcelFile(slots, selectedMonths = null) {
  const workbook = new ExcelJS.Workbook();
  let worksheet;
  let currentMonth;
  let rowIndex = 2;

  const data = Object.entries(slots)
    .map(([date, slots]) => ({ date, slots }))
    .sort((a, b) => new Date(a.date) - new Date(b.date));

  for (const { date, slots: slotArray } of data) {
    const dateObj = new Date(date);
    const month = dateObj.getMonth();
    const dayOfWeek = (dateObj.getDay() + 6) % 7;

    if (selectedMonths && !selectedMonths.includes(month)) {
      continue;
    }

    if (month !== currentMonth) {
      const monthYear = dateObj.toLocaleString('default', { month: 'long', year: 'numeric' }).replace(/^\w/, c => c.toUpperCase());

      currentMonth = month;
      worksheet = workbook.addWorksheet(monthYear);
      rowIndex = 2;

      createHeaders(worksheet);
    }

    const dayColumnStart = (dayOfWeek * 4) + 1;
    const dayColumnEnd = dayColumnStart + 3;

    createDayGroup(worksheet, rowIndex, dayColumnStart, dayColumnEnd, dateObj);

    for (let i = 0; i < 4; i++) {
      const doctorsRequired = slotArray[i].doctorsRequired;

      worksheet.getCell(rowIndex + 1, dayColumnStart + i).value = `${slotArray[i].start}-${slotArray[i].end}`;

      for (let j = 0; j < doctorsRequired; j++) {
        worksheet.getCell(rowIndex + 2 + j, dayColumnStart + i).value = `Doctor ${j + 1}`;
      }
    }

    fillSlots(worksheet, rowIndex, dayColumnStart, slotArray, DOCTOR_NAMES);
    mergeRemainingCells(worksheet, rowIndex, dayColumnStart, slotArray);
    applyBorderToDay(worksheet, rowIndex, dayColumnStart, dayColumnEnd);

    worksheet.columns.forEach(column => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, cell => {
        const columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = maxLength < 10 ? 10 : maxLength;
    });

    if (dayOfWeek === 6) {
      rowIndex += 6;
    }
  }

  const file = await workbook.xlsx.writeBuffer();

  saveAs(new Blob([file]), 'Schedule.xlsx');
}