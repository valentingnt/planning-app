import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'

const calendarSlots = {
  isHoliday: [
    { start: "08:00", end: "12:00", doctorsRequired: 4 },
    { start: "12:00", end: "16:00", doctorsRequired: 2 },
    { start: "16:00", end: "20:00", doctorsRequired: 1 },
    { start: "20:00", end: "00:00", doctorsRequired: 1 }
  ],
  holidayEve: [
    { start: "08:00", end: "12:00", doctorsRequired: 3 },
    { start: "12:00", end: "16:00", doctorsRequired: 2 },
    { start: "16:00", end: "20:00", doctorsRequired: 2 },
    { start: "20:00", end: "00:00", doctorsRequired: 2 }
  ],
  isDuringSchoolHolidays: {
    isWeekDay: [
      { start: "08:00", end: "13:00", doctorsRequired: 2 },
      { start: "13:00", end: "18:00", doctorsRequired: 2 },
      { start: "18:00", end: "22:00", doctorsRequired: 1 },
      { start: "18:00", end: "00:00", doctorsRequired: 1 }
    ],
    isweekEnd: [
      { start: "08:00", end: "12:00", doctorsRequired: 4 },
      { start: "12:00", end: "16:00", doctorsRequired: 2 },
      { start: "16:00", end: "20:00", doctorsRequired: 1 },
      { start: "20:00", end: "00:00", doctorsRequired: 1 }
    ]
  },
  isNotDuringSchoolHolidays: {
    isWeekDay: [
      [
        { start: "08:00", end: "13:00", doctorsRequired: 1 },
        { start: "13:00", end: "18:00", doctorsRequired: 1 },
        { start: "18:00", end: "22:00", doctorsRequired: 1 },
        { start: "18:00", end: "00:00", doctorsRequired: 1 }
      ],
      [
        { start: "08:00", end: "14:00", doctorsRequired: 1 },
        { start: "14:00", end: "20:00", doctorsRequired: 1 },
        { start: "18:00", end: "22:00", doctorsRequired: 1 },
        { start: "20:00", end: "00:00", doctorsRequired: 1 }
      ]
    ],
    saturday: [
      { start: "08:00", end: "12:00", doctorsRequired: 3 },
      { start: "12:00", end: "16:00", doctorsRequired: 2 },
      { start: "16:00", end: "20:00", doctorsRequired: 2 },
      { start: "20:00", end: "00:00", doctorsRequired: 2 }
    ],
    sunday: [
      { start: "08:00", end: "12:00", doctorsRequired: 3 },
      { start: "12:00", end: "16:00", doctorsRequired: 2 },
      { start: "16:00", end: "20:00", doctorsRequired: 2 },
      { start: "20:00", end: "00:00", doctorsRequired: 1 }
    ],
  },
}

const monthSelector = document.getElementById('month-selector')

monthSelector.addEventListener('change', ({ target: { value: month, checked } }) => {
  document.querySelectorAll('table').forEach(table => {
    if (table.getAttribute('data-month') === month) {
      table.style.display = checked ? 'table' : 'none'
    }
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
    acc[date] = getSlotsForDay(day, calendarSlots)
    return acc
  }, {})

  return slots
}

function createMonth(dateObj) {
  let month = `<h2>${dateObj.toLocaleString('default', { month: 'long' })}</h2>`
  month += `<table data-month="${dateObj.getMonth()}"><tr><th>Lun</th><th>Mar</th><th>Mer</th><th>Jeu</th><th>Ven</th><th>Sam</th><th>Dim</th></tr>`

  return month
}

function createDay(dateObj) {
  return `<td class="day"><p class="days-number">${dateObj.getDate()}</p><br>`
}

function createSlot(slot) {
  let slotHtml = `<div class="slot"><div class="time">${slot.start} - ${slot.end}</div>`

  for (let j = 0; j < slot.doctorsRequired; j++) {
    slotHtml += `<div class="doctor">Doctor ${j + 1}</div>`
  }

  slotHtml += '</div>'

  return slotHtml
}

function fillEmptyDays(dayOfWeek) {
  let emptyDays = ''
  for (let i = 0; i < dayOfWeek; i++) {
    emptyDays += '<td></td>'
  }

  return emptyDays
}

function startNewMonth(html, dateObj, dayOfWeek) {
  html += '</tr></table>'
  html += createMonth(dateObj)
  html += '<tr>' + fillEmptyDays(dayOfWeek)

  return html
}

function endMonth(html, dayOfWeek) {
  if (dayOfWeek !== 0) {
    html += fillEmptyDays(7 - dayOfWeek) + '</tr>'
  }

  html += '</table>'

  return html
}

function startNewWeek(html) {
  html += '</tr><tr>'

  return html
}

function endWeek(html) {
  html += '</tr>'

  return html
}

function slotsToHTML(slots) {
  let html = ''
  let dates = Object.keys(slots).sort()
  let currentMonth = null
  let dayOfWeek
  let lastDate = null

  dates.forEach(date => {
    let dateObj = new Date(date)
    let month = dateObj.getMonth()

    if (month !== currentMonth) {
      if (currentMonth !== null) {
        html = endMonth(html, dayOfWeek)
      }

      dayOfWeek = (dateObj.getDay() + 6) % 7
      html = startNewMonth(html, dateObj, dayOfWeek)
      currentMonth = month
    }

    html += createDay(dateObj)

    slots[date].forEach((slot, i) => {
      html += createSlot(slot)
    })

    html += '</td>'

    dayOfWeek = (dayOfWeek + 1) % 7
    if (dayOfWeek === 0) {
      html = endWeek(html)
      if (date !== dates[dates.length - 1]) {
        html = startNewWeek(html)
      }
    }

    lastDate = date
  })

  html = endMonth(html, dayOfWeek)

  return html
}


createCalendar(2024).then((slots) => {
  const html = slotsToHTML(slots)
  document.getElementById('calendar').innerHTML = html
})

async function downloadExcelFile(slots, selectedMonths) {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Calendar')

  worksheet.columns = [
    { header: 'Date', key: 'date' },
    { header: 'Slots', key: 'slots' },
  ]

  for (const date in slots) {
    const month = date.split('-')[1]
    if (selectedMonths.includes(month)) {
      const slotData = slots[date].map(slot => `${slot.start} - ${slot.end}`).join(', ')
      worksheet.addRow({ date, slots: slotData })
    }
  }

  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })

  saveAs(blob, 'Calendar.xlsx')
}

function downloadSelectedView() {
  const selectedMonths = Array.from(document.querySelectorAll('#month-selector input:checked')).map(input => input.value)
  createCalendar(2024).then((slots) => {
    downloadExcelFile(slots, selectedMonths)
  })
}