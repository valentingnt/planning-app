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

document.getElementById('month-selector').addEventListener('change', ({ target: { value: month, checked } }) => {
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

  console.log(results)

  return results
}

function getDaysNameOfTheYear(year) {
  const result = {}

  const months = [
    "01",
    "02",
    "03",
    "04",
    "05",
    "06",
    "07",
    "08",
    "09",
    "10",
    "11",
    "12"
  ]

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
  return date >= startDate && date <= endDate
}

function getSlotsForDay(day, calendarSlots) {
  if (day.isHoliday) {
    return calendarSlots.isHoliday;
  } else if (day.isHolidayEve) {
    return calendarSlots.holidayEve;
  } else if (day.isSchoolHoliday) {
    return day.dayName === 'samedi' || day.dayName === 'dimanche' ?
      calendarSlots.isDuringSchoolHolidays.isweekEnd :
      calendarSlots.isDuringSchoolHolidays.isWeekDay;
  } else {
    if (day.dayName === 'samedi') {
      return calendarSlots.isNotDuringSchoolHolidays.saturday;
    } else if (day.dayName === 'dimanche') {
      return calendarSlots.isNotDuringSchoolHolidays.sunday;
    } else {
      return calendarSlots.isNotDuringSchoolHolidays.isWeekDay[0];
    }
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

      const isHoliday = holidaysLookup[date] !== undefined
      const isSchoolHoliday = Object.values(schoolHolidaysLookup).some(holiday => isDateInRange(date, holiday))

      acc[date] = {
        dayName,
        isHoliday,
        isSchoolHoliday
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

function slotsToHTML(slots) {
  let html = ''
  let dates = Object.keys(slots).sort()
  let currentMonth = null
  let openRow = false
  let dayOfWeek
  let lastDate = null

  dates.forEach(date => {
    let dateObj = new Date(date)
    let month = dateObj.getMonth()

    if (month !== currentMonth) {
      if (currentMonth !== null) {
        while (dayOfWeek !== 0) {
          html += '<td></td>'
          dayOfWeek = (dayOfWeek + 1) % 7
        }
        html += '</tr>'
        html += '</table>'
      }

      html += `<h2>${dateObj.toLocaleString('default', { month: 'long' })}</h2>`
      html += `<table data-month="${dateObj.getMonth()}"><tr><th>Lun</th><th>Mar</th><th>Mer</th><th>Jeu</th><th>Ven</th><th>Sam</th><th>Dim</th></tr>`
      currentMonth = month
      openRow = false
    }

    dayOfWeek = (dateObj.getDay() + 6) % 7

    if (!openRow) {
      html += '<tr>'
      for (let i = 0; i < dayOfWeek; i++) {
        html += '<td></td>'
      }
      openRow = true
    }

    html += '<td>'

    html += dateObj.getDate() + '<br>'

    slots[date].forEach((slot, i) => {
      html += `<div class="slot"><div class="time">${slot.start} - ${slot.end}</div>`
      for (let j = 0; j < slot.doctorsRequired; j++) {
        html += `<div class="doctor">Doctor ${j + 1}</div>`
      }
      html += '</div>'
    })

    html += '</td>'

    if (dayOfWeek === 6) {
      html += '</tr>'
      openRow = false
    }

    lastDate = date
  })

  if (openRow) {
    let lastDateOfMonth = new Date(lastDate);
    lastDateOfMonth.setMonth(lastDateOfMonth.getMonth() + 1);
    lastDateOfMonth.setDate(0);
    let lastDayOfMonth = (lastDateOfMonth.getDay() + 6) % 7;

    if (lastDayOfMonth !== 0) { // Change 6 to 0
      while (dayOfWeek !== 0) { // Change 6 to 0
        html += '<td></td>';
        dayOfWeek = (dayOfWeek + 1) % 7;
      }
    }
    html += '</tr>';
  }

  return html
}


createCalendar(2024).then((slots) => {
  const html = slotsToHTML(slots)

  document.getElementById('calendar').innerHTML = html

})