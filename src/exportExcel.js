import * as XLSX from 'xlsx'
import { supabase } from './supabaseClient'

const VEHICLE_ORDER = [
  '2W',
  'CAR_SUV',
  'LCV',
  'HCV',
  'AGRI',
  'BUS',
  'OTHER'
]

const LABELS = {
  '2W': '2 Wheeler',
  CAR_SUV: 'Car / SUV',
  LCV: 'LCV',
  HCV: 'HCV',
  AGRI: 'Agri',
  BUS: 'Bus',
  OTHER: 'Other'
}

export async function exportDayToExcel(date, shift) {

  // Get all records for this shift
  const { data, error } = await supabase
    .from('vehicle_events')
    .select('vehicle_type, traffic_type, created_at')
    .eq('count_date', date)
    .eq('shift', shift)

  if (error) {
    alert('Could not fetch data: ' + error.message)
    return
  }

  /*
    SHIFT TIMINGS

    Morning:
    09:00 AM - 09:00 PM

    Night:
    09:00 PM - 09:00 AM
  */

  const startHour =
    shift === 'Morning' ? 9 : 21

  const numberOfHours = 12

  // -----------------------------------------
  // CREATE HOURLY SLOTS
  // -----------------------------------------

  const slots = {}

  for (let i = 0; i < numberOfHours; i++) {

    const hour =
      (startHour + i) % 24

    const nextHour =
      (hour + 1) % 24

    const label =
      `${String(hour).padStart(2, '0')}:00-${String(nextHour).padStart(2, '0')}:00`

    slots[label] = {}

    VEHICLE_ORDER.forEach(vehicle => {

      slots[label][vehicle] = {
        TURN_IN: 0,
        PASS_THROUGH: 0
      }

    })
  }

  // -----------------------------------------
  // PUT EVENTS INTO HOURLY SLOTS
  // -----------------------------------------

  data?.forEach(row => {

    const dateObj =
      new Date(row.created_at)

    const hour =
      dateObj.getHours()

    let slotHour = null

    for (let i = 0; i < numberOfHours; i++) {

      const currentHour =
        (startHour + i) % 24

      if (hour === currentHour) {
        slotHour = currentHour
        break
      }
    }

    if (slotHour === null) {
      return
    }

    const nextHour =
      (slotHour + 1) % 24

    const label =
      `${String(slotHour).padStart(2, '0')}:00-${String(nextHour).padStart(2, '0')}:00`

    if (
      slots[label] &&
      slots[label][row.vehicle_type] &&
      row.traffic_type
    ) {

      slots[label][row.vehicle_type][row.traffic_type]++

    }

  })

  // -----------------------------------------
  // EXCEL HEADER
  // -----------------------------------------

  /*
  
  Time Slot

  2 Wheeler
      Turn-In | Pass-Through | Highway

  Car / SUV
      Turn-In | Pass-Through | Highway

  LCV
      Turn-In | Pass-Through | Highway

  ...

  Highway Total
  Turn-In Total
  Turn-In %

  */

  const headerRow1 = ['Time Slot']
  const headerRow2 = ['']

  VEHICLE_ORDER.forEach(vehicle => {

    headerRow1.push(LABELS[vehicle])
    headerRow1.push('')
    headerRow1.push('')

    headerRow2.push('Turn-In')
    headerRow2.push('Pass-Through')
    headerRow2.push('Highway')

  })

  headerRow1.push('Highway Total')
  headerRow1.push('Turn-In Total')
  headerRow1.push('Turn-In %')

  headerRow2.push('')
  headerRow2.push('')
  headerRow2.push('')

  const rows = [
    headerRow1,
    headerRow2
  ]

  // -----------------------------------------
  // GRAND TOTALS
  // -----------------------------------------

  const grandTotal = {}

  VEHICLE_ORDER.forEach(vehicle => {

    grandTotal[vehicle] = {
      TURN_IN: 0,
      PASS_THROUGH: 0,
      HIGHWAY: 0
    }

  })

  let grandHighwayTotal = 0
  let grandTurnInTotal = 0

  // -----------------------------------------
  // HOURLY ROWS
  // -----------------------------------------

  Object.entries(slots).forEach(
    ([timeSlot, counts]) => {

      let rowHighwayTotal = 0
      let rowTurnInTotal = 0

      const row = [timeSlot]

      VEHICLE_ORDER.forEach(vehicle => {

        const turnIn =
          counts[vehicle].TURN_IN

        const passThrough =
          counts[vehicle].PASS_THROUGH

        // IMPORTANT:
        // Highway = Turn-In + Pass-Through
        const highway =
          turnIn + passThrough

        row.push(turnIn)
        row.push(passThrough)
        row.push(highway)

        rowTurnInTotal += turnIn
        rowHighwayTotal += highway

        grandTotal[vehicle].TURN_IN += turnIn
        grandTotal[vehicle].PASS_THROUGH += passThrough
        grandTotal[vehicle].HIGHWAY += highway

      })

      // Highway Total
      row.push(rowHighwayTotal)

      // Turn-In Total
      row.push(rowTurnInTotal)

      // Turn-In %
      const percentage =
        rowHighwayTotal > 0
          ? (
              (rowTurnInTotal /
                rowHighwayTotal) *
              100
            ).toFixed(2) + '%'
          : '0.00%'

      row.push(percentage)

      // Only show hours where something happened
      if (rowHighwayTotal > 0) {
        rows.push(row)
      }

      grandHighwayTotal +=
        rowHighwayTotal

      grandTurnInTotal +=
        rowTurnInTotal

    }
  )

  // -----------------------------------------
  // GRAND TOTAL ROW
  // -----------------------------------------

  const grandRow = ['Grand Total']

  VEHICLE_ORDER.forEach(vehicle => {

    grandRow.push(
      grandTotal[vehicle].TURN_IN
    )

    grandRow.push(
      grandTotal[vehicle].PASS_THROUGH
    )

    grandRow.push(
      grandTotal[vehicle].HIGHWAY
    )

  })

  grandRow.push(grandHighwayTotal)
  grandRow.push(grandTurnInTotal)

  const grandPercentage =
    grandHighwayTotal > 0
      ? (
          (grandTurnInTotal /
            grandHighwayTotal) *
          100
        ).toFixed(2) + '%'
      : '0.00%'

  grandRow.push(grandPercentage)

  rows.push(grandRow)

  // -----------------------------------------
  // CREATE WORKSHEET
  // -----------------------------------------

  const worksheet =
    XLSX.utils.aoa_to_sheet(rows)

  // -----------------------------------------
  // MERGE HEADERS
  // -----------------------------------------

  worksheet['!merges'] = []

  let column = 1

  VEHICLE_ORDER.forEach(() => {

    worksheet['!merges'].push({
      s: {
        r: 0,
        c: column
      },

      e: {
        r: 0,
        c: column + 2
      }
    })

    column += 3

  })

  // Time Slot
  worksheet['!merges'].push({
    s: {
      r: 0,
      c: 0
    },

    e: {
      r: 1,
      c: 0
    }
  })

  // Highway Total
  worksheet['!merges'].push({
    s: {
      r: 0,
      c: column
    },

    e: {
      r: 1,
      c: column
    }
  })

  // Turn-In Total
  worksheet['!merges'].push({
    s: {
      r: 0,
      c: column + 1
    },

    e: {
      r: 1,
      c: column + 1
    }
  })

  // Turn-In %
  worksheet['!merges'].push({
    s: {
      r: 0,
      c: column + 2
    },

    e: {
      r: 1,
      c: column + 2
    }
  })

  // -----------------------------------------
  // COLUMN WIDTHS
  // -----------------------------------------

  worksheet['!cols'] = [
    { wch: 20 },

    ...VEHICLE_ORDER.flatMap(() => [
      { wch: 11 }, // Turn-In
      { wch: 14 }, // Pass-Through
      { wch: 11 }  // Highway
    ]),

    { wch: 15 }, // Highway Total
    { wch: 15 }, // Turn-In Total
    { wch: 12 }  // Turn-In %
  ]

  // -----------------------------------------
  // FREEZE HEADER
  // -----------------------------------------

  worksheet['!freeze'] = {
    xSplit: 1,
    ySplit: 2
  }

  // -----------------------------------------
  // CREATE WORKBOOK
  // -----------------------------------------

  const workbook =
    XLSX.utils.book_new()

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    `${shift} Summary`
  )

  // -----------------------------------------
  // DOWNLOAD
  // -----------------------------------------

  XLSX.writeFile(
    workbook,
    `BPCL-Vehicle-Count-${date}-${shift}.xlsx`
  )
}