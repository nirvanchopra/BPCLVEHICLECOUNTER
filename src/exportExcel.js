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

  // =====================================================
  // GET ALL RECORDS
  // Supabase normally returns max 1000 records per request.
  // Fetch in batches so Excel gets ALL records.
  // =====================================================

  let allData = []
  let from = 0
  const pageSize = 1000

  while (true) {

    const { data: pageData, error } = await supabase
      .from('vehicle_events')
      .select('vehicle_type, traffic_type, created_at')
      .eq('count_date', date)
      .eq('shift', shift)
      .range(from, from + pageSize - 1)

    if (error) {
      alert('Could not fetch data: ' + error.message)
      return
    }

    allData = [
      ...allData,
      ...(pageData || [])
    ]

    // Last batch reached
    if (!pageData || pageData.length < pageSize) {
      break
    }

    from += pageSize
  }

  console.log(
    `Exporting ${allData.length} vehicle events`
  )

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

  // =====================================================
  // CREATE HOURLY SLOTS
  // =====================================================

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

  // =====================================================
  // PUT EVENTS INTO HOURLY SLOTS
  // =====================================================

  allData.forEach(row => {

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
      slots[label][row.vehicle_type]
    ) {

      if (row.traffic_type === 'TURN_IN') {

        slots[label][row.vehicle_type].TURN_IN++

      }

      else if (row.traffic_type === 'PASS_THROUGH') {

        slots[label][row.vehicle_type].PASS_THROUGH++

      }

    }

  })

  // =====================================================
  // EXCEL HEADER
  // =====================================================

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

  // =====================================================
  // GRAND TOTALS
  // =====================================================

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

  // =====================================================
  // HOURLY ROWS
  // =====================================================

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

        /*
          IMPORTANT:

          Turn-In + automatically represents:
          1 Turn-In vehicle
          AND therefore 1 Highway vehicle.

          Highway + represents:
          1 Pass-Through vehicle.

          Therefore:

          HIGHWAY =
          TURN-IN + PASS-THROUGH
        */

        const highway =
          turnIn + passThrough

        row.push(turnIn)
        row.push(passThrough)
        row.push(highway)

        rowTurnInTotal += turnIn
        rowHighwayTotal += highway

        grandTotal[vehicle].TURN_IN +=
          turnIn

        grandTotal[vehicle].PASS_THROUGH +=
          passThrough

        grandTotal[vehicle].HIGHWAY +=
          highway

      })

      // =================================================
      // HOURLY TOTALS
      // =================================================

      row.push(rowHighwayTotal)
      row.push(rowTurnInTotal)

      const percentage =
        rowHighwayTotal > 0
          ? (
              (rowTurnInTotal /
                rowHighwayTotal) * 100
            ).toFixed(2) + '%'
          : '0.00%'

      row.push(percentage)

      // Only show hours with vehicles
      if (rowHighwayTotal > 0) {
        rows.push(row)
      }

      grandHighwayTotal +=
        rowHighwayTotal

      grandTurnInTotal +=
        rowTurnInTotal

    }
  )

  // =====================================================
  // GRAND TOTAL ROW
  // =====================================================

  const grandRow = ['Grand Total']

  VEHICLE_ORDER.forEach(vehicle => {

    const turnIn =
      grandTotal[vehicle].TURN_IN

    const passThrough =
      grandTotal[vehicle].PASS_THROUGH

    /*
      FINAL VEHICLE HIGHWAY TOTAL

      Highway = Turn-In + Pass-Through
    */

    const highway =
      turnIn + passThrough

    grandRow.push(turnIn)
    grandRow.push(passThrough)
    grandRow.push(highway)

  })

  // =====================================================
  // FINAL GRAND TOTAL
  // =====================================================

  /*
    Do NOT add Highway again.

    Highway already contains:
    Turn-In + Pass-Through
  */

  grandRow.push(grandHighwayTotal)
  grandRow.push(grandTurnInTotal)

  const grandPercentage =
    grandHighwayTotal > 0
      ? (
          (grandTurnInTotal /
            grandHighwayTotal) * 100
        ).toFixed(2) + '%'
      : '0.00%'

  grandRow.push(grandPercentage)

  rows.push(grandRow)

  // =====================================================
  // CREATE WORKSHEET
  // =====================================================

  const worksheet =
    XLSX.utils.aoa_to_sheet(rows)

  // =====================================================
  // MERGE HEADERS
  // =====================================================

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

  // =====================================================
  // COLUMN WIDTHS
  // =====================================================

  worksheet['!cols'] = [

    { wch: 20 },

    ...VEHICLE_ORDER.flatMap(() => [
      { wch: 11 },
      { wch: 14 },
      { wch: 11 }
    ]),

    { wch: 15 },
    { wch: 15 },
    { wch: 12 }

  ]

  // =====================================================
  // FREEZE HEADER
  // =====================================================

  worksheet['!freeze'] = {
    xSplit: 1,
    ySplit: 2
  }

  // =====================================================
  // CREATE WORKBOOK
  // =====================================================

  const workbook =
    XLSX.utils.book_new()

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    `${shift} Summary`
  )

  // =====================================================
  // DOWNLOAD
  // =====================================================

  XLSX.writeFile(
    workbook,
    `BPCL-Vehicle-Count-${date}-${shift}.xlsx`
  )
}