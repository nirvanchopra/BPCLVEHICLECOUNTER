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
    Morning:
    09:00 AM - 09:00 PM

    Night:
    09:00 PM - 09:00 AM
  */

  let startHour
  let numberOfHours = 12

  if (shift === 'Morning') {
    startHour = 9
  } else {
    startHour = 21
  }

  // Create hourly slots
  const slots = {}

  for (let i = 0; i < numberOfHours; i++) {

    const hour = (startHour + i) % 24
    const nextHour = (hour + 1) % 24

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

  // Put every vehicle into its correct time slot
  data?.forEach(row => {

    const dateObj = new Date(row.created_at)
    const hour = dateObj.getHours()

    let slotHour = false

    for (let i = 0; i < numberOfHours; i++) {

      const currentHour = (startHour + i) % 24

      if (hour === currentHour) {
        slotHour = currentHour
        break
      }
    }

    if (slotHour === false) return

    const nextHour = (slotHour + 1) % 24

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

  /*
    HEADER

    Row 1:
    Time Slot | 2 Wheeler | Car/SUV | LCV | ...

    Row 2:
              | Turn In | Pass Through | Turn In | Pass Through ...
  */

  const headerRow1 = ['Time Slot']
  const headerRow2 = ['']

  VEHICLE_ORDER.forEach(vehicle => {
    headerRow1.push(LABELS[vehicle])
    headerRow1.push('')
    
    headerRow2.push('Turn In')
    headerRow2.push('Pass Through')
  })

  headerRow1.push('Total')
  headerRow1.push('Total Turn In')
  headerRow1.push('% Turn In')

  headerRow2.push('')
  headerRow2.push('')
  headerRow2.push('')

  const rows = [
    headerRow1,
    headerRow2
  ]

  // Grand totals
  const grandTotal = {}

  VEHICLE_ORDER.forEach(vehicle => {
    grandTotal[vehicle] = {
      TURN_IN: 0,
      PASS_THROUGH: 0
    }
  })

  let grandTotalVehicles = 0
  let grandTotalTurnIn = 0

  // Add hourly rows
  Object.entries(slots).forEach(([timeSlot, counts]) => {

    let rowTotal = 0
    let rowTurnIn = 0

    const row = [timeSlot]

    VEHICLE_ORDER.forEach(vehicle => {

      const turnIn = counts[vehicle].TURN_IN
      const passThrough = counts[vehicle].PASS_THROUGH

      row.push(turnIn)
      row.push(passThrough)

      rowTurnIn += turnIn
      rowTotal += turnIn + passThrough

      grandTotal[vehicle].TURN_IN += turnIn
      grandTotal[vehicle].PASS_THROUGH += passThrough
    })

    // Total
    row.push(rowTotal)

    // Total Turn In
    row.push(rowTurnIn)

    // Percentage
    const percentage =
      rowTotal > 0
        ? ((rowTurnIn / rowTotal) * 100).toFixed(2) + '%'
        : '0.00%'

    row.push(percentage)

    // Only show hours where something was counted
    if (rowTotal > 0) {
      rows.push(row)
    }

    grandTotalVehicles += rowTotal
    grandTotalTurnIn += rowTurnIn
  })

  // Grand Total row
  const grandRow = ['Grand Total']

  VEHICLE_ORDER.forEach(vehicle => {

    grandRow.push(grandTotal[vehicle].TURN_IN)
    grandRow.push(grandTotal[vehicle].PASS_THROUGH)

  })

  grandRow.push(grandTotalVehicles)
  grandRow.push(grandTotalTurnIn)

  const grandPercentage =
    grandTotalVehicles > 0
      ? ((grandTotalTurnIn / grandTotalVehicles) * 100).toFixed(2) + '%'
      : '0.00%'

  grandRow.push(grandPercentage)

  rows.push(grandRow)

  // Create worksheet
  const worksheet = XLSX.utils.aoa_to_sheet(rows)

  /*
    Merge vehicle headers

    2 Wheeler -> Turn In / Pass Through
    Car/SUV   -> Turn In / Pass Through
    etc.
  */

  worksheet['!merges'] = []

  let column = 1

  VEHICLE_ORDER.forEach(() => {

    worksheet['!merges'].push({
      s: { r: 0, c: column },
      e: { r: 0, c: column + 1 }
    })

    column += 2
  })

  // Merge Time Slot vertically
  worksheet['!merges'].push({
    s: { r: 0, c: 0 },
    e: { r: 1, c: 0 }
  })

  // Merge Total vertically
  worksheet['!merges'].push({
    s: { r: 0, c: column },
    e: { r: 1, c: column }
  })

  worksheet['!merges'].push({
    s: { r: 0, c: column + 1 },
    e: { r: 1, c: column + 1 }
  })

  worksheet['!merges'].push({
    s: { r: 0, c: column + 2 },
    e: { r: 1, c: column + 2 }
  })

  // Column widths
  worksheet['!cols'] = [
    { wch: 18 },

    ...VEHICLE_ORDER.flatMap(() => [
      { wch: 12 },
      { wch: 14 }
    ]),

    { wch: 12 },
    { wch: 16 },
    { wch: 12 }
  ]

  // Freeze header
  worksheet['!freeze'] = {
    xSplit: 1,
    ySplit: 2
  }

  // Create workbook
  const workbook = XLSX.utils.book_new()

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    `${shift} Summary`
  )

  // File name
  XLSX.writeFile(
    workbook,
    `BPCL-Vehicle-Count-${date}-${shift}.xlsx`
  )
}