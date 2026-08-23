import { useState, useEffect } from 'react'
import { supabase } from './supabaseClient'
import { exportDayToExcel } from './exportExcel'
import './Counter.css'

const VEHICLE_TYPES = [
  { key: '2W', label: '2 Wheeler', icon: '🏍️' },
  { key: 'CAR_SUV', label: 'Car / SUV', icon: '🚗' },
  { key: 'LCV', label: 'LCV', icon: '🚐' },
  { key: 'HCV', label: 'HCV', icon: '🚛' },
  { key: 'AGRI', label: 'Agri', icon: '🚜' },
  { key: 'BUS', label: 'Bus', icon: '🚌' },
  { key: 'OTHER', label: 'Other', icon: '🚘' },
]

const TRAFFIC_TYPES = {
  TURN_IN: 'TURN_IN',
  PASS_THROUGH: 'PASS_THROUGH',
}

function getDefaultShift() {
  const hour = new Date().getHours()

  // Morning: 9 AM - 9 PM
  if (hour >= 9 && hour < 21) {
    return 'Morning'
  }

  // Night: 9 PM - 9 AM
  return 'Night'
}

function getShiftDate(shift) {
  const now = new Date()

  // Night shift after midnight belongs to previous operational date
  if (shift === 'Night' && now.getHours() < 9) {
    now.setDate(now.getDate() - 1)
  }

  return now.toLocaleDateString('en-CA')
}

function Counter({ session }) {
  const [shift, setShift] = useState(getDefaultShift())

  const [todayCounts, setTodayCounts] = useState({
    TURN_IN: {},
    PASS_THROUGH: {},
  })

  const [loading, setLoading] = useState(false)

  const shiftDate = getShiftDate(shift)

  useEffect(() => {
    loadTodayCounts()
  }, [shift])

  // -----------------------------------------
  // LOAD COUNTS
  // -----------------------------------------
  const loadTodayCounts = async () => {
    const { data, error } = await supabase
      .from('vehicle_events')
      .select('vehicle_type, traffic_type')
      .eq('count_date', shiftDate)
      .eq('shift', shift)

    if (error) {
      console.error('Error loading counts:', error)
      return
    }

    const counts = {
      TURN_IN: {},
      PASS_THROUGH: {},
    }

    data?.forEach(row => {
      if (!counts[row.traffic_type]) {
        return
      }

      counts[row.traffic_type][row.vehicle_type] =
        (counts[row.traffic_type][row.vehicle_type] || 0) + 1
    })

    setTodayCounts(counts)
  }

  // -----------------------------------------
  // ADD VEHICLE
  // -----------------------------------------
  const handleIncrease = async (vehicleType, trafficType) => {
    setLoading(true)

    const { error } = await supabase
      .from('vehicle_events')
      .insert({
        vehicle_type: vehicleType,
        user_id: session.user.id,
        shift,
        count_date: shiftDate,
        traffic_type: trafficType,
      })

    if (error) {
      console.error('Error counting vehicle:', error)

      alert(
        'Could not save count: ' + error.message
      )

      setLoading(false)
      return
    }

    setTodayCounts(prev => ({
      ...prev,

      [trafficType]: {
        ...prev[trafficType],

        [vehicleType]:
          (prev[trafficType][vehicleType] || 0) + 1,
      },
    }))

    setLoading(false)
  }

  // -----------------------------------------
  // REMOVE VEHICLE
  // -----------------------------------------
  const handleDecrease = async (
    vehicleType,
    trafficType
  ) => {
    const currentCount =
      todayCounts[trafficType]?.[vehicleType] || 0

    if (currentCount <= 0) {
      return
    }

    /*
      Turn-In minus:
      removes the latest TURN_IN entry.

      Highway minus:
      removes the latest PASS_THROUGH entry.

      This keeps the database logically correct.
    */

    const { data, error } = await supabase
      .from('vehicle_events')
      .select('id')
      .eq('vehicle_type', vehicleType)
      .eq('shift', shift)
      .eq('count_date', shiftDate)
      .eq('traffic_type', trafficType)
      .order('created_at', {
        ascending: false,
      })
      .limit(1)
      .maybeSingle()

    if (error) {
      console.error(
        'Error finding vehicle:',
        error
      )

      alert('Could not decrease count.')
      return
    }

    if (!data) {
      alert(
        'There is no matching vehicle entry to decrease.'
      )
      return
    }

    const { error: deleteError } =
      await supabase
        .from('vehicle_events')
        .delete()
        .eq('id', data.id)

    if (deleteError) {
      console.error(
        'Error deleting vehicle:',
        deleteError
      )

      alert(
        'Could not decrease count: ' +
        deleteError.message
      )

      return
    }

    setTodayCounts(prev => ({
      ...prev,

      [trafficType]: {
        ...prev[trafficType],

        [vehicleType]: Math.max(
          0,
          (prev[trafficType][vehicleType] || 0) - 1
        ),
      },
    }))
  }

  // -----------------------------------------
  // TOTAL TURN-IN
  // -----------------------------------------
  const turnInTotal =
    VEHICLE_TYPES.reduce(
      (total, vehicle) =>
        total +
        (todayCounts.TURN_IN[vehicle.key] || 0),
      0
    )

  // -----------------------------------------
  // TOTAL PASS-THROUGH
  // -----------------------------------------
  const passThroughTotal =
    VEHICLE_TYPES.reduce(
      (total, vehicle) =>
        total +
        (todayCounts.PASS_THROUGH[vehicle.key] || 0),
      0
    )

  // -----------------------------------------
  // HIGHWAY TOTAL
  //
  // HIGHWAY = TURN-IN + PASS-THROUGH
  // -----------------------------------------
  const highwayTotal =
    turnInTotal + passThroughTotal

  // -----------------------------------------
  // TURN-IN PERCENTAGE
  // -----------------------------------------
  const turnInPercentage =
    highwayTotal > 0
      ? ((turnInTotal / highwayTotal) * 100).toFixed(2)
      : '0.00'

  return (
    <div className="counter-page">

      {/* HEADER */}
      <div className="counter-header">

        <div>
          <div className="app-title">
            BPCL Vehicle Counter
          </div>

          <div className="app-subtitle">
            Vehicle movement recording system
          </div>
        </div>

        <div className="shift-badge">
          <span className="shift-dot"></span>
          {shift} Shift
        </div>

      </div>


      {/* SHIFT + TOTAL */}
      <div className="top-section">

        <div className="shift-box">

          <label>
            Current Shift
          </label>

          <select
            value={shift}
            onChange={e =>
              setShift(e.target.value)
            }
          >
            <option value="Morning">
              Morning
            </option>

            <option value="Night">
              Night
            </option>
          </select>

          <div className="shift-time">
            {shift === 'Morning'
              ? '09:00 AM — 09:00 PM'
              : '09:00 PM — 09:00 AM'}
          </div>

        </div>


        {/* HIGHWAY TOTAL */}
        <div className="total-box">

          <div className="total-label">
            Highway Total
          </div>

          <div className="total-number">
            {highwayTotal}
          </div>

          <div className="total-date">
            {shiftDate}
          </div>

        </div>

      </div>


      {/* TRAFFIC SUMMARY */}
      <div className="traffic-summary">

        {/* TURN-IN */}
        <div className="traffic-summary-box turn-in-summary">

          <div className="traffic-summary-title">
            ↪ Turn-In
          </div>

          <div className="traffic-summary-number">
            {turnInTotal}
          </div>

        </div>


        {/* HIGHWAY */}
        <div className="traffic-summary-box highway-summary">

          <div className="traffic-summary-title">
            → Highway
          </div>

          <div className="traffic-summary-number">
            {highwayTotal}
          </div>

        </div>


        {/* PERCENTAGE */}
        <div className="traffic-summary-box percentage-summary">

          <div className="traffic-summary-title">
            Turn-In %
          </div>

          <div className="traffic-summary-number">
            {turnInPercentage}%
          </div>

        </div>

      </div>


      {/* VEHICLE COUNT */}
      <div className="section-title">
        Vehicle Count
      </div>


      {/* COLUMN HEADINGS */}
      <div className="traffic-column-headings">

        <div>
          Vehicle
        </div>

        <div>
          ↪ Turn-In
        </div>

        <div>
          → Highway
        </div>

      </div>


      {/* VEHICLE COUNTERS */}
      <div className="vehicle-grid">

        {VEHICLE_TYPES.map(vehicle => {

          const turnInCount =
            todayCounts.TURN_IN[
              vehicle.key
            ] || 0

          const passThroughCount =
            todayCounts.PASS_THROUGH[
              vehicle.key
            ] || 0

          /*
            IMPORTANT:

            Highway count is NOT stored separately.

            Highway =
            Turn-In + Pass-Through
          */
          const highwayCount =
            turnInCount + passThroughCount

          return (
            <div
              className="vehicle-card"
              key={vehicle.key}
            >

              {/* VEHICLE */}
              <div className="vehicle-info">

                <div className="vehicle-icon">
                  {vehicle.icon}
                </div>

                <div className="vehicle-name">
                  {vehicle.label}
                </div>

              </div>


              {/* TURN-IN */}
              <div className="traffic-counter">

                <div className="vehicle-count">
                  {turnInCount}
                </div>

                <div className="counter-controls">

                  <button
                    className="minus-button"
                    onClick={() =>
                      handleDecrease(
                        vehicle.key,
                        TRAFFIC_TYPES.TURN_IN
                      )
                    }
                    disabled={
                      turnInCount === 0 ||
                      loading
                    }
                  >
                    −
                  </button>

                  <button
                    className="plus-button"
                    onClick={() =>
                      handleIncrease(
                        vehicle.key,
                        TRAFFIC_TYPES.TURN_IN
                      )
                    }
                    disabled={loading}
                  >
                    +
                  </button>

                </div>

              </div>


              {/* HIGHWAY */}
              <div className="traffic-counter">

                <div className="vehicle-count">
                  {highwayCount}
                </div>

                <div className="counter-controls">

                  {/*
                    Highway + means:
                    PASS-THROUGH vehicle.

                    If Turn-In is pressed,
                    Highway automatically increases too.
                  */}

                  <button
                    className="minus-button"
                    onClick={() =>
                      handleDecrease(
                        vehicle.key,
                        TRAFFIC_TYPES.PASS_THROUGH
                      )
                    }
                    disabled={
                      passThroughCount === 0 ||
                      loading
                    }
                  >
                    −
                  </button>

                  <button
                    className="plus-button"
                    onClick={() =>
                      handleIncrease(
                        vehicle.key,
                        TRAFFIC_TYPES.PASS_THROUGH
                      )
                    }
                    disabled={loading}
                  >
                    +
                  </button>

                </div>

              </div>

            </div>
          )
        })}

      </div>


      {/* EXCEL */}
      <button
        className="export-button"
        onClick={() =>
          exportDayToExcel(
            shiftDate,
            shift
          )
        }
      >
        <span>📊</span>
        Export {shift} Shift Report
      </button>


      <div className="footer-note">
        Turn-In vehicles are included in Highway Total.
      </div>

    </div>
  )
}

export default Counter